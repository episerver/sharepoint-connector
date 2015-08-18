using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers.Constructions;
using EPiServer.Construction;
using EPiServer.Core;
using EPiServer.Core.Transfer;
using EPiServer.DataAbstraction;
using EPiServer.Logging;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Validation;
using EPiServer.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.Providers
{
    /// <summary>
    /// This ContentProvider provides SharePoint's media to EPiServer's AssetTree
    /// </summary>
    public class SharePointMediaProvider : ContentProvider
    {
        private static readonly ILogger _logger = LogManager.GetLogger();

        private readonly IContentEvents _contentEvents;
        private readonly ContentMediaResolver _contentMediaResolver;
        private readonly IContentRepository _contentRepository;
        private IEnumerable<ISharePointDataBuilder> _contentBuilders;

        private ISharePointContentCacheProvider _contentCache;
        private ISharePointContentMappingRepository _sharePointRepository;
        private Func<ContentReference, object> _getCMSCachedContent;
        private Func<ContentReference> _getEntryPoint;

        private readonly string _providerKey;

        /// <summary>
        /// Gets the provider key that this ContentProvider instance has. It is unique for the EPiServer instance.
        /// This is the value that is set on <see cref="ContentReference.ProviderName "/> on each ContentReference served by this provider instance.
        /// </summary>
        /// <value>The provider key.</value>
        public override string ProviderKey
        {
            get
            {
                return _providerKey;
            }
        }

        public ISharePointContentCacheProvider ContentCache
        {
            get
            {
                return _contentCache;
            }
        }

        #region Ctors

        public SharePointMediaProvider(string providerKey)
            : this(providerKey,
                ServiceLocator.Current.GetInstance<ContentFactory>(),
                ServiceLocator.Current.GetInstance<ServiceAccessor<IContentTypeRepository>>(),
                ServiceLocator.Current.GetInstance<TemplateResolver>(),
                ServiceLocator.Current.GetInstance<UrlSegmentHandler>(),
                ServiceLocator.Current.GetInstance<ILanguageBranchRepository>(),
                ServiceLocator.Current.GetInstance<IFolderIDGenerator>(),
                ServiceLocator.Current.GetInstance<IContentCopyHandler>(),
                ServiceLocator.Current.GetInstance<IValidationService>(),
                ServiceLocator.Current.GetInstance<IContentEvents>(),
                ServiceLocator.Current.GetInstance<ContentMediaResolver>(),
                ServiceLocator.Current.GetInstance<IContentRepository>(),
                ServiceLocator.Current.GetAllInstances<ISharePointDataBuilder>(),
                ServiceLocator.Current.GetInstance<ISharePointContentMappingRepository>(),
                (contentLink) => { return CacheManager.Get(DataFactoryCache.PageMasterLanguageCacheKey(contentLink)); },
                () => { return SharePointSettings.Current.EntryPoint; },
                ServiceLocator.Current.GetInstance<ISharePointContentCacheProvider>()) { }

        public SharePointMediaProvider(
            string providerKey,
            ContentFactory contentFactory,
            ServiceAccessor<IContentTypeRepository> serviceAccessor,
            TemplateResolver templateResolver,
            UrlSegmentHandler urlSegmentHandler,
            ILanguageBranchRepository languageBranchRepository,
            IFolderIDGenerator folderIDGenerator,
            IContentCopyHandler contentCopyHandler,
            IValidationService validationService,
            IContentEvents contentEvents,
            ContentMediaResolver contentMediaResolver,
            IContentRepository contentRepository,
            IEnumerable<ISharePointDataBuilder> sharePointDataBuilders,
            ISharePointContentMappingRepository sharePointRepository,
            Func<ContentReference, object> getCMSCachedContent,
            Func<ContentReference> getEntryPoint,
            ISharePointContentCacheProvider contentCache)
            : base(contentFactory,
                serviceAccessor,
                templateResolver,
                urlSegmentHandler,
                languageBranchRepository,
                folderIDGenerator,
                contentCopyHandler,
                validationService)
        {
            _providerKey = providerKey;
            _contentEvents = contentEvents;
            _contentCache = contentCache;
            _contentMediaResolver = contentMediaResolver;
            _contentRepository = contentRepository;
            _contentBuilders = sharePointDataBuilders;
            _sharePointRepository = sharePointRepository;
            _getCMSCachedContent = getCMSCachedContent;
            _getEntryPoint = getEntryPoint;

            if (_contentEvents != null)
            {
                _contentEvents.CreatedContent += ContentEvents_CreatedContent;
            }
        }

        #endregion

        #region Public

        /// <summary>
        /// Lookups the list in the cache
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns>The content link</returns>
        public ContentReference LookupList(string key)
        {
            var helper = new SharePointHelper();
            var idListAndView = helper.GetListAndViewIds(key);
            if (idListAndView == null)
            {
                return ContentReference.EmptyReference;
            }

            List<IModelWrapper> list;
            _contentCache.Folders.TryGetValue(_getEntryPoint().ToString(), out list);

            if (list != null)
            {
                foreach (var item in list)
                {
                    if (idListAndView[1].Equals(item.View.Id.ToString()))
                    {
                        return item.ContentLink;
                    }
                }
            }

            return ContentReference.EmptyReference;
        }

        /// <summary>
        /// Gets the content and the model wrapper for specified content link.
        /// </summary>
        /// <param name="contentLink">The content link.</param>
        /// <param name="modelWrapper">The model wrapper.</param>
        public IContent LoadContent(ContentReference contentLink, out IModelWrapper modelWrapper)
        {
            return LoadOrBuildContent(contentLink, out modelWrapper);
        }

        /// <summary>
        /// Gets the model wrapper for a content link.
        /// </summary>
        /// <param name="contentLink">The content link.</param>
        public IModelWrapper GetModelWrapper(ContentReference contentLink)
        {
            ContentReference parentLink;
            return GetModelWrapper(contentLink, out parentLink);
        }

        #endregion

        #region Override

        public override ContentReference Save(IContent content, DataAccess.SaveAction action)
        {
            _logger.Information("SharePointMediaProvider Save - contentLink: {0}", content.ContentLink);
            return LoadOrBuildContent(content.ContentLink, null).ContentLink;
        }

        public override IContent GetDefaultContent(IContent parentContent, int contentTypeID, ILanguageSelector languageSelector)
        {
            _logger.Information("SharePointMediaProvider GetDefaultContent - parentContent: {0} - contentTypeID: {1}", parentContent.ContentLink, contentTypeID);
            var content = base.ContentFactory.CreateContent(base.ContentTypeRepository.Load(contentTypeID));
            content.ContentTypeID = contentTypeID;

            return content;
        }

        protected override IContent LoadContent(ContentReference contentLink, ILanguageSelector languageSelector)
        {
            _logger.Information("SharePointMediaProvider LoadContent - contentLink: {0}", contentLink);
            return LoadOrBuildContent(contentLink, languageSelector);
        }

        /// <summary>
        /// Loads existing content or build a new one from the given content link and language selector objects
        /// </summary>
        /// <param name="contentLink">The given content link object</param>
        /// <param name="languageSelector">The given language selector object</param>
        /// <returns></returns>
        private IContent LoadOrBuildContent(ContentReference contentLink, ILanguageSelector languageSelector)
        {
            IModelWrapper modelWrapper;
            return LoadOrBuildContent(contentLink, out modelWrapper);
        }

        /// <summary>
        /// Loads existing content or build a new one from the given content link and language selector objects
        /// </summary>
        /// <param name="contentLink">The given content link object</param>
        /// <param name="languageSelector">The given language selector object</param>
        /// <returns></returns>
        private IContent LoadOrBuildContent(ContentReference contentLink, out IModelWrapper modelWrapper)
        {
            modelWrapper = null;
            var mappedItem = _sharePointRepository.GetIdentity(contentLink);
            if (mappedItem == null)
            {
                _logger.Information("Not found MappedIdentity for contentLink: {0}. Return null.", contentLink);
                return null;
            }

            // For sharepoint media we first see if we have master language in cache
            IContent content = _getCMSCachedContent(contentLink) as IContent;
            if (content != null)
            {
                _logger.Information("Found content for contentLink: {0} in CMS cache, return.", contentLink);
                return content;
            }
            _contentCache.BuildMapping(contentLink);
            
            ContentReference parentContentLink;
            modelWrapper = GetModelWrapper(contentLink, out parentContentLink);

            // Return null if listItem has deleted or unmap, 
            if (modelWrapper == null)
            {
                _logger.Information("Cannot get the modelWrapper for contentLink: {0}, return null", contentLink);
                return null;
            }

            // Type of listItem: FileTypes of FolderType
            var sharePointType = modelWrapper.IsFile ? GetFileType(modelWrapper.ListItem.File) : typeof(ContentFolder);

            var builders = _contentBuilders.Where(b => b.HandleType(sharePointType));
            if (builders == null || builders.Count() <= 0)
            {
                var errorString = string.Format("ContentLink '{0}' maps to not supported sharepoint content type.", contentLink);
                _logger.Error(errorString);
                throw new NotSupportedException(errorString);
            }

            // Create new content with share point content type
            var requestedType = ContentTypeRepository.Load(sharePointType);
            content = _contentRepository.GetDefault<IContent>(parentContentLink, requestedType.ID);
            content.ParentLink = parentContentLink;
            content.ContentGuid = mappedItem.ContentGuid;
            content.ContentLink = mappedItem.ContentLink;

            // Assign read only for this content
            var iContentSecurable = content as IContentSecurable;
            if (iContentSecurable != null)
            {
                var securableDescriptor = iContentSecurable.GetContentSecurityDescriptor();
                if (securableDescriptor != null)
                {
                    securableDescriptor.IsInherited = true;
                    securableDescriptor.ContentLink = content.ContentLink;
                }

            }

            _logger.Information("Build content for contentLink: {0}", contentLink);
            foreach (var builder in builders)
            {
                content = builder.Build(content, contentLink, parentContentLink, modelWrapper);
            }

            return content;
        }

        /// <summary>
        /// Gets the model wrapper and parent content link for a content link.
        /// </summary>
        /// <param name="contentLink">The content link.</param>
        /// <param name="parentContentLink">The parent content link.</param>
        private IModelWrapper GetModelWrapper(ContentReference contentLink, out ContentReference parentContentLink)
        {
            IModelWrapper modelWrapper = null;
            parentContentLink = ContentReference.EmptyReference;

            // Find the oppropriate file and its parent
            foreach (var item in _contentCache.Files)
            {
                var file = item.Value.Where(v => v.ContentLink.Equals(contentLink)).FirstOrDefault();
                if (file != null)
                {
                    parentContentLink = new ContentReference(item.Key);
                    modelWrapper = file;
                    _logger.Information("The content link {0} is file. The mapping url {1}", contentLink, file.ToMappingUrl());
                    break;
                }
            }

            // Find the oppropriate folder/subfolder and its parent
            if (modelWrapper == null)
            {
                foreach (var item in _contentCache.Folders)
                {
                    var folder = item.Value.Where(v => v.ContentLink.Equals(contentLink)).FirstOrDefault();
                    if (folder != null)
                    {
                        parentContentLink = new ContentReference(item.Key);
                        modelWrapper = folder;
                        _logger.Information("The content link {0} is folder. The mapping url {1}", contentLink, folder.ToMappingUrl());
                        break;
                    }
                }
            }
            return modelWrapper;
        }

        protected override IList<GetChildrenReferenceResult> LoadChildrenReferencesAndTypes(ContentReference contentLink, string languageID, out bool languageSpecific)
        {
            languageSpecific = false; // no implementation of this yet

            var children = new List<GetChildrenReferenceResult>();
            IContent parent;
            try
            {
                parent = _contentRepository.Get<IContent>(contentLink);
            }
            catch (ContentNotFoundException)
            {
                _logger.Error("Content not found for contentLink: {0}. Return no children.", contentLink);
                return children;
            }
            _contentCache.BuildMapping(contentLink);
            var key = contentLink.ToString();
            var entryPoint = _getEntryPoint();
            if (contentLink.CompareToIgnoreWorkID(entryPoint))
            {
                key = entryPoint.ToString();
            }

            // Get files child of this content link
            var item = _contentCache.Files.Where(f => f.Key.Equals(key)).FirstOrDefault();
            if (item.Value != null)
            {
                foreach (var child in item.Value)
                {
                    children.Add(new GetChildrenReferenceResult()
                        {
                            ContentLink = child.ContentLink,
                            IsLeafNode = true,
                            ModelType = GetFileType(child.ListItem.File)
                        });
                }
            }

            // Get folder child of this content link
            item = _contentCache.Folders.Where(f => f.Key.Equals(key)).FirstOrDefault();
            if (item.Value != null)
            {
                foreach (var child in item.Value)
                {
                    children.Add(new GetChildrenReferenceResult()
                    {
                        ContentLink = child.ContentLink,
                        IsLeafNode = false,
                        ModelType = typeof(ContentFolder)
                    });
                }
            }

            _logger.Information("Return children for contentLink: {0} - childrenCount: {1}", contentLink, children.Count);
            return children;
        }

        #endregion

        #region Private

        /// <summary>
        ///     Match model types of template depend on extension of file
        /// </summary>
        /// <param name="file">The given file</param>
        /// <returns>the type of file. eg: Image, Video, Media</returns>
        private Type GetFileType(SPFile file)
        {
            if (file == null)
            {
                return null;
            }
            var fileInfo = new FileInfo(file.Name);
            return ContentTypeRepository.Load(_contentMediaResolver.GetFirstMatching(String.IsNullOrWhiteSpace(fileInfo.Extension) ? "." : fileInfo.Extension)).ModelType;
        }

        /// <summary>
        /// Clear source content link cache incase copy/paste event
        /// </summary>
        /// <param name="sender">The given sender object</param>
        /// <param name="e">The given content event argurment object</param>
        private void ContentEvents_CreatedContent(object sender, ContentEventArgs e)
        {
            // Verifies the given event object is CopyContentEventArgs or not
            var copyContentEventArgs = e as CopyContentEventArgs;
            if (copyContentEventArgs == null)
            {
                return;
            }

            // Verifies the given event object have source ContentLink or not
            var sourceContentLink = copyContentEventArgs.SourceContentLink;
            if (sourceContentLink == null || !sourceContentLink.IsExternalProvider || !sourceContentLink.ProviderName.Equals(ProviderKey))
            {
                return;
            }

            CacheManager.RemoveLocalOnly(DataFactoryCache.PageCommonCacheKey(sourceContentLink));
        }

        #endregion
    }
}