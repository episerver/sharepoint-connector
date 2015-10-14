using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Construction;
using EPiServer.Core;
using EPiServer.Core.Transfer;
using EPiServer.DataAbstraction;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using EPiServer.Validation;
using EPiServer.Web;
using System;
using System.Collections.Generic;
using System.IO;

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

        #region Ctors

        public SharePointMediaProvider(string providerKey)
            : this(providerKey,
                ServiceLocator.Current.GetInstance<ContentFactory>(),
                ServiceLocator.Current.GetInstance<ServiceAccessor<IContentTypeRepository>>(),
                ServiceLocator.Current.GetInstance<TemplateResolver>(),
                ServiceLocator.Current.GetInstance<UrlSegmentHandler>(),
                ServiceLocator.Current.GetInstance<ILanguageBranchRepository>(),
                ServiceLocator.Current.GetInstance<IContentCopyHandler>(),
                ServiceLocator.Current.GetInstance<IValidationService>(),
                ServiceLocator.Current.GetInstance<IContentEvents>(),
                ServiceLocator.Current.GetInstance<ContentMediaResolver>(),
                ServiceLocator.Current.GetInstance<IContentRepository>(),
                ServiceLocator.Current.GetInstance<ISharePointContentMappingRepository>(),
                (contentLink) => { return CacheManager.Get(DataFactoryCache.PageMasterLanguageCacheKey(contentLink)); },
                () => { return SharePointSettings.Current.EntryPoint; }) { }

        public SharePointMediaProvider(
            string providerKey,
            ContentFactory contentFactory,
            ServiceAccessor<IContentTypeRepository> serviceAccessor,
            TemplateResolver templateResolver,
            UrlSegmentHandler urlSegmentHandler,
            ILanguageBranchRepository languageBranchRepository,
            IContentCopyHandler contentCopyHandler,
            IValidationService validationService,
            IContentEvents contentEvents,
            ContentMediaResolver contentMediaResolver,
            IContentRepository contentRepository,
            ISharePointContentMappingRepository sharePointRepository,
            Func<ContentReference, object> getCMSCachedContent,
            Func<ContentReference> getEntryPoint)
            : base(contentFactory,
                serviceAccessor,
                templateResolver,
                urlSegmentHandler,
                languageBranchRepository,
                contentCopyHandler,
                validationService)
        {
            _providerKey = providerKey;
            _contentEvents = contentEvents;
            _contentMediaResolver = contentMediaResolver;
            _contentRepository = contentRepository;
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

            return ContentReference.EmptyReference;
        }

         #endregion

        #region Override

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

        protected override IContent LoadContent(ContentReference contentLink, ILanguageSelector languageSelector)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}