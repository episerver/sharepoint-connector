using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.DataAccess;
using EPiServer.Logging;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using System;
using System.Collections.Specialized;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// The helper class for content provider
    /// At the moment, it only support SharePointMediaProvider class
    /// </summary>
    public class ContentProviderHelper : IContentProviderHelper
    {
        private IServiceLocator _serviceLocator;
        private static readonly ILogger _logger = LogManager.GetLogger();

        private static IContentRepository _contentRepository;
        private static IContentProviderManager _contentProviderManager;
        private static IContentSecurityRepository _contentSecurityRepository;

        private static object _lockObject = new Object();
        private static SharePointMediaProvider _sharePointMediaProvider = null;

        private static readonly AccessLevel _noAccessLevel = AccessLevel.NoAccess;
        private static readonly AccessLevel _readAccessLevel = AccessLevel.Read;

        public const string ContentProviderKey = "sharepointcontent";

        public ContentProviderHelper()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
            _contentProviderManager = ServiceLocator.Current.GetInstance<IContentProviderManager>();
            _contentSecurityRepository = ServiceLocator.Current.GetInstance<IContentSecurityRepository>();
        }

        internal ContentProviderHelper(IServiceLocator serviceLocator)
        {
            _serviceLocator = serviceLocator;
            _contentRepository = serviceLocator.GetInstance<IContentRepository>();
            _contentProviderManager = serviceLocator.GetInstance<IContentProviderManager>();
            _contentSecurityRepository = serviceLocator.GetInstance<IContentSecurityRepository>();
        }
        /// <summary>
        /// Return entry point from settings
        /// </summary>
        private static ContentReference EntryPoint
        {
            get { return SharePointSettings.Current.EntryPoint; }
        }

        /// <summary>
        /// Gets or create a new instance of the SharePointMediaProvider class
        /// </summary>
        public static SharePointMediaProvider SharePointMediaProvider
        {
            get
            {
                if (_sharePointMediaProvider == null)
                {
                    lock (_lockObject)
                    {
                        if (_sharePointMediaProvider == null)
                        {
                            try
                            {
                                _sharePointMediaProvider = (_contentProviderManager.GetProvider(ContentProviderKey) as SharePointMediaProvider)
                                    ?? new SharePointMediaProvider(ContentProviderKey);
                            }
                            catch (ContentProviderNotFoundException exc)
                            {
                                _logger.Debug(String.Format("{0} content provider was not found", typeof(SharePointMediaProvider)), exc);
                                _sharePointMediaProvider = new SharePointMediaProvider(ContentProviderKey);
                            }
                            catch (Exception ex)
                            {
                                _logger.Debug(String.Format("{0} content provider was not found", typeof(SharePointMediaProvider)), ex);
                            }
                        }
                    }
                }

                return _sharePointMediaProvider;
            }
        }

        /// <summary>
        /// Map configured SharePoint's media to the EPiServer Asset tree
        /// </summary>
        public static void CreateSharePointMediaProvider()
        {
            // Do nothing if entry point is empty
            if (ContentReference.IsNullOrEmpty(EntryPoint))
            {
                _logger.Debug("The EntryPoint is null");
                return;
            }

            // Create SharePoint Media Provider and add to ProviderMap
            InitializeSharePointMediaProvider();

            var supportedLists = SharePointHelper.SPServicesHub.SPListsViewsService
                .GetLists((int)ListBaseType.DocumentLibrary);

            var helper = new SharePointHelper();
            var selectedLists = helper.GetSelectedListViewIds(supportedLists);

            foreach (var list in selectedLists)
            {
                var cachedContentLink = SharePointMediaProvider.LookupList(list.Key);
                // If we can not lookup this list's content link in cache then ignore it.
                // Its maybe sharepoint return not found on requested file id. (timeout for instance)
                if (ContentReference.IsNullOrEmpty(cachedContentLink))
                {
                    continue;
                }

                // Create mapping folders for SharePoint lists under the entry point
                // Create a folder for the list under SharePoint entry point
                var sharePointMappingFolder = _contentRepository.GetDefault<ContentFolder>(EntryPoint);
                sharePointMappingFolder.Name = list.Value;
                sharePointMappingFolder.ContentLink = cachedContentLink;
                sharePointMappingFolder.ParentLink = EntryPoint;

                _contentRepository.Save(sharePointMappingFolder, SaveAction.Publish, _noAccessLevel);
            }

            // Set access level to 'Read' on the SharePoint's root folder and its children
// hem_remove - removed the event listener that sets all the nodes to read only for now
//            AccessRightsHelper.SetAccessRight(EntryPoint, AccessLevel.Read);
//           _contentSecurityRepository.ContentSecuritySaved -= ContentSecurityRepository_ContentSecuritySaved;
//           _contentSecurityRepository.ContentSecuritySaved += ContentSecurityRepository_ContentSecuritySaved;
        }

        /// <summary>
        /// Delete the SharePoint Media Provider
        /// </summary>
        public static void DeleteSharePointMediaProvider()
        {
            // Remove the provider
            if (SharePointMediaProvider != null)
            {
                _contentProviderManager.ProviderMap.RemoveProvider(SharePointMediaProvider.Name);
            }

            _contentSecurityRepository.ContentSecuritySaved -= ContentSecurityRepository_ContentSecuritySaved;
        }

        /// <summary>
        /// Create SharePoint Media Provider for the given entry point
        /// </summary>
        /// <param name="entryPointContenLink">The given content reference of the entry point</param>
        private static void InitializeSharePointMediaProvider()
        {
            _contentProviderManager.ProviderMap.RemoveProvider(ContentProviderKey);

            // Init the provider and its settings
            var config = new NameValueCollection();
            config.Add("entryPoint", EntryPoint.ID.ToString());
            config.Add("name", ContentProviderKey);
            config.Add("type", typeof(SharePointMediaProvider).AssemblyQualifiedName);
            config.Add("capabilities", string.Join(",", ContentProviderCapabilities.Create, ContentProviderCapabilities.Edit, ContentProviderCapabilities.Copy));

            _sharePointMediaProvider = new SharePointMediaProvider(ContentProviderKey);
            _sharePointMediaProvider.Initialize(ContentProviderKey, config);
            _contentProviderManager.ProviderMap.AddProvider(_sharePointMediaProvider);
        }

        /// <summary>
        /// Sets access right each time the given content changed its security level
        /// </summary>
        /// <param name="sender">The given sender object</param>
        /// <param name="e">The given event</param>
        private static void ContentSecurityRepository_ContentSecuritySaved(object sender, ContentSecurityEventArg e)
        {
            ResetAccessRight(e.ContentLink);
        }

        /// <summary>
        /// Reset access right on the given content reference object and its descendents
        /// </summary>
        /// <param name="contentLink">The given content reference object</param>
        private static void ResetAccessRight(ContentReference contentLink)
        {
            var sharePointRoot = GetSharePointRoot(contentLink);
            if (sharePointRoot == null)
            {
                _logger.Debug("The SharePoint root not found");
                return;
            }

            var existingEntries = _contentSecurityRepository.Get(sharePointRoot).Entries;
            if (existingEntries == null || existingEntries.Count() == 0)
            {
                _logger.Debug("The entry not existed");
                return;
            }

            if (existingEntries.Count() == 1 && existingEntries.First().Access.Equals(_readAccessLevel))
            {
                _logger.Debug("The existing entry is read only");
                return;
            }
            var accessHelper = new AccessRightsHelper();
            accessHelper.SetAccessRight(sharePointRoot, _readAccessLevel);
        }

        /// <summary>
        /// Gets SharePoint root object from the given content reference object
        /// </summary>
        /// <param name="contentLink">The given content reference object</param>
        /// <returns>The SharePoint root content reference object</returns>
        private static ContentReference GetSharePointRoot(ContentReference contentLink)
        {
            if (contentLink == null)
            {
                _logger.Debug("GetSharePointRoot: The content link is null");
                return null;
            }

            var content = _contentRepository.Get<IContent>(contentLink);
            if (content != null)
            {
                // CASE: directly set access right on the SharePoint root asset
                if (ContentReference.Equals(content.ContentLink, EntryPoint))
                {
                    return content.ContentLink;
                }

                // CASE: set access right on one of the SharePoint root assest ancestors
                // ex.:
                //      Set access right on the "Global Assets" or "Root"
                var descendents = _contentRepository.GetDescendents(contentLink);
                if (descendents == null || descendents.Count() == 0)
                {
                    _logger.Debug("The are not any descendent of content link", contentLink);
                    return null;
                }

                foreach (var item in descendents)
                {
                    content = _contentRepository.Get<IContent>(item);
                    if (content != null)
                    {
                        if (ContentReference.Equals(content.ContentLink, EntryPoint))
                        {
                            return content.ContentLink;
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Recursively searches for specific ContentFolder given a name and a startRef (if null uses SiteDefinition.Current.GlobalAssetsRoot)
        /// </summary>
        /// <param name="name"></param>
        /// <param name="startRef">if null, it will use SiteDefinition.Current.GlobalAssetsRoot as starting ref</param>
        /// <returns></returns>
        public ContentReference FindRecursiveFolders(string name, ContentReference startRef = null)
        {
            if( startRef == null )
                startRef = SiteDefinition.Current.GlobalAssetsRoot;

            var rootFolders = _contentRepository.GetChildren<ContentFolder>(startRef);
            foreach (var rootFolder in rootFolders)
            {
                if (rootFolder.Name == name)
                    return rootFolder.ContentLink;
                else
                {
                    ContentReference cref = FindRecursiveFolders(name, rootFolder.ContentLink);
                    if (cref != null)
                        return cref;
                }
            }

            return null;
        }

        /// <summary>
        /// checks the current folder for the media item.
        /// </summary>
        /// <param name="name"></param>
        /// <param name="startRef"></param>
        /// <returns></returns>
        public ContentReference FindMediaInFolder(string name, ContentReference folderRef)
        {
            // Check this folder
            var medialist = _contentRepository.GetChildren<MediaData>(folderRef);
            foreach (var item in medialist)
            {
                if (item.Name == name)
                    return item.ContentLink;
            }
            return null;
        }

        /// <summary>
        /// Set the given access rights for the given content reference
        /// </summary>
        /// <param name="contentReference">The given content reference</param>
        /// <param name="accessLevel">The given access rights</param>
        public void SetAccessRight(ContentReference contentReference, AccessLevel accessLevel)
        {
            var contentSecurity = _contentSecurityRepository.Get(contentReference).CreateWritableClone() as IContentSecurityDescriptor;
            if (contentSecurity == null)
            {
                return;
            }

            // The Access Control List does not support adding or modifying entries when it is set to inherited
            // So that, convert it to local
            if (contentSecurity.IsInherited)
            {
                contentSecurity.ToLocal();
            }

            contentSecurity.Clear();
            // Add access right for the default groups
            contentSecurity.AddEntry(new AccessControlEntry("Everyone", accessLevel));
            _contentSecurityRepository.Save(contentReference, contentSecurity, SecuritySaveType.Replace);
        }
    }
}