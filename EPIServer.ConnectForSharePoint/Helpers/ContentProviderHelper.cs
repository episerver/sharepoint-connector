using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Logging;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using System;
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
        /// Given a sharepoint object quid we will try to find it in the specified folder
        /// </summary>
        /// <param name="SPItemGuid"></param>
        /// <param name="folderRef"></param>
        /// <returns></returns>
        public ContentReference FindMediaInFolder(Guid SPItemGuid, ContentReference folderRef)
        {
            var medialist = _contentRepository.GetChildren<IContent>(folderRef);
            foreach (var item in medialist)
            {
                if (item is ISharePointItem)
                {
                    var mediaEX = item as ISharePointItem;
                    if (SPItemGuid.ToString().Equals(mediaEX.SharePointGUID))
                        return item.ContentLink;
                }
            }
            return null;
        }
    }
}