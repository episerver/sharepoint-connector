using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using System;
using System.Diagnostics.CodeAnalysis;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// The helper class for content provider
    /// At the moment, it only support SharePointMediaProvider class
    /// </summary>
    public class ContentProviderHelper : IContentProviderHelper
    {
        private IServiceLocator _serviceLocator;

        private static IContentRepository _contentRepository;

        [ExcludeFromCodeCoverageAttribute]
        public ContentProviderHelper()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
        }

        internal ContentProviderHelper(IServiceLocator serviceLocator)
        {
            _serviceLocator = serviceLocator;
            _contentRepository = serviceLocator.GetInstance<IContentRepository>();
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

            var folders = _contentRepository.GetChildren<ContentFolder>(startRef);
            foreach (var folder in folders)
            {
                if (folder.Name == name)
                    return folder.ContentLink;
                else
                {
                    ContentReference cref = FindRecursiveFolders(name, folder.ContentLink);
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