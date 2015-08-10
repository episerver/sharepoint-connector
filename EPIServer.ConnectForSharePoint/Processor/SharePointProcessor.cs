using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Core;
using EPiServer.DataAccess;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessor : ISharePointProcessor
    {
        // The content repository we use to create folders and objects in episerver
        IContentRepository _contentRepository;

        public SharePointProcessor()
        {
            _contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
        }

        internal SharePointProcessor( IContentRepository repository)
        {
            _contentRepository = repository;
        }

        public void AddDocument(SPListItem item)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Adds a new content folder into CMS 
        /// </summary>
        /// <param name="item"></param>
        public void AddFolder(SPListItem item)
        {
            // Set access level to 'Read' on the SharePoint's root folder and its children
            ContentReference cref = new ContentReference(item.entryPointId);
            AccessRightsHelper.SetAccessRight(cref, AccessLevel.FullAccess);

            // Create a folder for the list under SharePoint entry point
            var cmsFolder = _contentRepository.GetDefault<ContentFolder>(cref);
            cmsFolder.Name = item.Folder.Name;
            cmsFolder.ParentLink = SharePointSettings.Current.EntryPoint;
        
            // Save the folder
            _contentRepository.Save(cmsFolder, SaveAction.Publish, AccessLevel.FullAccess);
        }

        public void DeleteFolder(SPListItem item)
        {
            // Set access level to 'Read' on the SharePoint's root folder and its children
            ContentReference cref = new ContentReference(item.entryPointId);
            AccessRightsHelper.SetAccessRight(cref, AccessLevel.FullAccess);

            ContentReference rootRef = new ContentReference(item.entryPointId);
            var cmsFolder = _contentRepository.GetDefault<ContentFolder>(rootRef);
            var children = _contentRepository.GetChildren<ContentFolder>(rootRef);
            foreach (var child in children)
            {
                if( child.Name.Equals(item.Folder.Name) )
                {
                    _contentRepository.Delete(child.ContentLink,true,AccessLevel.Delete);
                    break;
                }
            }
        }

        public void UpdateDocument(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void UpdateFolder(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void DeleteDocument(SPListItem item)
        {
            throw new NotImplementedException();
        }
    }
}
