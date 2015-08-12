using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.DataAccess;
using EPiServer.Framework.Blobs;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using EPiServer.Web.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Hosting;

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

        public ContentReference FindMediaFolderByID(ContentReference parentFolder, ContentReference defaultFolder = null)
        {
            var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
            var rootFolders = contentRepository.GetChildren<ContentFolder>(SiteDefinition.Current.GlobalAssetsRoot);

            foreach (var rootFolder in rootFolders)
            {
                if (rootFolder.ContentLink.ID == parentFolder.ID) return rootFolder.ContentLink;
            }

            return defaultFolder;
        }

        public void AddDocument(SPListItem item)
        {
            var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
            var contentTypeRepository = ServiceLocator.Current.GetInstance<IContentTypeRepository>();
            var mediaDataResolver = ServiceLocator.Current.GetInstance<ContentMediaResolver>();
            var blobFactory = ServiceLocator.Current.GetInstance<BlobFactory>();
            var fileName = item.File.Name; //  @"C:\test.jpg";

            //Get a suitable MediaData type from extension
            var mediaType = mediaDataResolver.GetFirstMatching(Path.GetExtension(fileName));
            var contentType = contentTypeRepository.Load(mediaType);

            // make sure the parent has the proper access so we can create objects in it
            ContentReference parentFolder;
            if (item.entryPointId != -1)
            {
                parentFolder = new ContentReference(item.entryPointId);
            }
            else
            {
                parentFolder = FindRecursive(item.ParentFolderName, SiteDefinition.Current.GlobalAssetsRoot);
            }
            AccessRightsHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess); 

            //Get a new empty media object
            var media = contentRepository.GetDefault<MediaData>(parentFolder, contentType.ID);

            // Set attibutes of the media
            media.Name = fileName;
            media.Created = item.File.TimeCreated;

            // Create a blob in the binary container
            if (item.File.DataStream != null)
            {
                var blob = blobFactory.CreateBlob(media.BinaryDataContainer, Path.GetExtension(fileName));
                blob.Write(item.File.DataStream);
                media.BinaryData = blob;
            }

            try
            {
                var file1ID = contentRepository.Save(media, SaveAction.Publish, AccessLevel.FullAccess);
                // contentRepository.Move(file1ID, parentContent);
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception caught.", e);
            }
        }

        /// <summary>
        /// Adds a new content folder into CMS 
        /// </summary>
        /// <param name="item"></param>
        public void AddFolder(SPListItem item)
        {
            // Set access level to 'Read' on the SharePoint's root folder and its children
            ContentReference cref;
            if (item.entryPointId != -1)
            {
                cref = new ContentReference(item.entryPointId);
            }
            else
            {
                cref = FindRecursive(item.ParentFolderName, SiteDefinition.Current.GlobalAssetsRoot); 
            }
            AccessRightsHelper.SetAccessRight(cref, AccessLevel.FullAccess);

            // Create a folder for the list under SharePoint entry point
            var cmsFolder = _contentRepository.GetDefault<ContentFolder>(cref);
            cmsFolder.Name = item.Folder.Name;
            cmsFolder.ParentLink = cref;
        
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

        public ContentReference FindRecursive(string name, ContentReference startRef)
        {
            var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
            var rootFolders = contentRepository.GetChildren<ContentFolder>(startRef);
            foreach (var rootFolder in rootFolders)
            {
                if (rootFolder.Name == name)
                    return rootFolder.ContentLink;
                else
                {
                    ContentReference cref = FindRecursive(name, rootFolder.ContentLink);
                    if (cref != null)
                        return cref;
                }
            }

            return null;
        }
    }
}
