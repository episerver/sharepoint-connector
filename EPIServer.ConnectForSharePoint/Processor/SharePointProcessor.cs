using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.DataAccess;
using EPiServer.Framework.Blobs;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using System;
using System.IO;

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
                // Only search from the entry point down
                parentFolder = ContentProviderHelper.FindRecursiveFolders(item.ParentFolderName, SharePointSettings.Current.EntryPoint);
            }
            AccessRightsHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess); 

            // Now check the folder to see if the item already exists. 
            ContentReference cref = ContentProviderHelper.FindMediaInFolder(fileName, parentFolder);
            if (cref == null)
            {
                //Get a new empty media object
                var media = contentRepository.GetDefault<MediaData>(parentFolder, contentType.ID);

                // Set attibutes of the media
                media.Name = fileName;
                media.Created = item.File.TimeCreated;

                try
                {
                    // Get the data and add it to the item
                    var relativeUrl = SharePointHelper.SPServicesHub.SPFileStreamService.GetFileRelativeUrl(item, false);
                    var data = SharePointHelper.SPServicesHub.SPFileStreamService.GetFileSteam(relativeUrl, false);

                    var blob = blobFactory.CreateBlob(media.BinaryDataContainer, Path.GetExtension(fileName));
                    blob.Write(data);
                    media.BinaryData = blob;
                }
                catch (Exception e)
                {
                    // todo we should be using logging
                    // failed to read data
                }

                try
                {
                    var file1ID = contentRepository.Save(media, SaveAction.Publish, AccessLevel.FullAccess);
                }
                catch (Exception e)
                {
                    // todo we should be using logging
                    // failed to save data in cms
                }
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
                cref = ContentProviderHelper.FindRecursiveFolders(item.ParentFolderName, SharePointSettings.Current.EntryPoint); 
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
            ContentReference RootRef = new ContentReference(item.entryPointId);
            AccessRightsHelper.SetAccessRight(RootRef, AccessLevel.FullAccess);

            var cmsFolder = _contentRepository.GetDefault<ContentFolder>(RootRef);
            var children = _contentRepository.GetChildren<ContentFolder>(RootRef);
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
