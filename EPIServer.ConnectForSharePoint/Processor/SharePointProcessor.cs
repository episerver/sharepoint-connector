using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.DataAccess;
using EPiServer.Framework.Blobs;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessor : ISharePointProcessor
    {
        IServiceLocator _serviceLocator;
        IContentSecurityRepository _contentSecurityRepository;
        IContentProviderHelper _contentProviderHelper;
        IFileStreamService _fileStreamService;

        [ExcludeFromCodeCoverage] // Exclude this constructor
        public SharePointProcessor()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
            _contentProviderHelper = new ContentProviderHelper();
            _fileStreamService = SharePointHelper.SPServicesHub.SPFileStreamService;
        }

        internal SharePointProcessor(IServiceLocator serviceLocator, IContentProviderHelper contentProviderHelper, IFileStreamService fileStreamService)
        {
            _serviceLocator = serviceLocator;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
            _contentProviderHelper = contentProviderHelper;
            _fileStreamService = fileStreamService;
        }

        public void AddDocument(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();
            var contentTypeRepository = _serviceLocator.GetInstance<IContentTypeRepository>();
            var mediaDataResolver = _serviceLocator.GetInstance<ContentMediaResolver>();
            var blobFactory = _serviceLocator.GetInstance<BlobFactory>();
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
                parentFolder = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName, new ContentReference(item.rootCMSFolderEntryPointID));
            }
            _contentProviderHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess); 

            // Now check the folder to see if the item already exists. 
            ContentReference cref = _contentProviderHelper.FindMediaInFolder(fileName, parentFolder);
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
                    var relativeUrl = _fileStreamService.GetFileRelativeUrl(item);
                    var data = _fileStreamService.GetFileSteam(relativeUrl);

                    var blob = blobFactory.CreateBlob(media.BinaryDataContainer, Path.GetExtension(fileName));

                    // Can't test this part of the code path...
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
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();
            if (item.entryPointId != -1)
            {
                cref = new ContentReference(item.entryPointId);
            }
            else
            {
                cref = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName, 
                    new ContentReference(item.rootCMSFolderEntryPointID)); 
            }

            _contentProviderHelper.SetAccessRight(cref, AccessLevel.FullAccess);

            // Create a folder for the list under SharePoint entry point
            var cmsFolder = contentRepository.GetDefault<ContentFolder>(cref);
            cmsFolder.Name = item.Folder.Name;
            cmsFolder.ParentLink = cref;

            // Save the folder
            contentRepository.Save(cmsFolder, SaveAction.Publish, AccessLevel.FullAccess);
        }

        public void DeleteFolder(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();
            // Set access level to 'Read' on the SharePoint's root folder and its children
            ContentReference RootRef = new ContentReference(item.entryPointId);

            _contentProviderHelper.SetAccessRight(RootRef, AccessLevel.FullAccess);

            var children = contentRepository.GetChildren<ContentFolder>(RootRef);
            foreach (var child in children)
            {
                if( child.Name.Equals(item.Folder.Name) )
                {
                    contentRepository.Delete(child.ContentLink,true,AccessLevel.Delete);
                    break;
                }
            }
        }

        public void UpdateDocument(SPListItem item)
        {
            // When you implement these make sure you do the unit tests
            throw new NotImplementedException();
        }

        public void UpdateFolder(SPListItem item)
        {
            // When you implement these make sure you do the unit tests
            throw new NotImplementedException();
        }

        public void DeleteDocument(SPListItem item)
        {
            // When you implement these make sure you do the unit tests
            throw new NotImplementedException();
        }


    }
}
