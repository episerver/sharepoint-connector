using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.DataAccess;
using EPiServer.Framework.Blobs;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using EPiServer.Logging;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using EPiServer.ConnectForSharePoint.Rest;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessor : ISharePointProcessor
    {
        IServiceLocator _serviceLocator;
        IContentProviderHelper _contentProviderHelper;
        IFileStreamService _fileStreamService;
        IListsViewsService _listViewsService;
        ISharePointProcessorHelper _helper;

        private ILogger _logger;

        [ExcludeFromCodeCoverage] // Exclude this constructor
        public SharePointProcessor()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentProviderHelper = new ContentProviderHelper();
            
            var contextService = SharePointSettings.Current.IsModernAuth ?
            new EndpointContextService
            {
                EndpointContext = new SPRestClient(SharePointSettings.Current)
            } :
            new EndpointContextService
            {
                EndpointContext = new SPRestClient(SharePointSettings.Current, SharePointHelper.GetLatestCookies)
            };

            _fileStreamService = new FileStreamService { EndpointContextService = contextService };
            _listViewsService = new ListsViewsService { EndpointContextService = contextService };

            _helper = new SharePointProcessorHelper();
            _logger = LogManager.GetLogger();
        }

        internal SharePointProcessor(IServiceLocator serviceLocator, IContentProviderHelper contentProviderHelper, 
            IFileStreamService fileStreamService, IListsViewsService listViewsService, IAccessRightsHelper accessRightsHelper)
        {
            _serviceLocator = serviceLocator;
            _helper = _serviceLocator.GetInstance<ISharePointProcessorHelper>();
            _contentProviderHelper = contentProviderHelper;
            _fileStreamService = fileStreamService;
            _listViewsService = listViewsService;
            _logger = _serviceLocator.GetInstance<ILogger>();
        }

        public void AddDocument(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();
            var contentTypeRepository = _serviceLocator.GetInstance<IContentTypeRepository>();
            var mediaDataResolver = _serviceLocator.GetInstance<ContentMediaResolver>();
            var blobFactory = _serviceLocator.GetInstance<BlobFactory>();
            var fileName = item.File.Name; //  @"C:\test.jpg";

            _logger.Debug("- Adding doc : " + fileName); 

            //Get a suitable MediaData type from extension
            Type typetobuild = typeof(SharepointGenericMedia);
            var mediaType = mediaDataResolver.GetFirstMatching(Path.GetExtension(fileName));
            if (typeof(ImageData).IsAssignableFrom(mediaType))
                typetobuild = typeof(SharepointImageFile);
            else if (typeof(VideoData).IsAssignableFrom(mediaType))
                typetobuild = typeof(SharepointVideoFile);

            var contentType = contentTypeRepository.Load(typetobuild);

            // make sure the parent has the proper access so we can create objects in it
            ContentReference parentFolder;
            if (item.entryPointId != -1)
            {
                parentFolder = new ContentReference(item.entryPointId);
            }
            else
            {
                FolderData fd = new FolderData();
                var fdata = fd.GetFolder(item.ParentFolderSharePointGuid);
                parentFolder = new ContentReference(fdata.RefId);
            }

            // Now check the folder to see if the item already exists. 
            ContentReference cref = _contentProviderHelper.FindMediaInFolder(item.GUID, parentFolder);
            if (cref == null)
            {
                //Get a new empty media object
                var media = contentRepository.GetDefault<MediaData>(parentFolder, contentType.ID);

                // Set attibutes of the media
                media.Name = fileName;
                media.Created = item.File.TimeCreated;
                if (media is ISharePointItem) // ExtraMediaData is our custom attributes, basically a mapping from episerver object to sharepoint object
                {
                    var mediaEX = media as ISharePointItem;
                    mediaEX.SharePointGUID = item.GUID.ToString();
                    mediaEX.SharePointTimeLastModified = item.File.TimeLastModified;
                }

                try
                {
                    // Get the data and add it to the item
                    // Note: do not remove the getFileStream method, its public and available to 
                    // plugin authors so that they can get the data associated with a particular
                    // file in sharepoint
                    var data = _helper.GetDocumentStream(item);
                    var blob = blobFactory.CreateBlob(media.BinaryDataContainer, Path.GetExtension(fileName));

                    // Can't test this part of the code path...
                    blob.Write(data);
                    media.BinaryData = blob;
                }
                catch (Exception e)
                {
                    _logger.Error("Failed to create blog", e);
                }

                try
                {
                    var file1ID = contentRepository.Save(media, SaveAction.Publish, AccessLevel.FullAccess);
                }
                catch (Exception e)
                {
                    _logger.Error("Failed to save to cms", e);
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
                FolderData fd = new FolderData();
                var fdata = fd.GetFolder(item.ParentFolderSharePointGuid);
                cref = new ContentReference(fdata.RefId);
            }

            if (cref != null)
            {
                // Create a folder for the list under SharePoint entry point
                var cmsFolder = contentRepository.GetDefault<ContentFolder>(cref);
                cmsFolder.Name = item.Folder.Name;
                cmsFolder.ParentLink = cref;

                _logger.Debug("- Adding folder : " + item.Folder.Name);

                // Save the folder
                var savedRef = contentRepository.Save(cmsFolder, SaveAction.Publish, AccessLevel.FullAccess);

                // Save the sharepoint data
                var fd = new FolderData(savedRef.ID, item.GUID, item.Created, item.Modified);
                fd.Save();
            }
            else
            {
                _logger.Error("Parent folder " + item.ParentFolderName + " was not found. " + item.Folder.Name + " cannot be added. ");
            }
        }

        public void UpdateFolder(SPListItem item)
        {
            FolderData fd = new FolderData();
            var fdata = fd.GetFolder(item.GUID);
            if (fdata.RefId != -1)
            {
                var repo = _serviceLocator.GetInstance<IContentRepository>();
                var cmsFolder = repo.Get<ContentFolder>(new ContentReference(fdata.RefId)).
                    CreateWritableClone() as ContentFolder;
                cmsFolder.Name = item.Folder.Name;

                _logger.Debug("- Updating folder : " + item.Folder.Name);

                // figure out if the item has been moved.
                var parentfd = fd.GetFolder(item.ParentFolderSharePointGuid);
                if( parentfd.RefId != -1)
                {
                    if (cmsFolder.ParentLink.ID != parentfd.RefId)
                    {
                        repo.Move(cmsFolder.ContentLink, new ContentReference(parentfd.RefId));
                    }
                }

                var savedRef = repo.Save(cmsFolder, SaveAction.Publish, AccessLevel.FullAccess);
                fdata.Modified = item.Modified;
                fdata.Save();
            }
        }

        public void DeleteFolder(SPListItem item)
        {
            var repo = _serviceLocator.GetInstance<IContentRepository>();
            FolderData fd = new FolderData();
            var folderData = fd.GetFolder(item.GUID);
            if (folderData.RefId != -1)
            {
                ContentReference rootRef = new ContentReference(folderData.RefId);

                var children = repo.GetChildren<ContentFolder>(rootRef);
                foreach (var child in children)
                {
                    // before we delete this folder, delete all the sub folders
                    DeleteSubFolders(repo, child);

                    repo.Delete(child.ContentLink, true);
                    FolderData.Delete(fd.GetFolder(child.ContentLink.ID));
                }

                _logger.Debug("- Deleting folder : " + item.Folder.Name);

                // now delete this folder
                repo.Delete(new ContentReference(folderData.RefId), true);
                FolderData.Delete(folderData);
            }
        }

        /// <summary>
        /// Traverses to the bottom of the tree under this folder and deletes each item back to the ContentFolder specified.
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="repo">Pass this in so we aren't opening up a connection to the db every time, especially when we recurse through the tree.</param>
        internal void DeleteSubFolders(IContentRepository repo, ContentFolder folder)
        {
            var folders = repo.GetChildren<ContentFolder>(folder.ContentLink);
            foreach (var f in folders)
            {
                // recurse over each child before deleting parent folder
                DeleteSubFolders(repo, f);

                // delete any media in the current folder
                var medialist = repo.GetChildren<MediaData>(folder.ContentLink);
                foreach (var m in medialist)
                {
                    try
                    {
                        repo.Delete(m.ContentLink, true);
                    } catch { } // intentionally empty catch, MAI-933
                }               

                try
                {
                    repo.Delete(f.ContentLink, true, AccessLevel.Delete);
                } catch { } // intentionally empty catch, MAI-933

                var fd = new FolderData();
                FolderData.Delete(fd.GetFolder(f.ContentLink.ID));
            }
        }

        public void UpdateDocument(SPListItem item)
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
                FolderData fd = new FolderData();
                var fdata = fd.GetFolder(item.ParentFolderSharePointGuid);
                parentFolder = new ContentReference(fdata.RefId);
            }

            if (item.ContentLink != null)
            {
                //Get a new empty media object
                var media = contentRepository.Get<MediaData>(item.ContentLink).CreateWritableClone() as MediaData;
                if( media.ParentLink.ID != parentFolder.ID )
                {
                    contentRepository.Move(media.ContentLink, parentFolder);
                }

                if (media is ISharePointItem)
                {
                    var mediaEX = media as ISharePointItem;

                    // Only update if the last modified dates dont match
                    if (item.File.TimeLastModified.CompareTo(mediaEX.SharePointTimeLastModified) != 0)
                    {
                        _logger.Debug("- Update Document : " + fileName);

                        // Set attibutes of the media
                        media.Name = fileName;
                        media.Created = item.File.TimeCreated;
                        mediaEX.SharePointGUID = item.GUID.ToString();
                        mediaEX.SharePointTimeLastModified = item.File.TimeLastModified;
                        try
                        {
                            // Get the data and add it to the item
                            // Note: do not remove the getFileStream method, its public and available to 
                            // plugin authors so that they can get the data associated with a particular
                            // file in sharepoint
                            var data = _helper.GetDocumentStream(item);
                            var blob = blobFactory.CreateBlob(media.BinaryDataContainer, Path.GetExtension(fileName));

                            // Can't test this part of the code path...
                            blob.Write(data);
                            media.BinaryData = blob;
                        }
                        catch
                        {
                            // todo we should be using logging
                            // failed to read data
                        }

                        try
                        {
                            var file1ID = contentRepository.Save(media, SaveAction.Publish, AccessLevel.FullAccess);
                        }
                        catch
                        {
                            // todo we should be using logging
                            // failed to save data in cms
                        }
                    }
                    else
                    {
                        _logger.Debug("- Nothing to Update for Document : " + fileName);
                    }
                }
            }
        }

        public void DeleteDocument(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();

            try
            {
                _logger.Debug("- Delete Document : " + item.File.Title);
                contentRepository.Delete(item.ContentLink, true, AccessLevel.Delete);
            }
            catch(Exception e)
            {
                _logger.Error("Failed to delete document data from cms", e);
            }
        }

        public void AddListItem(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();
            var contentTypeRepository = _serviceLocator.GetInstance<IContentTypeRepository>();
            var mediaDataResolver = _serviceLocator.GetInstance<ContentMediaResolver>();
            var fileName = item.DisplayName; //  @"C:\test.jpg";

            _logger.Debug("- Add Custom List item : " + fileName);

            //Get a suitable MediaData type from extension
            Type typetobuild = typeof(SharePointBlockData);
            var contentType = contentTypeRepository.Load(typetobuild);

            // make sure the parent has the proper access so we can create objects in it
            ContentReference parentFolder;
            if (item.entryPointId != -1)
            {
                parentFolder = new ContentReference(item.entryPointId);
            }
            else
            {
                FolderData fd = new FolderData();
                var fdata = fd.GetFolder(item.ParentFolderSharePointGuid);
                parentFolder = new ContentReference(fdata.RefId);
            }

            // Now check the folder to see if the item already exists. 
            ContentReference cref = _contentProviderHelper.FindMediaInFolder(item.GUID, parentFolder);
            if (cref == null)
            {
                //Get a new empty media object
                var media = contentRepository.GetDefault<SharePointBlockData>(parentFolder, contentType.ID);
                media.Json = _helper.GetRawListItemData(item); // get the json data from the list item

                // Set attibutes of the media
                var content = media as IContent;
                if (content != null)
                    content.Name = fileName;

                // Set the view being used.
                if (item.selectedView == null)
                    media.ViewID = item.ParentList.Views.First().Id.ToString();
                else
                    media.ViewID = item.selectedView.Id.ToString();
                 
                media.Created = item.Created;
                media.SharePointGUID = item.GUID.ToString();
                media.SharePointTimeLastModified = item.Modified;
                media.ListID = item.ParentList.Id.ToString();

                try
                {
                    var file1ID = contentRepository.Save(media as IContent, SaveAction.Publish, AccessLevel.FullAccess);
                }
                catch (Exception e)
                {
                    _logger.Error("Failed to save to cms", e);
                }
            }
        }

        public void UpdateListItem(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();
            var contentTypeRepository = _serviceLocator.GetInstance<IContentTypeRepository>();
            var mediaDataResolver = _serviceLocator.GetInstance<ContentMediaResolver>();
            var fileName = item.DisplayName; //  @"C:\test.jpg";

            //Get a suitable MediaData type from extension
            Type typetobuild = typeof(SharePointBlockData);
            var contentType = contentTypeRepository.Load(typetobuild);

            // make sure the parent has the proper access so we can create objects in it
            ContentReference parentFolder;
            if (item.entryPointId != -1)
            {
                parentFolder = new ContentReference(item.entryPointId);
            }
            else
            {
                FolderData fd = new FolderData();
                var fdata = fd.GetFolder(item.ParentFolderSharePointGuid);
                parentFolder = new ContentReference(fdata.RefId);
            }

            if (item.ContentLink != null)
            {
                //Get the media object
                var media = contentRepository.Get<SharePointBlockData>(item.ContentLink).CreateWritableClone() as SharePointBlockData;
                if (media is ISharePointItem) // ExtraMediaData is our custom attributes, basically a mapping from episerver object to sharepoint object
                {
                    var mediaEX = media as ISharePointItem;
                    // Get the current view.
                    SPListView view = null;
                    if (item.selectedView == null)
                        view = item.ParentList.Views.First();
                    else
                        view = item.selectedView;

                    // Only update if the last modified dates dont match
                    // or the view has changed.
                    if (item.Modified.CompareTo(mediaEX.SharePointTimeLastModified) != 0 ||
                        view.Id != Guid.Parse(mediaEX.ViewID) )
                    {
                        _logger.Debug("- Update Custom List item : " + fileName);

                        media.Json = _helper.GetRawListItemData(item); // get the json data from the list item

                        // Set attibutes of the media
                        var content = media as IContent;
                        if (content != null)
                            content.Name = fileName;

                        media.Created = item.Created;

                        mediaEX.ViewID = view.Id.ToString();
                        mediaEX.SharePointGUID = item.GUID.ToString();
                        mediaEX.SharePointTimeLastModified = item.Modified;
                        try
                        {
                            var file1ID = contentRepository.Save(media as IContent, SaveAction.Publish, AccessLevel.FullAccess);
                        }
                        catch (Exception e)
                        {
                            _logger.Error("Failed to save to cms", e);
                        }
                    }
                    else
                    {
                        _logger.Debug("- Nothing to Update Custom List item : " + fileName);
                    }
                }
            }
        }

        public void DeleteListItem(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();

            try
            {
                _logger.Debug("- Delete Custom List item : " + item.Title);

                contentRepository.Delete(item.ContentLink, true, AccessLevel.Delete);
            }
            catch (Exception e)
            {
                _logger.Error("Failed to delete block data from cms", e);
            }
        }
    }
}
