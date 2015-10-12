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
using EPiServer;
using System.Xml;
using System.Net;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessor : ISharePointProcessor
    {
        IServiceLocator _serviceLocator;
        IContentSecurityRepository _contentSecurityRepository;
        IContentProviderHelper _contentProviderHelper;
        IFileStreamService _fileStreamService;
        IListsViewsService _listViewsService;

        private ILogger _logger;

        [ExcludeFromCodeCoverage] // Exclude this constructor
        public SharePointProcessor()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
            _contentProviderHelper = new ContentProviderHelper();
            _fileStreamService = SharePointHelper.SPServicesHub.SPFileStreamService;
            _listViewsService = SharePointHelper.SPServicesHub.SPListsViewsService;
            _logger = LogManager.GetLogger();
        }

        internal SharePointProcessor(IServiceLocator serviceLocator, IContentProviderHelper contentProviderHelper, 
            IFileStreamService fileStreamService, IListsViewsService listViewsService)
        {
            _serviceLocator = serviceLocator;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
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
                // Only search from the entry point down
                parentFolder = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName, new ContentReference(item.rootCMSFolderEntryPointID));
            }
            _contentProviderHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess);

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
                    var data = getFileStream(item);
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
                // Its ok to search for the parent folder by name, its not going to be changing ever during this 
                // message.
                cref = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName,
                    new ContentReference(item.rootCMSFolderEntryPointID));
            }

            _contentProviderHelper.SetAccessRight(cref, AccessLevel.FullAccess);

            // Create a folder for the list under SharePoint entry point
            var cmsFolder = contentRepository.GetDefault<ContentFolder>(cref);
            cmsFolder.Name = item.Folder.Name;
            cmsFolder.ParentLink = cref;

            // Save the folder
            var savedRef = contentRepository.Save(cmsFolder, SaveAction.Publish, AccessLevel.FullAccess);

            // Save the sharepoint data
            var fd = new FolderData(savedRef.ID, item.GUID, item.Created, item.Modified);
            fd.Save();
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
                var savedRef = repo.Save(cmsFolder, SaveAction.Publish, AccessLevel.FullAccess);

                fdata.Modified = item.Modified;
                fdata.Save();
            }
        }

        public void DeleteFolder(SPListItem item)
        {
            var repo = _serviceLocator.GetInstance<IContentRepository>();

            // Set access level to 'Read' on the SharePoint's root folder and its children
            ContentReference rootRef = new ContentReference(item.entryPointId);
            _contentProviderHelper.SetAccessRight(rootRef, AccessLevel.FullAccess);

            var children = repo.GetChildren<ContentFolder>(rootRef);
            foreach (var child in children)
            {
                if (child.Name.Equals(item.Folder.Name, StringComparison.OrdinalIgnoreCase))
                {
                    // before we delete this folder, delete all the sub folders
                    DeleteSubFolders(repo, child);

                    repo.Delete(child.ContentLink, true, AccessLevel.Delete);
                    FolderData fd = new FolderData();
                    FolderData.Delete(fd.GetFolder(child.ContentLink.ID));
                    break;
                }
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
                    repo.Delete(m.ContentLink, true);
                }

                repo.Delete(f.ContentLink, true, AccessLevel.Delete);
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
                // Only search from the entry point down
                parentFolder = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName, new ContentReference(item.rootCMSFolderEntryPointID));
            }
            _contentProviderHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess);

            if (item.ContentLink != null)
            {
                //Get a new empty media object
                var media = contentRepository.Get<MediaData>(item.ContentLink).CreateWritableClone() as MediaData;

                if (media is ISharePointItem)
                {
                    var mediaEX = media as ISharePointItem;

                    // Only update if the last modified dates dont match
                    if (item.File.TimeLastModified.CompareTo(mediaEX.SharePointTimeLastModified) != 0)
                    {
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
                            var data = getFileStream(item);
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
                }
            }
        }

        public void DeleteDocument(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();

            // Now check the folder to see if the item already exists. 
            ContentReference cref = new ContentReference(item.entryPointId);
            try
            {
                contentRepository.Delete(cref, true, AccessLevel.Delete);
            }
            catch(Exception e)
            {
                _logger.Error("Failed to delete document data from cms", e);
            }
        }

        /// <summary>
        /// Parses the sharepoint json data for a listitem, and returns a list of ListItemField
        /// </summary>
        /// <param name="jsondata"></param>
        /// <returns></returns>
        public List<ListItemField> getFields(Guid ListID, string jsondata)
        {
            var list = new List<ListItemField>();
            JObject obj = JObject.Parse(jsondata);
            if (obj["d"]["results"].HasValues == false)
                return list;

            // get the dictionary
            var columnDictionary = _listViewsService.GetFieldItemDictionary(ListID);

            var data = obj["d"]["results"][0].Value<JObject>();
            List<string> keys = data.Properties().Select(p => p.Name).ToList();
            foreach (var prop in data.Properties())
            {
                string property = prop.Name;
                string propvalue = prop.Value.ToString();
                string url = "";
                string description = "";
                bool addToList = true;

                // get the DisplayName aka title from the property dictionary for this list
                SPFieldItem field;
                if (columnDictionary.TryGetValue(property, out field))
                {
                    property = field.Title;
                }
                else
                {
                    // well this is wierd, sometimes the value we get from the sharepoint webservice
                    // has data in a different format like OData__"property" - hack
                    //
                    foreach( var entry in columnDictionary)
                    {
                        if( property.Contains(entry.Key) )
                        {
                            property = entry.Value.Title;
                            break;
                        }
                    }
                }

                try
                {
                    if (obj["d"]["results"].HasValues == true )
                    {
                        var data2 = obj["d"]["results"][0][prop.Name].Value<JObject>();
                        foreach (var prop2 in data2.Properties())
                        {
                            string property2 = prop2.Name;
                            string propvalue2 = prop2.Value.ToString();

                            if (string.Equals(property2, "Url", StringComparison.OrdinalIgnoreCase)) {
                                url = propvalue2;
                            }
                            if (string.Equals(property2, "Description", StringComparison.OrdinalIgnoreCase)) {
                                description = propvalue2;
                                propvalue = description;
                            }
                            if (string.Equals(property2, "__deferred", StringComparison.OrdinalIgnoreCase))
                            { // Dont add these things they are urls to objects in sharepoint for the folder, file etc.
                                addToList = false;
                            }
                        }
                    }
                } catch
                {

                }
                
                // More nodes we dont want to add, these actually come back by default with the view
                if (string.Equals(property, "FileSystemObjectType", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "__metadata", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "IDDD", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "ID", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "created", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "modified", StringComparison.OrdinalIgnoreCase))
                {
                    addToList = false;
                }

                if ( addToList )
                    list.Add(new ListItemField() { Name = property, Value = propvalue, Url = url, Description = description });
            }
            return list;
        }

        /// <summary>
        /// Gets the stream data assocatied with a specific SPListItem
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public Stream getFileStream(SPListItem item)
        {
            var relativeUrl = _fileStreamService.GetFileRelativeUrl(item);
            return _fileStreamService.GetFileSteam(relativeUrl);
        }

        /// <summary>
        /// Gets the json encoded data associated with the specific list item.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public string getListItemData(SPListItem item)
        {
            var restUrl = item.FileUrl; // http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'f2c6da74-b21c-4f9b-ad50-bb51ae7b2bd5')/Items(1)/File
            return _fileStreamService.GetListItemJsonData(item);
        }

        public void AddListItem(SPListItem item)
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
                // Only search from the entry point down
                parentFolder = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName, new ContentReference(item.rootCMSFolderEntryPointID));
            }
            _contentProviderHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess);

            // Now check the folder to see if the item already exists. 
            ContentReference cref = _contentProviderHelper.FindMediaInFolder(item.GUID, parentFolder);
            if (cref == null)
            {
                //Get a new empty media object
                var media = contentRepository.GetDefault<SharePointBlockData>(parentFolder, contentType.ID);
                media.Body = new XhtmlString(getListItemData(item)); // get the xml data from the list item

                // Set attibutes of the media
                var content = media as IContent;
                if (content != null)
                    content.Name = fileName;

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
                // Only search from the entry point down
                parentFolder = _contentProviderHelper.FindRecursiveFolders(item.ParentFolderName, new ContentReference(item.rootCMSFolderEntryPointID));
            }
            _contentProviderHelper.SetAccessRight(parentFolder, AccessLevel.FullAccess);

            if (item.ContentLink != null)
            {
                //Get the media object
                var media = contentRepository.Get<SharePointBlockData>(item.ContentLink).CreateWritableClone() as SharePointBlockData;
                if (media is ISharePointItem) // ExtraMediaData is our custom attributes, basically a mapping from episerver object to sharepoint object
                {
                    var mediaEX = media as ISharePointItem;

                    // Only update if the last modified dates dont match
                    if (item.Modified.CompareTo(mediaEX.SharePointTimeLastModified) != 0)
                    {
                        media.Body = new XhtmlString(getListItemData(item)); // get the xml data from the list item

                        // Set attibutes of the media
                        var content = media as IContent;
                        if (content != null)
                            content.Name = fileName;

                        media.Created = item.Created;

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
                }
            }
        }

        public void DeleteListItem(SPListItem item)
        {
            var contentRepository = _serviceLocator.GetInstance<IContentRepository>();

            // Now check the folder to see if the item already exists. 
            ContentReference cref = new ContentReference(item.entryPointId);
            try
            {
                contentRepository.Delete(cref, true, AccessLevel.Delete);
            }
            catch (Exception e)
            {
                _logger.Error("Failed to delete block data from cms", e);
            }
        }
    }
}
