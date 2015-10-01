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

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessor : ISharePointProcessor
    {
        IServiceLocator _serviceLocator;
        IContentSecurityRepository _contentSecurityRepository;
        IContentProviderHelper _contentProviderHelper;
        IFileStreamService _fileStreamService;
        private ILogger _logger;

        [ExcludeFromCodeCoverage] // Exclude this constructor
        public SharePointProcessor()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
            _contentProviderHelper = new ContentProviderHelper();
            _fileStreamService = SharePointHelper.SPServicesHub.SPFileStreamService;
            _logger = LogManager.GetLogger();
        }

        internal SharePointProcessor(IServiceLocator serviceLocator, IContentProviderHelper contentProviderHelper, IFileStreamService fileStreamService)
        {
            _serviceLocator = serviceLocator;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
            _contentProviderHelper = contentProviderHelper;
            _fileStreamService = fileStreamService;
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
                if (media is ExtraMediaData) // ExtraMediaData is our custom attributes, basically a mapping from episerver object to sharepoint object
                {
                    var mediaEX = media as ExtraMediaData;
                    mediaEX.SharePointGUID = item.GUID.ToString();
                    mediaEX.SharePointMajorVersion = item.File.MajorVersion;
                    mediaEX.SharePointMinorVersion = item.File.MinorVersion;
                    //                                       mediaEX.ParentListID = item.File.ListItem.
                    mediaEX.SharePointTimeCreated = item.File.TimeCreated;
                    mediaEX.SharePointTimeLastModified = item.File.TimeLastModified;
                    mediaEX.SharePointRelativeUrl = item.File.ServerRelativeUrl;
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
                    DeleteSubFolders(child);

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
        internal void DeleteSubFolders(ContentFolder folder)
        {
            var repo = _serviceLocator.GetInstance<IContentRepository>();
            var folders = repo.GetChildren<ContentFolder>(folder.ContentLink);
            foreach (var f in folders)
            {
                DeleteSubFolders(f);

                // delete any media in the folder
                var medialist = repo.GetChildren<MediaData>(folder.ContentLink);
                foreach (var m in medialist)
                {
                    repo.Delete(m.ContentLink, true);
                }

                repo.Delete(f.ContentLink, true, AccessLevel.Delete);
                FolderData fd = new FolderData();
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

            // Now check the folder to see if the item already exists. 
            ContentReference cref = _contentProviderHelper.FindMediaInFolder(item.GUID, parentFolder);
            if (cref != null)
            {
                //Get a new empty media object
                var media = contentRepository.Get<MediaData>(cref).CreateWritableClone() as MediaData;

                if (media is ExtraMediaData)
                {
                    var mediaEX = media as ExtraMediaData;

                    // Only update if the last modified dates dont match
                    if (item.File.TimeLastModified.CompareTo(mediaEX.SharePointTimeLastModified) != 0)
                    {
                        // Set attibutes of the media
                        media.Name = fileName;
                        media.Created = item.File.TimeCreated;
                        mediaEX.SharePointGUID = item.GUID.ToString();
                        mediaEX.SharePointMajorVersion = item.File.MajorVersion;
                        mediaEX.SharePointMinorVersion = item.File.MinorVersion;
                        mediaEX.SharePointTimeCreated = item.File.TimeCreated;
                        mediaEX.SharePointTimeLastModified = item.File.TimeLastModified;
                        mediaEX.SharePointRelativeUrl = item.File.ServerRelativeUrl;
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
            var contentTypeRepository = _serviceLocator.GetInstance<IContentTypeRepository>();

            // make sure the parent has the proper access so we can delete objects in it
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
            if (cref != null)
            {
                //Get a new empty media object
                var media = contentRepository.Get<MediaData>(cref).CreateWritableClone() as MediaData;
                if (media != null)
                {
                    try
                    {
                        contentRepository.Delete(cref, true, AccessLevel.Delete);
                    }
                    catch
                    {
                        // todo we should be using logging
                        // failed to save data in cms
                    }
                }
            }
        }

        /// <summary>
        /// Parses the sharepoint xml for a list item, and returs a list of ListItemField
        /// </summary>
        /// <param name="XmlData"></param>
        /// <returns></returns>
        public List<ListItemField> getFields(string XmlData)
        {
            var list = new List<ListItemField>();
            try
            {
                using (XmlReader reader = XmlReader.Create(new StringReader(XmlData)))
                {
                    bool inProperties = false;
                    bool inHyperLink = false;
                    string name = "";
                    string value = "";
                    string type = "";
                    string description = "";
                    string url = "";
                    string lastElementName = "";

                    // Parse the file and display each of the nodes.
                    while (reader.Read())
                    {
                        switch (reader.NodeType)
                        {
                            case XmlNodeType.Element:

                                if (inProperties && inHyperLink != true)
                                {
                                    name = reader.Name;
                                    type = reader.GetAttribute("m:type");
                                    if (string.Equals(type, "SP.FieldUrlValue", StringComparison.OrdinalIgnoreCase))
                                        inHyperLink = true;
                                }

                                if (string.Equals(reader.Name, "m:properties", StringComparison.OrdinalIgnoreCase))
                                    inProperties = true;

                                lastElementName = reader.Name;
                                break;
                            case XmlNodeType.Text:
                                if (inProperties)
                                {
                                    if (!inHyperLink)
                                        value = reader.Value;
                                    else
                                    {
                                        if (string.Equals(lastElementName, "d:Description", StringComparison.OrdinalIgnoreCase))
                                            description = reader.Value;

                                        if (string.Equals(lastElementName, "d:Url", StringComparison.OrdinalIgnoreCase))
                                            url = reader.Value;
                                    }
                                }
                                break;

                            case XmlNodeType.XmlDeclaration:
                            case XmlNodeType.ProcessingInstruction:
                                Console.WriteLine(reader.Name + " " + reader.Value);
                                break;
                            case XmlNodeType.Comment:
                                Console.WriteLine(reader.Value);
                                break;

                            case XmlNodeType.EndElement:
                                if (inProperties)
                                {
                                    if (!inHyperLink)
                                    {
                                        if (!string.IsNullOrEmpty(name))
                                        {
                                            list.Add(new ListItemField()
                                            {
                                                Name = name.Substring(2), // remove the d:                                                    Type = type,
                                                Value = value,
                                                Description = description,
                                                Url = url
                                            });
                                            name = type = value = description = url = "";
                                        }
                                    }
                                    else // in hyperlink
                                    {
                                        if (string.Equals(name, reader.Name, StringComparison.OrdinalIgnoreCase))
                                        {
                                            inHyperLink = false;
                                            list.Add(new ListItemField()
                                            {
                                                Name = name.Substring(2), // remove the d:
                                                Type = type,
                                                Value = value,
                                                Description = description,
                                                Url = url
                                            });
                                            name = type = value = description = url = "";
                                        }
                                    }
                                }

                                if (string.Equals(reader.Name, "m:properties", StringComparison.OrdinalIgnoreCase))
                                    inProperties = false;

                                break;
                        }
                    }
                }
            }
            catch
            {
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
        /// Gets the xml data associated with the specific list item.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public XhtmlString getListItemData(SPListItem item)
        {
            var restUrl = item.FileUrl; // http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'f2c6da74-b21c-4f9b-ad50-bb51ae7b2bd5')/Items(1)/File

            // Need to remove the "/File" part in the Url
            var pos = restUrl.LastIndexOf('/');
            if (pos >= 0)
                restUrl = restUrl.Substring(0, pos);
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(_fileStreamService.GetListItemXmlData(restUrl));

            return new XhtmlString(xDoc.InnerXml); 
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
                media.Body = getListItemData(item); // get the xml data from the list item

                // Set attibutes of the media
                var content = media as IContent;
                if (content != null)
                    content.Name = fileName;

                media.Created = item.Created;
                if (media is ExtraMediaData) // ExtraMediaData is our custom attributes, basically a mapping from episerver object to sharepoint object
                {
                    var mediaEX = media as ExtraMediaData;
                    mediaEX.SharePointGUID = item.GUID.ToString();
                    mediaEX.SharePointTimeCreated = item.Created;
                    mediaEX.SharePointTimeLastModified = item.Modified;
                    mediaEX.SharePointRelativeUrl = item.FileUrl;
                }

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

            // Now check the folder to see if the item already exists. 
            ContentReference cref = _contentProviderHelper.FindMediaInFolder(item.GUID, parentFolder);
            if (cref != null)
            {
                //Get the media object
                var media = contentRepository.Get<SharePointBlockData>(cref).CreateWritableClone() as SharePointBlockData;
                if (media is ExtraMediaData) // ExtraMediaData is our custom attributes, basically a mapping from episerver object to sharepoint object
                {
                    var mediaEX = media as ExtraMediaData;

                    // Only update if the last modified dates dont match
                    if (item.Modified.CompareTo(mediaEX.SharePointTimeLastModified) != 0)
                    {
                        media.Body = getListItemData(item); // get the xml data from the list item

                        // Set attibutes of the media
                        var content = media as IContent;
                        if (content != null)
                            content.Name = fileName;

                        media.Created = item.Created;

                        mediaEX.SharePointGUID = item.GUID.ToString();
                        mediaEX.SharePointTimeCreated = item.Created;
                        mediaEX.SharePointTimeLastModified = item.Modified;
                        mediaEX.SharePointRelativeUrl = item.FileUrl;
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
            var contentTypeRepository = _serviceLocator.GetInstance<IContentTypeRepository>();

            // make sure the parent has the proper access so we can delete objects in it
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
            if (cref != null)
            {
                //Get a new empty media object
                var media = contentRepository.Get<BlockData>(cref).CreateWritableClone() as BlockData;
                if (media != null)
                {
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
    }
}
