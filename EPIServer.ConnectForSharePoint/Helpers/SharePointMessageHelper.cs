using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Models.FileTypes;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public class SharePointMessageHelper : ISharePointMessageHelper
    {
        private ILogger _logger;
        private ISharePointHelper _helper;
        private IContentProviderHelper _contentProviderHelper;
        private ISharePointSettingsHelper _settingsHelper;
        private IMessageEmitterFactory _emitterFactory;
        private IServiceLocator _serviceLocator;

        public SharePointMessageHelper()
        {
            _logger = LogManager.GetLogger();
            _helper = new SharePointHelper();
            _contentProviderHelper = new ContentProviderHelper();
            _settingsHelper = new SharePointSettingsHelper();
            _emitterFactory = new SharePointMessageEmitterFactory();
            _serviceLocator = ServiceLocator.Current;
        }

        internal SharePointMessageHelper(ILogger logger, ISharePointHelper helper, IContentProviderHelper contentHelper, ISharePointSettingsHelper settings, IMessageEmitterFactory emitter, IServiceLocator serviceLocator)
        {
            _logger = logger;
            _helper = helper;
            _contentProviderHelper = contentHelper;
            _settingsHelper = settings;
            _emitterFactory = emitter;
            _serviceLocator = serviceLocator;
        }

        /// <summary>
        /// Adds all the subfolders and subitems for the given parent folder, additionally deletes any removed items in a folder.
        /// </summary>
        /// <param name="parentFolder">The parent folder to use as selection criterea items and as a source for the cms folder name</param>
        /// <param name="selectedList">The list of items associated with all the selected views</param>
        /// <param name="selectedView">The selected views in the configuration</param>
        /// <param name="ParentFolderName">An optional parent folder name to use, i.e. in the case of views as a parent folder the name is actually different. </param>
        public void ProcessSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, string ParentFolderName = null)
        {
            var items = _helper.GetItems(selectedList, selectedView, parentFolder);
            var iterator = items.GetEnumerator();
            while (iterator.MoveNext())
            {
                var spItem = iterator.Current;
                spItem.entryPointId = -1; // we cant use the entrypointid, we have to use the parent folder name 
                spItem.rootCMSFolderEntryPointID = _settingsHelper.GetEntryPoint().ID;
                if (string.IsNullOrEmpty(ParentFolderName))
                    spItem.ParentFolderName = parentFolder.Name;
                else
                    spItem.ParentFolderName = ParentFolderName;

                if (spItem.Folder != null)
                {
                    // Add the folder if its not already there
                    var contentRef = AddOrUpdateFolder(spItem, _settingsHelper.GetEntryPoint());
                    if (contentRef != null)
                    {
                        // The folder was previously created so we can assume that we have been here before and this is a refresh.
                        // We must now check for media items that no longer exist in share point and must be deleted from cms.
                        // 1) get all the children items from that cms folder (contentRef)
                        // 2) iterate through each media item and get the quid from our extra media data instance.
                        // 3) check to see if the quid is in the sharepoint item list for this folder.
                        // 4) if not post the delete item message.
                        IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();
                        var cmsListItems = repo.GetChildren<MediaData>(contentRef);
                        var sharepointListItems = _helper.GetItems(selectedList, selectedView, spItem.Folder);
                        foreach (var item in cmsListItems)
                        {
                            // If its content
                            if (item is ExtraMediaData)
                            {
                                var mediaItem = item as ExtraMediaData;
                                Boolean delete = true;

                                // now iterate through the sharepoint list and check to see if its there. 
                                // If its not, it needs to be deleted.
                                foreach (var spItem2 in sharepointListItems)
                                {
                                    if (spItem2.GUID == new Guid(mediaItem.SharePointGUID))
                                    {
                                        delete = false; // the item is in both lists, we wont delete it. 
                                        break;
                                    }
                                }
                                if (delete == true)
                                {
                                    _emitterFactory.Create().Emit(new SharePointContentDeleteMessage()
                                    {
                                        Item = new SPListItem()
                                        {
                                            entryPointId = contentRef.ID, // id of the folder the item is in
                                            GUID = new Guid(mediaItem.SharePointGUID),
                                        },
                                    });
                                }
                            }
                        }

                        // Now we must do the same thing for removed folders
                        var cmsFolderItems = repo.GetChildren<ContentFolder>(contentRef);
                        FolderData fd = new FolderData();
                        foreach (var item in cmsFolderItems)
                        {
                            Boolean delete = true;

                            // now iterate through the sharepoint list and check to see if its there. 
                            // If its not, it needs to be deleted.
                            foreach (var spItem2 in sharepointListItems)
                            {
                                fd = fd.GetFolder(item.ContentLink.ID);

                                if (spItem2.GUID == fd.FolderGuid)
                                {
                                    delete = false; // the item is in both lists, we wont delete it. 
                                    break;
                                }
                            }
                            if (delete == true)
                            {
                                _emitterFactory.Create().Emit(new SharePointFolderDeleteMessage()
                                {
                                    Item = new SPListItem()
                                    {
                                        entryPointId = contentRef.ID, // id of the folder the item is in
                                        GUID = fd.FolderGuid,
                                        Folder = new SPFolder() { Name = item.Name}
                                    },
                                });
                            }
                        }
                    }

                    ProcessSubItems(spItem.Folder, selectedList, selectedView);
                }
                else if (spItem.File != null)
                {
                    ContentReference itemref = null; 
                    var parentFolderRef = _contentProviderHelper.FindRecursiveFolders(spItem.ParentFolderName, _settingsHelper.GetEntryPoint());
                    if( parentFolderRef != null )
                    {
                        // Determine if we are adding or updating and send the appropriate message.
                        itemref = _contentProviderHelper.FindMediaInFolder(spItem.GUID, parentFolderRef);
                    }

                    // if there is no parent folder, or the content doesnt exist send a Add message
                    if (parentFolderRef == null || itemref == null)
                    {
                        _logger.Information("Adding SharePointContentAddMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointContentAddMessage()
                        {
                            Item = spItem,
                        });
                    }
                    else
                    {
                        _logger.Information("Adding SharePointContentUpdateMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointContentUpdateMessage()
                        {
                            Item = spItem,
                        });
                    }
                }
            }
        }

        /// <summary>
        /// Deletes all folders in the specified splists collection, UNLESS there list ids to exclude from deletion.
        /// </summary>
        /// <param name="splists"></param>
        /// <param name="listids"></param>
        public void QueueFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null)
        {
            var selectedLists = _helper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists)
            {
                // The idListAndView is an object that contains two guids, one for the list and one for the view.
                var idListAndView = _helper.GetListAndViewIds(item.Key);

                // Get the selected list using the first guid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second guid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                // Make sure that the view guid and the view guid is in the listids, if its not then it needs to be deleted.
                var FolderName = selectedList.Title + " - " + selectedView.Title; // assume we are deleting it.
                if (listids != null) // we are some folders
                {
                    foreach (var id in listids)
                    {
                        var listIds = id.Split('~');
                        if (selectedList.Id == Guid.Parse(listIds[0]) && selectedView.Id == Guid.Parse(listIds[1]))
                        {
                            // We found a match, this means the view is in the current selection - don't delete it!  
                            // else all the items in the folder will be deleted and existing pages will lose content.
                            FolderName = string.Empty;
                            break;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(FolderName))
                {
                    // The message on the queue will be SPListItem
                    // Because we have only a list (like "documents") and a view (like "all documents") we need to create a SPListItem
                    // that contains some aspects of both.
                    SPListItem tbd = new SPListItem()
                    {
                        Title = selectedList.Title + " - " + selectedView.Title, // the title will be a combo of the list and view names
                        GUID = selectedView.Id,
                        entryPointId = _settingsHelper.GetEntryPoint().ID,
                        Folder = new SPFolder()
                        {
                            Name = FolderName,
                        },
                    };

                    _logger.Information("Deleting SharePointFolderAddMessage to the queue " + tbd.Title);
                    _emitterFactory.Create().Emit(new SharePointFolderDeleteMessage()
                    {
                        Item = tbd,
                    });
                }
            }
        }

        /// <summary>
        /// Adds all items to be added (SPFolders and SPItems) to the queue to be added
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        public void QueueItemsForAdd(string listids, IEnumerable<SPList> splists)
        {
            var selectedLists = _helper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists) // For each view selected
            {
                // The idListAndView is an object that contains two quids, one for the list and one for the view.
                var idListAndView = _helper.GetListAndViewIds(item.Key);

                // Get the selected list using the first quid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second quid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));
                var FolderName = selectedList.Title + " - " + selectedView.Title;
                var parentFolder = selectedList.RootFolder;

                // Dummy splistitem used for message to processor
                SPListItem viewListItem = new SPListItem()
                {
                    GUID = new Guid(idListAndView[0]),
                    entryPointId = _settingsHelper.GetEntryPoint().ID,
                    Title = selectedView.Title,
                    Folder = new SPFolder()
                    {
                        Name = FolderName
                    }
                };

                // Add or update the folder if its not already there
                AddOrUpdateFolder(viewListItem, _settingsHelper.GetEntryPoint());

                // Process the sub items of this folder.
                ProcessSubItems(parentFolder, selectedList, selectedView, FolderName);
            }
        }

        private ContentReference AddOrUpdateFolder(SPListItem item, ContentReference entryPoint)
        {
            var contentRef = _contentProviderHelper.FindRecursiveFolders(item.Folder.Name, entryPoint);
            if (contentRef == null)
            {
                Boolean add = true;

                // before we add it, lets see if the item is in the SPFolderMappings dynamic data store, if it is 
                // we can assume the name changed and we just need to updated it
                FolderData fd = new FolderData();
                FolderData foundData = fd.GetFolder(item.GUID);
                if (foundData.RefId != -1)
                {
                    // well we found it but it doesnt mean its the correct one. If the user changed the dest folder
                    // these items could have been marked for delete in previouse methods and are still in the queue
                    IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();
                    ContentFolder cf = repo.Get<ContentFolder>(new ContentReference(foundData.RefId));
                    if (cf != null) // did the item actually exist?
                    {
                        // walk up the parent list and see if its the current selected folder, if its not
                        // there it means the item is in the delete queue and the user selected a different folder.
                        // If it is under the current folder, it means the folder needs to be updated becuase it changed
                        // in sharepoint
                        Boolean done = false;
                        do
                        {
                            if (cf.ParentLink.ID == entryPoint.ID)
                            {
                                _logger.Information("SharePointFolderUpdateMessage to the queue : " + item.Folder.Name);
                                _emitterFactory.Create().Emit(new SharePointFolderUpdateMessage()
                                {
                                    Item = item,
                                });
                                add = false;
                                done = true;
                            } else
                                cf = repo.Get<ContentFolder>(new ContentReference(cf.ParentLink.ID));

                            if( cf == null || cf.ParentLink == null )
                            {
                                done = true;
                            }
                        } while (!done);
                    }
                    else
                    {
                        // remove the folder data, its invalid.
                        FolderData.Delete(foundData);
                    }
                }
                if( add )
                {
                    // Add the view folder.
                    _logger.Information("SharePointFolderAddMessage to the queue : " + item.Folder.Name);
                    _emitterFactory.Create().Emit(new SharePointFolderAddMessage()
                    {
                        Item = item,
                    });
                }
            }
            return contentRef;
        }
    }
}