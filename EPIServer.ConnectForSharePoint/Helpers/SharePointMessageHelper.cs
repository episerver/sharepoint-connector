using System;
using System.Collections.Generic;
using System.Linq;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.Web;
using EPiServer.DataAbstraction;

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

        private void DeleteDocumentsAndItems(IEnumerable<IContent> cmsListItems, IEnumerable<SPListItem> sharepointListItems)
        {
            foreach (var item in cmsListItems)
            {
                // If its content
                if (item is ISharePointItem)
                {
                    var mediaItem = item as ISharePointItem;
                    Boolean found = false;

                    // now iterate through the sharepoint list and check to see if its there. 
                    // If its not, it needs to be deleted.
                    foreach (var spItem2 in sharepointListItems)
                    {
                        if (spItem2.GUID == new Guid(mediaItem.SharePointGUID))
                        {
                            found = true; // the item is in both lists, we wont delete it. 
                            break;
                        }
                    }
                    if (!found)
                    {
                        // if its a list item, delete it
                        if (mediaItem.Type == SharePointItemType.ListItem )
                        {
                            _emitterFactory.Create().Emit(new SharePointListItemDeleteMessage()
                            {
                                Item = new SPListItem()
                                {
                                    entryPointId = item.ContentLink.ID, // id of the folder the item is in
                                    GUID = new Guid(mediaItem.SharePointGUID),
                                },
                            });
                        }
                        else if(mediaItem.Type == SharePointItemType.LibraryItem )
                        {
                            _emitterFactory.Create().Emit(new SharePointContentDeleteMessage()
                            {
                                Item = new SPListItem()
                                {
                                    entryPointId = item.ContentLink.ID, // id of the folder the item is in
                                    GUID = new Guid(mediaItem.SharePointGUID),
                                },
                            });
                        }
                    }
                }
            }
        }

        public void RemoveSharePointAddedFolders(ContentReference folderRef, IEnumerable<SPListItem> sharepointListItems)
        {
            IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();

            var cmsFolderItems = repo.GetChildren<ContentFolder>(folderRef);
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
                            entryPointId = folderRef.ID, // id of the folder the item is in
                            GUID = fd.FolderGuid,
                            Folder = new SPFolder() { Name = item.Name }
                        },
                    });
                }
            }
        }

        /// <summary>
        /// manages the folder by updating the name if it has changed, adding and removing items 
        /// </summary>
        /// <param name="spItem"></param>
        /// <param name="selectedList"></param>
        /// <param name="selectedView"></param>
        private void manageFolder(SPListItem spItem, SPList selectedList, SPListView selectedView, IEnumerable<SPListItem> sharepointListItems = null)
        {
            if (spItem.Folder != null)
            {
                // Add the folder if its not already there
                var folderRef = AddOrUpdateFolder(spItem, _settingsHelper.GetEntryPoint());
                if (folderRef != null)
                {
                    // The folder was previously created so we can assume that we have been here before and this is a refresh.
                    // get the list of items that should be in this folder and call processLists that will 
                    if (sharepointListItems == null)
                    {
                        sharepointListItems = _helper.GetItems(selectedList, selectedView, spItem.Folder);
                    }

                    RemoveSharePointAddedFolders(folderRef, sharepointListItems);
                }
            }
        }

        /// <summary>
        /// Adds or updates SPListItems that are not folders 
        /// </summary>
        /// <param name="spItem"></param>
        private void AddOrUpdate(SPListItem spItem, IEnumerable<IContent> cmsItems)
        {
            if (spItem.Folder == null)
            {
                IContent itemref = null;

                foreach (var cmsItem in cmsItems)
                {
                    if (cmsItem is ISharePointItem)
                    {
                        var thisItem = cmsItem as ISharePointItem;
                        if (spItem.GUID.ToString().Equals(thisItem.SharePointGUID))
                        {
                            itemref = cmsItem;
                            spItem.ContentLink = cmsItem.ContentLink;
                        }
                    }
                }

                if (spItem.File != null) // is it a document?
                {   
                    // if there is no parent folder, or the content doesnt exist send a Add message
                    if ( itemref == null)
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
                else // it is a list item
                {   
                    // if there is no parent folder, or the content doesnt exist send a Add message
                    if ( itemref == null)
                    {
                        _logger.Information("Adding SharePointListItemAddMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointListItemAddMessage()
                        {
                            Item = spItem,
                        });
                    }
                    else
                    {
                        _logger.Information("Adding SharePointListItemUpdateMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointListItemUpdateMessage()
                        {
                            Item = spItem,
                        });
                    }
                }
            }
        }

        /// <summary>
        /// Adds all the subfolders and subitems for the given parent folder, additionally deletes any removed items in a folder.
        /// </summary>
        /// <param name="parentFolder">The parent folder to use as selection criterea items and as a source for the cms folder name</param>
        /// <param name="selectedList">The list of items associated with all the selected views</param>
        /// <param name="selectedView">The selected views in the configuration</param>
        /// <param name="ParentFolderName">An optional parent folder name to use, i.e. in the case of views as a parent folder the name is actually different. </param>
        public List<SPListItem> ProcessSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, string ParentFolderName = null)
        {
            List<SPListItem> nonFolderItems = new List<SPListItem>();

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

                if (spItem.Folder == null)
                {
                    spItem.selectedView = selectedView;
                    nonFolderItems.Add(spItem);
                }
                else
                {
                    // Add the folder if its not already there
                    var folderRef = AddOrUpdateFolder(spItem, _settingsHelper.GetEntryPoint());
                    if (folderRef != null)
                    {
                        // The folder was previously created so we can assume that we have been here before and this is a refresh.
                        // get the list of items that should be in this folder and call processLists that will 
                        var subItems = _helper.GetItems(selectedList, selectedView, spItem.Folder);
                        RemoveSharePointAddedFolders(folderRef, subItems);
                    }

                    nonFolderItems.AddRange( ProcessSubItems(spItem.Folder, selectedList, selectedView) );
                }
            }
            return nonFolderItems;
        }

        /// <summary>
        /// Deletes all folders in the specified splists collection, UNLESS there list ids to exclude from deletion.
        /// </summary>
        /// <param name="splists"></param>
        /// <param name="listids"></param>
        public void queueRootFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null)
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

        private SPListItem CreateParentSPListItem(SPList selectedList, SPListView selectedView)
        {
            var FolderName = selectedList.Title + " - " + selectedView.Title;
            var parentFolder = selectedList.RootFolder;

            // Dummy SPListItem used for message to processor
            SPListItem item = new SPListItem()
            {
                GUID = selectedList.Id,
                entryPointId = _settingsHelper.GetEntryPoint().ID,
                Title = selectedView.Title,
                Folder = new SPFolder()
                {
                    Name = FolderName
                }
            };

            return item;
        }

        private void RemoveUserAddedFolders(ContentReference cf)
        {
            IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();
            var cmsFolderItems = repo.GetChildren<ContentFolder>(cf);
            FolderData fd = new FolderData();
            foreach (var folder in cmsFolderItems)
            {
                fd = fd.GetFolder(folder.ContentLink.ID);
                if (fd.RefId == -1)
                {
                    _emitterFactory.Create().Emit(new SharePointFolderDeleteMessage()
                    {
                        Item = new SPListItem()
                        {
                            entryPointId = cf.ID, // id of the folder the item is in
                            Folder = new SPFolder() { Name = folder.Name }
                        },
                    });
                }
            }
        }

        /// <summary>
        /// Adds all items to be added (SPFolders and SPItems) to the queue to be added
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        public void BeginProcessing(string listids, IEnumerable<SPList> splists)
        {
            List<SPListItem> AllSPListItems = new List<SPListItem>();

            // get our selected lists and iterate through them
            var selectedLists = _helper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists) // For each list/view selected
            {
                // The idListAndView is an object that contains two guids, one for the list and one for the view.
                var idListAndView = _helper.GetListAndViewIds(item.Key);

                // split the list and view Guid and get the list and view objects 
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));
                var sharepointListItems = _helper.GetItems(selectedList, selectedView, selectedList.RootFolder);

                // create and or update our root folder for this list/view combo
                SPListItem rootFolder = CreateParentSPListItem(selectedList, selectedView);
                var folderRef = AddOrUpdateFolder(rootFolder, _settingsHelper.GetEntryPoint());
                if(folderRef != null ) // if folder exists
                {
                    // Remove any folders the user may have added in cms or deleted from sharepoint
                    RemoveUserAddedFolders(folderRef);
                    RemoveSharePointAddedFolders(folderRef, sharepointListItems);
                }

                // All list items that are not folders need the view object
                foreach (var spItem in sharepointListItems)
                {
                    if (spItem.Folder == null)
                    {
                        spItem.selectedView = selectedView;
                        spItem.ParentFolderName = rootFolder.Folder.Name;
                        AllSPListItems.Add(spItem); 
                    }
                }

                // Process the sub items of this folder.
                AllSPListItems.AddRange( 
                    ProcessSubItems(selectedList.RootFolder, selectedList, selectedView, rootFolder.Folder.Name) );
            }

            // get all of our content from cms
            var contentItems = FindSharePointItems();
            foreach (var item in AllSPListItems)
            {
                AddOrUpdate(item, contentItems);
            }
            
            // delete all items (if any) from cms in this list that have been removed from sharepoint
            DeleteDocumentsAndItems(contentItems, AllSPListItems);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="entryPoint">The current entry point that is saved in the settings.</param>
        /// <returns></returns>
        private ContentReference AddOrUpdateFolder(SPListItem item, ContentReference entryPoint)
        {
            var contentRef = _contentProviderHelper.FindRecursiveFolders(item.Folder.Name, entryPoint);
            if (contentRef == null)
            {
                var add = true;

                // before we add it, lets see if the item is in the SPFolderMappings dynamic data store, if it is 
                // we can assume the name changed and we just need to update it
                var fd = new FolderData();
                var foundData = fd.GetFolder(item.GUID);

                if (foundData.RefId != -1)
                {
                    // we found it, but it doesn't mean it's the correct one. If the user changed the dest folder
                    // these items could have been marked for delete in previous methods and are still in the queue
                    var repo = _serviceLocator.GetInstance<IContentRepository>();
                    var cf = repo.Get<ContentFolder>(new ContentReference(foundData.RefId));

                    if (cf != null) // did the item actually exist?
                    {
                        // walk up the parent list and see if it's the current selected folder, if it's not
                        // there it means the item is in the delete queue and the user selected a different folder.
                        // If it is under the current folder, it means the folder needs to be updated becuase it changed
                        // in sharepoint.
                        var done = false;
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
                            }
                            else
                            {
                                cf = repo.Get<ContentFolder>(new ContentReference(cf.ParentLink.ID));
                            }

                            if( cf == null || cf.ParentLink == SiteDefinition.Current.GlobalAssetsRoot)
                            {
                                done = true;
                            }
                        } while (!done);
                    }
                    else
                    {
                        // Remove the folder data, it's invalid.
                        FolderData.Delete(foundData);
                    }
                }

                if (add)
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

        private IEnumerable<IContent> FindSharePointItems()
        {
            var contentTypeRepo = _serviceLocator.GetInstance<IContentTypeRepository>();

            var cmsTypes = contentTypeRepo.List().Where(ctr => ctr.ModelType != null && typeof(ISharePointItem).IsAssignableFrom(ctr.ModelType)).ToList();
            var refrences = new List<ContentReference>();
            foreach (var contentType in cmsTypes)
            {
                var items = _serviceLocator.GetInstance<IContentModelUsage>().ListContentOfContentType(contentType).Select(x => x.ContentLink.ToReferenceWithoutVersion()).Distinct().ToList();
                refrences.AddRange(items);
            }
            var dataResults = _serviceLocator.GetInstance<IContentRepository>()
                .GetItems(refrences, new LoaderOptions() { LanguageLoaderOption.FallbackWithMaster() })
                .Where(r => !string.IsNullOrEmpty((r as ISharePointItem).SharePointGUID) && r.IsDeleted == false);

            return dataResults;
        }

    }
}