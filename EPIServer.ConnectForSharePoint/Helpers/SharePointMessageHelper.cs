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
using EPiServer.DataAbstraction;
using System.Diagnostics.CodeAnalysis;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public class SharePointMessageHelper : ISharePointMessageHelper
    {
        private ILogger _logger;
        private ISharePointHelper _helper;
        private ISharePointProcessorHelper _processorHelper;
        private IContentProviderHelper _contentProviderHelper;
        private ISharePointSettingsHelper _settingsHelper;
        private IMessageEmitterFactory _emitterFactory;
        private IServiceLocator _serviceLocator;
        internal IList<SPListItem> _listSPFolders = new List<SPListItem>();
        internal IList<SPListItem> _ListSPItems = new List<SPListItem>();

        [ExcludeFromCodeCoverageAttribute]
        public SharePointMessageHelper()
        {
            _logger = LogManager.GetLogger();
            _helper = new SharePointHelper();
            _contentProviderHelper = new ContentProviderHelper();
            _settingsHelper = new SharePointSettingsHelper();
            _emitterFactory = new SharePointMessageEmitterFactory();
            _serviceLocator = ServiceLocator.Current;
            _processorHelper = new SharePointProcessorHelper();
        }

        internal SharePointMessageHelper(ILogger logger, ISharePointHelper helper, IContentProviderHelper contentHelper, ISharePointSettingsHelper settings, 
            IMessageEmitterFactory emitter, IServiceLocator serviceLocator)
        {
            _logger = logger;
            _helper = helper;
            _contentProviderHelper = contentHelper;
            _settingsHelper = settings;
            _emitterFactory = emitter;

            _serviceLocator = serviceLocator;
            _processorHelper = _serviceLocator.GetInstance<ISharePointProcessorHelper>();
        }

        /// <summary>
        /// Begins the processing of the selected lists, reading all the items in pages and firing insert messages to the message queue
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        public void BeginProcessing(string listids, IEnumerable<SPList> splists)
        {
            var selectedLists = _helper.GetSelectedListViewIds(splists);
            foreach (var list_item in selectedLists)
            {
                _listSPFolders.Clear(); // will contain just folders
                _ListSPItems.Clear();   // will contain just documents and items 

                // The idListAndView is an object that contains two guids, one for the list and one for the view.
                var idListAndView = _helper.GetListAndViewIds(list_item.Key);

                // split the list and view Guid and get the root folder items for this list
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Invalidate the metadata cache for this list, this ensures that if the metadata changes
                // we can fix it when the user hits save, refresh, or the scheduled task runs.
                _processorHelper.GetMetaData(selectedList, true);

                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                // create and/or update our root folder for this list/view
                SPListItem rootFolder = CreateParentSPListItem(selectedList, selectedView);
                var folderRef = AddOrUpdateFolder(rootFolder);
                _listSPFolders.Add(rootFolder);

                // now get all the items, paged, and process each item.
                string nextUri;
                processItemList(selectedList, rootFolder,
                    _helper.GetPagedItems(selectedList,
                    selectedView, rootFolder.Folder, out nextUri));

                // process any more batches, if any
                while (!String.IsNullOrEmpty(nextUri))
                {
                    processItemList(selectedList, rootFolder,
                        _helper.GetNextPagedItems(selectedList, nextUri, out nextUri));
                }

                // if the list folder already existed we need to check if anything has 
                // changed since the last processing loop
                // i.e. remove any user added folders or any folders removed from sharepoint.
                if (folderRef != null)
                {
                    RemoveUserAddedFolders(folderRef);
                    RemoveSharePointAddedFolders(folderRef, _listSPFolders);
                }
            }
        }

        /// <summary>
        /// Processes each item in the items list and fires the appropriate 
        /// CRUD message for each Share Point item type. 
        /// </summary>
        /// <param name="list">The Sharepoint list the items are related to</param>
        /// <param name="rootFolder">The root folder for the items</param>
        /// <param name="items">The items to process</param>
        internal void processItemList(SPList list, SPListItem rootFolder, IEnumerable<SPListItem> items)
        {
            // the folder service is used to load the parent folder.
            var folderService = SharePointHelper.SPServicesHub.SPFilesFoldersService;

            // iterate over all the items
            foreach (var item in items)
            {
                // is the item a folder?
                if (item.Folder != null)
                {
                    item.entryPointId = -1;
                    var parent = folderService.LoadParentFolder(item.Folder);

                    // if the parents name is the same as the list title
                    // this means its a root folder
                    if (parent.Name == list.Title)
                    {
                        item.ParentFolderSharePointGuid = rootFolder.GUID;
                    }
                    else
                    {
                        // We need to find the parent folder in our folder list.
                        // Note that the message queue might not have processed yet
                        // but thats ok, we just need the Guid of the parent item
                        var i = _listSPFolders.FirstOrDefault(x => x.Folder.ServerRelativeUrl == parent.ServerRelativeUrl);
                        item.ParentFolderSharePointGuid = i.GUID;
                    }

                    // Determins if we are sending the update or create message. 
                    AddOrUpdateFolder(item);

                    // Make sure the folder list item gets added to our internal list
                    // incase we need to find in the next batch.
                    _listSPFolders.Add(item);
                }
                else
                {
                    // All items in this sharepoint list get added to this list
                    _ListSPItems.Add(item);

                    // Todo: 
                    // Modify the splist item with the proper item data to insert it.
                    // The only way to figure out its parent folder is to 
                    // parse its relative url and find the sharepoint folder Guid
                    // in the _allsharepointfolders list. 
                }
            }
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
                                    ContentLink = item.ContentLink, // id of the folder the item is in
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
                                    ContentLink = item.ContentLink, // id of the folder the item is in
                                    GUID = new Guid(mediaItem.SharePointGUID),
                                },
                            });
                        }
                    }
                }
            }
        }

        /// <summary>
        /// If any subfolders of the specified folderRef exist, that are not in the sharepoint list
        /// we will fire the delete message.
        /// </summary>
        /// <param name="folderRef"></param>
        /// <param name="sharepointListItems"></param>
        private void RemoveSharePointAddedFolders(ContentReference folderRef, IEnumerable<SPListItem> sharepointListItems)
        {
            IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();
            FolderData fd = new FolderData();

            var cmsFolders = repo.GetChildren<ContentFolder>(folderRef);
            foreach (var cmsFolder in cmsFolders)
            {
                // Make sure we check any sub folders that have been removed. 
                RemoveSharePointAddedFolders(cmsFolder.ContentLink, sharepointListItems);

                Boolean delete = true;
                // get the folder from the big table
                fd = fd.GetFolder(cmsFolder.ContentLink.ID);
                if (fd.RefId != -1)
                {
                    foreach (var item in sharepointListItems)
                    {
                        if (item.Folder != null && item.GUID == fd.FolderGuid)
                        {
                            delete = false;
                            item.entryPointId = folderRef.ID;
                            break;
                        }
                    }

                    if (delete == true)
                    {
                        _emitterFactory.Create().Emit(new SharePointFolderDeleteMessage()
                        {
                            Item = new SPListItem()
                            {
                                entryPointId = fd.RefId, // id of the folder the item is in
                                GUID = fd.FolderGuid,
                                Folder = new SPFolder() { Name = cmsFolder.Name }
                            },
                        });
                    }
                }
            }
        }

        /// <summary>
        /// Adds or updates SPListItems that are not folders 
        /// </summary>
        /// <param name="spItem">The item to update or add</param>
        /// <param name="cmsItems">A list of all items currently in cms</param>
        internal void AddOrUpdateItem(SPListItem spItem, IEnumerable<IContent> cmsItems)
        {
            if (spItem.Folder == null)
            {
                foreach (var cmsItem in cmsItems) // Search for the item in our IContent List
                {
                    if (cmsItem is ISharePointItem)
                    {
                        var thisItem = cmsItem as ISharePointItem;
                        if (spItem.GUID.ToString().Equals(thisItem.SharePointGUID))
                        {
                            spItem.ContentLink = cmsItem.ContentLink;
                        }
                    }
                }

                if (spItem.File != null) // is it a document?
                {   
                    // if there is reference send Add message
                    if (spItem.ContentLink == null)
                    {
                        _logger.Information("Adding SharePointContentAddMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointContentAddMessage() { Item = spItem });
                    }
                    else // its an update
                    {
                        _logger.Information("Adding SharePointContentUpdateMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointContentUpdateMessage() { Item = spItem });
                    }
                }
                else // it is a list item
                {
                    // if there is reference send Add message
                    if (spItem.ContentLink == null)
                    {
                        _logger.Information("Adding SharePointListItemAddMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointListItemAddMessage() { Item = spItem });
                    }
                    else // its an update
                    {
                        _logger.Information("Adding SharePointListItemUpdateMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit(new SharePointListItemUpdateMessage() { Item = spItem });
                    }
                }
            }
        }

        /// <summary>
        /// Deletes all folders in the specified splists collection, UNLESS there list ids to exclude from deletion.
        /// </summary>
        /// <param name="splists"></param>
        /// <param name="listids"></param>
        public void QueueRootFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null)
        {
            var selectedLists = _helper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists)
            {
                bool delete = true;

                // The idListAndView is an object that contains two guids, one for the list and one for the view.
                var idListAndView = _helper.GetListAndViewIds(item.Key);

                // Get the selected list using the first guid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second guid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                if (listids != null) // we are some folders
                {
                    foreach (var id in listids)
                    {
                        var listIds = id.Split('~');
                        if (selectedList.Id == Guid.Parse(listIds[0]))
                        {
                            // We found a match, this means the view is in the current selection - don't delete it!  
                            // else all the items in the folder will be deleted and existing pages will lose content.
                            delete = false;
                            break;
                        }
                    }
                }

                if ( delete )
                {
                    // The message on the queue will be SPListItem
                    // Because we have only a list (like "documents") and a view (like "all documents") we need to create a SPListItem
                    // that contains some aspects of both.
                    _emitterFactory.Create().Emit(new SharePointFolderDeleteMessage()
                    {
                        Item = CreateParentSPListItem(selectedList, selectedView)
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
        /// 
        /// </summary>
        /// <param name="item"></param>
        /// <param name="entryPoint">The current entry point that is saved in the settings.</param>
        /// <returns>the content reference of the folder if found</returns>
        private ContentReference AddOrUpdateFolder(SPListItem item)
        {
            var add = true;
            ContentReference contentRef = null;

            // before we add it, lets see if the item is in the SPFolderMappings dynamic data store, if it is 
            // we can assume the name changed and we just need to update it
            var fd = new FolderData();
            var foundData = fd.GetFolder(item.GUID);

            if (foundData.RefId != -1)
            {
                var repo = _serviceLocator.GetInstance<IContentRepository>();
                ContentFolder cf;
                if( repo.TryGet<ContentFolder>(new ContentReference(foundData.RefId), out cf) )
                {
                    _logger.Information("SharePointFolderUpdateMessage to the queue : " + item.Folder.Name);
                    _emitterFactory.Create().Emit(new SharePointFolderUpdateMessage()
                    {
                        Item = item,
                    });
                    add = false;
                    contentRef = cf.ContentLink; // Make sure we return the content reference
                }
                else
                {
                    // Remove the folder data, it's invalid., we will emit the create folder message.
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