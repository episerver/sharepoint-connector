using System;
using System.Collections.Generic;
using System.Linq;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.ConnectForSharePoint.Rest;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public class SharepointToCmsUtilities : ISharepointToCmsUtilities
    {
        private IServiceLocator _serviceLocator;

        public SharepointToCmsUtilities(IServiceLocator locator)
        {
            _serviceLocator = locator;
        }

        /// <inheritdoc/>
        public IEnumerable<IContent> FindSharePointItems()
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

        /// <inheritdoc/>
        public IEnumerable<SPList> FindSupportedLists()
        {
            var contextService = SharePointSettings.Current.IsModernAuth ?
            new EndpointContextService
            {
                EndpointContext = new SPRestClient(SharePointSettings.Current)
            } :
            new EndpointContextService
            {
                EndpointContext = new SPRestClient(SharePointSettings.Current, SharePointHelper.GetLatestCookies)
            };

            var listViewService = new ListsViewsService { EndpointContextService = contextService };

            // Get all of the lists we support in sharepoint.
            var documentLists = listViewService.GetLists((int)ListBaseType.DocumentLibrary);
            var customLists = listViewService.GetLists((int)ListBaseType.GenericList);

            return documentLists.Concat(customLists);
        }

        /// <inheritdoc/>
        public void RemoveUserAddedFolders(IMessageEmitterFactory emmiter, ContentReference folderContentReference)
        {
            IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();

            var cmsFolderItems = repo.GetChildren<ContentFolder>(folderContentReference);
            FolderData fd = new FolderData();
            foreach (var folder in cmsFolderItems)
            {
                fd = fd.GetFolder(folder.ContentLink.ID);
                if (fd.RefId == -1)
                {
                    emmiter.Emit<SharePointFolderDeleteMessage>( new SPListItem()
                    {
                        entryPointId = folderContentReference.ID, // id of the folder the item is in
                        Folder = new SPFolder() { Name = folder.Name }
                    });
                }
            }
        }

        /// <inheritdoc/>
        public void RemoveSharePointAddedFolders(IMessageEmitterFactory emmiter, ContentReference folderRef, IEnumerable<SPListItem> sharepointListItems)
        {
            IContentRepository repo = _serviceLocator.GetInstance<IContentRepository>();
            FolderData fd = new FolderData();

            var cmsFolders = repo.GetChildren<ContentFolder>(folderRef);
            foreach (var cmsFolder in cmsFolders)
            {
                // Make sure we check any sub folders that have been removed. 
                RemoveSharePointAddedFolders(emmiter, cmsFolder.ContentLink, sharepointListItems);

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
                        emmiter.Emit<SharePointFolderDeleteMessage>(new SPListItem()
                        {
                            entryPointId = fd.RefId, // id of the folder the item is in
                            GUID = fd.FolderGuid,
                            Folder = new SPFolder() { Name = cmsFolder.Name }
                        });
                    }
                }
            }
        }

        /// <inheritdoc/>
        public void AddOrUpdateItem(IMessageEmitterFactory emmiter, SPListItem spItem, IEnumerable<IContent> cmsItems)
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
                        emmiter.Emit<SharePointContentAddMessage>(spItem);
                    }
                    else // its an update
                    {
                        emmiter.Emit<SharePointContentUpdateMessage>(spItem);
                    }
                }
                else // it is a list item
                {
                    // if there is reference send Add message
                    if (spItem.ContentLink == null)
                    {
                        emmiter.Emit<SharePointListItemAddMessage>(spItem);
                    }
                    else // its an update
                    {
                        emmiter.Emit<SharePointListItemUpdateMessage>(spItem);
                    }
                }
            }
        }

        /// <inheritdoc/>
        public ContentReference AddOrUpdateFolder(IMessageEmitterFactory emmiter, SPListItem item)
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
                if (repo.TryGet<ContentFolder>(new ContentReference(foundData.RefId), out cf))
                {
                    emmiter.Emit<SharePointFolderUpdateMessage>(item);
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
                emmiter.Emit<SharePointFolderAddMessage>(item);
            }
            return contentRef;
        }

        /// <inheritdoc/>
        public void DeleteDocumentsAndItems(IMessageEmitterFactory emmiter, IEnumerable<IContent> cmsListItems, IEnumerable<SPListItem> sharepointListItems)
        {
            foreach (var item in cmsListItems)
            {
                var mediaItem = item as ISharePointItem;
                if (item != null &&
                    sharepointListItems.FirstOrDefault<SPListItem>(i => i.GUID == new Guid(mediaItem.SharePointGUID)) == null)
                {
                    // if its a list item, delete it
                    if (mediaItem.Type == SharePointItemType.ListItem)
                    {
                        emmiter.Emit<SharePointListItemDeleteMessage>(new SPListItem()
                            {
                                File = new SPFile() { Name = item.Name },
                                ContentLink = item.ContentLink, // id of the folder the item is in
                                GUID = new Guid(mediaItem.SharePointGUID),
                            });
                    }
                    else if (mediaItem.Type == SharePointItemType.LibraryItem)
                    {
                        emmiter.Emit<SharePointContentDeleteMessage>( new SPListItem()
                        {
                            File = new SPFile() { Name = item.Name },
                            ContentLink = item.ContentLink, // id of the folder the item is in
                            GUID = new Guid(mediaItem.SharePointGUID)
                        });
                    }
                }
            }
        }

        public IEnumerable<KeyValuePair<string, string>> GetListViewPairs(IEnumerable<SPList> lists)
        {
            var idPairs = lists.Where(list => Enum.IsDefined(typeof(SupportedBaseTemplates), list.BaseTemplate)).
                SelectMany<SPList, KeyValuePair<string, string>>(spl =>
                {
                    return spl.Views.Where(view => !view.Hidden).Select<SPListView, KeyValuePair<string, string>>(spv =>
                    {
                        return new KeyValuePair<string, string>(String.Format("{0}~{1}", spl.Id, spv.Id),
                            String.Format("{0} - {1}", spl.Title, spv.Title));
                    });
                });

            return idPairs;
        }

        public IEnumerable<KeyValuePair<string, string>> GetSelectedListViewIds(IEnumerable<SPList> lists)
        {
            var listIds = SharePointSettings.Current.ListIds.Split(new string[] { ";" }, StringSplitOptions.None);
            var idListViewPairs = GetListViewPairs(lists);
            return idListViewPairs.Where(li => listIds.Contains(li.Key));
        }

        /// <inheritdoc/>
        public void QueueRootFoldersForDelete(IMessageEmitterFactory emmiter, IEnumerable<SPList> splists, IEnumerable<string> listids = null)
        {
            var selectedLists = GetSelectedListViewIds(splists);
            foreach (var item in selectedLists)
            {
                bool delete = true;

                // The idListAndView is an object that contains two guids, one for the list and one for the view.
                var idListAndView = item.Key.Split('~');

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

                if (delete)
                {
                    // The message on the queue will be SPListItem
                    // Because we have only a list (like "documents") and a view (like "all documents") we need to create a SPListItem
                    // that contains some aspects of both.
                    emmiter.Emit<SharePointFolderDeleteMessage>(CreateParentSPListItem(selectedList, selectedView));
                }
            }
        }

        /// <inheritdoc/>
        public SPListItem CreateParentSPListItem(SPList list, SPListView view)
        {
            var folderName = list.Title + " - " + view.Title;

            // Dummy SPListItem used for message to processor
            SPListItem item = new SPListItem()
            {
                GUID = list.Id,
                entryPointId = SharePointSettings.Current.EntryPoint.ID,
                Title = view.Title,
                Folder = new SPFolder()
                {
                    Name = folderName
                }

            };

            return item;
        }
    }
}