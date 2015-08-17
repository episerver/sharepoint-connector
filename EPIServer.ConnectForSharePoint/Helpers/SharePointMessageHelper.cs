using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Configuration;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public class SharePointMessageHelper : ISharePointMessageHelper
    {
        private ILogger _logger;
        private ISharePointHelper _helper;
        private IContentProviderHelper _contentProviderHelper;
        private ISharePointSettingsHelper _settingsHelper;
        private IMessageEmitterFactory _emitterFactory;

        public SharePointMessageHelper()
        {
            _logger = LogManager.GetLogger();
            _helper = new SharePointHelper();
            _contentProviderHelper = new ContentProviderHelper();
            _settingsHelper = new SharePointSettingsHelper();
            _emitterFactory = new SharePointMessageEmitterFactory();
        }

        internal SharePointMessageHelper(ILogger logger, ISharePointHelper helper, IContentProviderHelper contentHelper, ISharePointSettingsHelper settings, IMessageEmitterFactory emitter)
        {
            _logger = logger;
            _helper = helper;
            _contentProviderHelper = contentHelper;
            _settingsHelper = settings;
            _emitterFactory = emitter;
        }

        /// <summary>
        /// Adds all the subfolders and subitems for the given parent folder
        /// </summary>
        /// <param name="parentFolder">The parent folder to use as selection criterea items and as a source for the cms folder name</param>
        /// <param name="selectedList">The list of items associated with all the selected views</param>
        /// <param name="selectedView">The selected views in the configuration</param>
        /// <param name="ParentFolderName">An optional parent folder name to use, i.e. in the case of views as a parent folder the name is actually different. </param>
        public void AddSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, string ParentFolderName = null)
        {
            var items = _helper.GetItems(selectedList, selectedView, parentFolder);
            var iterator = items.GetEnumerator();
            while (iterator.MoveNext())
            {
                var spItem = iterator.Current;
                spItem.entryPointId = -1; // we cant use the entrypointid, we have to use the parent folder name 
                if (string.IsNullOrEmpty(ParentFolderName))
                    spItem.ParentFolderName = parentFolder.Name;
                else
                    spItem.ParentFolderName = ParentFolderName;

                if (spItem.Folder != null)
                {
                    // Only add the folder if its not already there.
                    // Add the folder if its not already there
                    var contentRef = _contentProviderHelper.FindRecursiveFolders(spItem.Folder.Name, _settingsHelper.GetEntryPoint());
                    if (contentRef == null)
                    {
                        _logger.Information("Adding SharePointFolderAddMessage to the queue " + spItem.Title);
                        _emitterFactory.Create().Emit<SharePointFolderAddMessage>(new SharePointFolderAddMessage()
                        {
                            Item = spItem,
                        });
                    }

                    AddSubItems(spItem.Folder, selectedList, selectedView);
                }
                else if (spItem.File != null)
                {
                    // now emit the message
                    _logger.Information("Adding SharePointFolderAddMessage to the queue " + spItem.Title);
                    _emitterFactory.Create().Emit<SharePointContentAddMessage>(new SharePointContentAddMessage()
                    {
                        Item = spItem,
                    });
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
            var selectedLists = SharePointHelper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists)
            {
                // The idListAndView is an object that contains two quids, one for the list and one for the view.
                var idListAndView = SharePointHelper.GetListAndViewIds(item.Key);

                // Get the selected list using the first quid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second quid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                // Make sure that the view quid and the view quid is in the listids, if its not then it needs to be deleted.
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
                        entryPointId = SharePointSettings.Current.EntryPoint.ID,
                        Folder = new SPFolder()
                        {
                            Name = FolderName,
                        },
                    };

                    _logger.Information("Deleting SharePointFolderAddMessage to the queue " + tbd.Title);
                    SharePointMessageEmitterFactory fact2 = new SharePointMessageEmitterFactory();
                    fact2.Create().Emit<SharePointFolderDeleteMessage>(new SharePointFolderDeleteMessage()
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
            var selectedLists = SharePointHelper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists) // For each view selected
            {
                // The idListAndView is an object that contains two quids, one for the list and one for the view.
                var idListAndView = SharePointHelper.GetListAndViewIds(item.Key);

                // Get the selected list using the first quid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second quid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));
                var FolderName = selectedList.Title + " - " + selectedView.Title;
                var parentFolder = selectedList.RootFolder;

                // Add the folder if its not already there
                var helper = new ContentProviderHelper();
                var contentRef = helper.FindRecursiveFolders(FolderName, SharePointSettings.Current.EntryPoint);
                if (contentRef == null)
                {
                    SPListItem viewListItem = new SPListItem()
                    {
                        entryPointId = SharePointSettings.Current.EntryPoint.ID,
                        Title = selectedView.Title,
                        Folder = new SPFolder()
                        {
                            Name = FolderName,
                        }

                    };

                    // Add the view folder.
                    SharePointMessageEmitterFactory spfact = new SharePointMessageEmitterFactory();
                    spfact.Create().Emit<SharePointFolderAddMessage>(new SharePointFolderAddMessage()
                    {
                        Item = viewListItem,
                    });
                }

                AddSubItems(parentFolder, selectedList, selectedView, FolderName);
            }
        }
    }
}