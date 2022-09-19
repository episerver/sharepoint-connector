using System;
using System.Collections.Generic;
using System.Linq;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using System.Diagnostics.CodeAnalysis;
using EPiServer.ConnectForSharePoint.SPServices;
using System.Threading;
using EPiServer.ConnectForSharePoint.Helpers;

namespace EPiServer.ConnectForSharePoint.Processor
{
    [ServiceConfiguration(ServiceType = typeof(IMetaDataProcessor), Lifecycle = ServiceInstanceScope.Singleton)]
    public class MetaDataProcessor : IMetaDataProcessor
    {
        private ILogger _logger;
        private ISharePointHelper _helper;
        private IMessageEmitterFactory _emitterFactory;
        private IServiceLocator _serviceLocator;

        private ISharepointToCmsUtilities SharepointToCmsUtils { get; set; }

        static internal Thread _processingThread;
        static internal string _listids;
        static internal IEnumerable<SPList> _splists;
        static Object _lock = new Object();
        static public Boolean _processingResult = true;
        static public Exception _processingException;

        // TODO : The following variables are used by the unit tests to verify behavior, needs to be examined and possible refactored out.
        internal bool _clearAllSPItemsWhenDone = true;
        internal IList<SPListItem> _SPFolders = new List<SPListItem>();
        internal IList<SPListItem> _SPItems = new List<SPListItem>();
        internal List<SPListItem> _AllSPItems = new List<SPListItem>();
        internal IEnumerable<IContent> _SPItemsInCMs;

        internal bool _stopSignaled;

        private bool _isTest;

        [ExcludeFromCodeCoverageAttribute]
        public MetaDataProcessor()
        {
            _logger = LogManager.GetLogger();
            _helper = new SharePointHelper();
            _emitterFactory = new SharePointMessageEmitterFactory();
            _serviceLocator = ServiceLocator.Current;

            SharepointToCmsUtils = new SharepointToCmsUtilities(ServiceLocator.Current);
        }

        internal MetaDataProcessor(ILogger logger, ISharePointHelper helper, IMessageEmitterFactory emitter, IServiceLocator serviceLocator)
        {
            _logger = logger;
            _helper = helper;
            _emitterFactory = emitter;

            _serviceLocator = serviceLocator;
            _isTest = true;

            SharepointToCmsUtils = new SharepointToCmsUtilities(_serviceLocator);
        }

        /// <summary>
        /// Starts the processing thread if its not already running
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        public void BeginProcessing(string listids, IEnumerable<SPList> splists)
        {
            if (_processingThread == null || !_processingThread.IsAlive)
            {
                _listids = listids;
                _splists = splists;
                _processingThread = new Thread(new ThreadStart(Process));
                _processingResult = true;
                _processingException = null;
                _stopSignaled = false;
                _processingThread.Start();
            }
        }

        public bool IsRunning()
        {
            bool retval = true;
            lock (_lock)
            {
                var emmiter = new SharePointMessageEmitterFactory();
                if (_processingThread == null || (!_processingThread.IsAlive && emmiter.QueueCount == 0))
                {
                    retval = false;
                }
            }
            return retval;
        }

        public void Stop()
        {
            _stopSignaled = true;

            var emitter = new SharePointMessageEmitterFactory();
            while ( IsRunning() )
            {
                emitter.Clear();
                Thread.Sleep(250);
            }
        }

        /// <summary>
        /// Begins the processing of the selected lists, reading all the items in pages and firing insert messages to the message queue
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        internal void Process()
        {
            try
            {
                var selectedLists = _helper.GetSelectedListViewIds(_splists);
                _SPItemsInCMs = SharepointToCmsUtils.FindSharePointItems();

                _logger.Debug("Sharepoint: Processing start. Current Items : " + _SPItemsInCMs.Count() +" Current folders: " + (new FolderData()).GetFolders().Count() );

                _AllSPItems.Clear();

                foreach (var list_item in selectedLists)
                {
                    if (_stopSignaled)
                    {
                        break;
                    }
                    _SPFolders.Clear(); // will contain just folders
                    _SPItems.Clear();   // will contain just documents and items 

                    // The idListAndView is an object that contains two guids, one for the list and one for the view.
                    var idListAndView = _helper.GetListAndViewIds(list_item.Key);

                    // split the list and view Guid and get the root folder items for this list
                    var selectedList = _splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                    // Invalidate the metadata cache for this list, this ensures that if the metadata changes
                    // we can fix it when the user hits save, refresh, or the scheduled task runs.
                    var processorHelper = _isTest ? _serviceLocator.GetInstance<ISharePointProcessorHelper>() : new SharePointProcessorHelper();

                    var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                    // create and/or update our root folder for this list/view
                    SPListItem rootFolder = SharepointToCmsUtils.CreateParentSPListItem(selectedList, selectedView);
                    var folderRef = SharepointToCmsUtils.AddOrUpdateFolder(_emitterFactory, rootFolder);
                    _SPFolders.Add(rootFolder);

                    _logger.Debug("Sharepoint: Processing SharePoint list: " + selectedList.Title + " ItemCount(including folders): " + selectedList.ItemCount);

                    // now get all the items, paged, and process each item.
                    string nextUri;
                    ProcessItemList(selectedList, selectedView, rootFolder,
                        _helper.GetPagedItems(selectedList,
                        selectedView, rootFolder.Folder, out nextUri), selectedView.Scope);

                    // process any more batches, if any
                    while (!String.IsNullOrEmpty(nextUri) && !_stopSignaled)
                    {
                        _logger.Debug("Sharepoint: " + _SPItems.Count() + " folders : " + _SPFolders.Count());
                        ProcessItemList(selectedList, selectedView, rootFolder,
                            _helper.GetNextPagedItems(selectedList, nextUri, out nextUri), selectedView.Scope);
                    }
                    _logger.Debug("Sharepoint: " + _SPItems.Count() + " folders : " + _SPFolders.Count());

                    // if the list folder already existed we need to check if anything has 
                    // changed since the last processing loop
                    // i.e. remove any user added folders or any folders removed from sharepoint.
                    if (folderRef != null && !_stopSignaled)
                    {
                        SharepointToCmsUtils.RemoveUserAddedFolders(_emitterFactory, folderRef);
                        SharepointToCmsUtils.RemoveSharePointAddedFolders(_emitterFactory, folderRef, _SPFolders);
                    }

                    _AllSPItems.AddRange(_SPItems);
                }

                // If we are stopped in the middle of processing dont call this method else we will delete all items that 
                // may have been there previously.
                if (!_stopSignaled)
                {
                    SharepointToCmsUtils.DeleteDocumentsAndItems(_emitterFactory, _SPItemsInCMs, _AllSPItems);
                }
            }
            catch ( Exception e)
            {
                // Unexpected error condition ending processing thread too soon.
                _logger.Error("Sharepoint: Error running update thread.", e);
                _processingResult = false;
                _processingException = e;
            }

            _logger.Debug( "Sharepoint: Processing Complete. Items in Queue:" + _emitterFactory.QueueCount);
            if( _clearAllSPItemsWhenDone )
            {
                _AllSPItems.Clear();
            }
        }

        /// <summary>
        /// Processes each item in the items list and fires the appropriate 
        /// CRUD message for each Share Point item type. 
        /// </summary>
        /// <param name="list">The Sharepoint list the items are related to</param>
        /// <param name="rootFolder">The root folder for the items</param>
        /// <param name="items">The items to process</param>
        internal void ProcessItemList(SPList list, SPListView view, SPListItem rootFolder, IEnumerable<SPListItem> items, int scope)
        {
            _logger.Debug("Sharepoint: Processing batch: " + items.Count());
            var fileFolderService = _serviceLocator.GetInstance<IFilesFoldersService>();

            // iterate over all the items
            foreach (var item in items)
            {
                if (_stopSignaled) break;

                item.entryPointId = -1;
                item.ParentFolderSharePointGuid = rootFolder.GUID;  // this is the root folder of the view.
                item.selectedView = view;

                // is the item a folder?
                if (item.Folder != null && scope == 0) // if scope is 0 we are adding folders else its a flat list.
                {
                    var parent = fileFolderService.LoadParentFolder(item.Folder);

                    // We need to find the parent folder in our folder list.
                    // Note that the message queue might not have processed yet
                    // but thats ok, we just need the Guid of the parent item
                    var i = _SPFolders.FirstOrDefault(x => x.Folder.ServerRelativeUrl == parent.ServerRelativeUrl);
                    if (i != null)
                    {
                        item.ParentFolderSharePointGuid = i.GUID;
                    }

                    // Determins if we are sending the update or create message. 
                    SharepointToCmsUtils.AddOrUpdateFolder(_emitterFactory, item);

                    // Make sure the folder list item gets added to our internal list
                    // incase we need to find in the next batch.
                    _SPFolders.Add(item);
                }
                else if (item.File != null)
                {
                    // Check to see if its a different parent folder by parsing out the parent folder url and look for it in the 
                    // folder list.
                    string url = item.File.ServerRelativeUrl.Remove(item.File.ServerRelativeUrl.IndexOf(item.File.Name) - 1);
                    var i = _SPFolders.FirstOrDefault(x => x.Folder.ServerRelativeUrl == url);
                    if (i != null)
                    {
                        item.ParentFolderSharePointGuid = i.GUID;
                    }

                    // Determins if we are sending the update or create message. 
                    SharepointToCmsUtils.AddOrUpdateItem(_emitterFactory, item, _SPItemsInCMs);

                    // All items in this sharepoint list get added to this list
                    _SPItems.Add(item);
                }
                else
                {
                    // Determins if we are sending the update or create message. 
                    SharepointToCmsUtils.AddOrUpdateItem(_emitterFactory, item, _SPItemsInCMs);

                    // All items in this sharepoint list get added to this list
                    _SPItems.Add(item);
                }
            }
        }
    }
}