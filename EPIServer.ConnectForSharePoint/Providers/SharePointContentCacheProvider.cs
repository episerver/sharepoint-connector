using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Caching;

namespace EPiServer.ConnectForSharePoint.Providers
{
    [ServiceConfiguration(ServiceType = typeof(ISharePointContentCacheProvider))]
    public class SharePointContentCacheProvider : ISharePointContentCacheProvider
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        private static readonly object _lockObject = new object();
        private static readonly string _dataMappingCacheKey = string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, "DataMapping");
        private static readonly string _supportMediaListCacheKey = string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, "SupportedMediaLists");
        private MemoryCache _cache = new MemoryCache(Consts.SHAREPOINT_CONTENT_CACHE_NAME);

        private static SharePointSettings _currentSettings
        {
            get { return SharePointSettings.Current; }
        }

        /// <summary>
        /// Return entry point from settings
        /// </summary>
        private static ContentReference EntryPoint
        {
            get
            {
                return _currentSettings.EntryPoint;
            }
        }

        public IEnumerable<SPList> LatestMediaLists
        {
            get
            {
                return GetSupportedMediaLists(true);
            }
        }

        public IEnumerable<SPList> CachedMediaLists
        {
            get
            {
                return GetSupportedMediaLists(false);
            }
        }

        public Dictionary<string, List<IModelWrapper>> Files
        {
            get
            {
                Dictionary<string, List<IModelWrapper>> copiedFiles = new Dictionary<string, List<IModelWrapper>>();
                lock (_lockObject)
                { 
                    copiedFiles = new Dictionary<string, List<IModelWrapper>>(GetDataMapping().Files);
                }
                return copiedFiles;
            }
        }

        public Dictionary<string, List<IModelWrapper>> Folders
        {
            get
            {
                Dictionary<string, List<IModelWrapper>> copiedFolders = new Dictionary<string, List<IModelWrapper>>();
                lock (_lockObject)
                { 
                    copiedFolders = new Dictionary<string, List<IModelWrapper>>(GetDataMapping().Folders);
                }
                return copiedFolders;
            }
        }

        #region public methods

        /// <summary>
        /// Builds the mapping of SharePoint files/folders and adds to cache.
        /// </summary>
        public void BuildMapping(ContentReference parentLink)
        {
            _logger.Information("Build mapping for content link: {0}.", parentLink);

            lock (_lockObject)
            {
                Dictionary<string, List<IModelWrapper>> files;
                Dictionary<string, List<IModelWrapper>> folders;

                GetFilesAndFoldersBelongToContent(parentLink, out files, out folders);

                AddToMapping(files, folders);
            }
        }

        /// <summary>
        /// Rebuilds the mapping of SharePoint files/folders and adds to cache.
        /// </summary>
        public void RebuildMapping()
        {
            _logger.Information("Remove all cached Files/Folders mapping.");

            ResetCache();

            CacheManager.Remove(_dataMappingCacheKey);
            BuildMapping(EntryPoint);
        }

        /// <summary>
        /// Removes a file from cache.
        /// </summary>
        /// <param name="contentLink">The parent link.</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file.</param>
        public void RemoveFileFromCache(ContentReference contentLink, string serverRelativeUrl)
        {
            _logger.Information("Remove file from cache, contentLink: {0} - serverRelativeUrl: {1}.", contentLink, serverRelativeUrl);

            lock (_lockObject)
            {
                foreach (string keyFolderParent in Files.Keys)
                {
                    // Note: find the deleted file from its folders and remove it out of the cache, but there are cases in Edit mode
                    // the deleted item is still be referenced by other object(s) e.g. blob then it still is listed on the folder tree, this is WON't fix.
                    // Example: in Edit mode, double click to open an item, go to SharePoint delete the associate file, the below if condition will
                    // at least be excuted suceefully for time(s) up to the folder(s) contains the file, but because of still be referenced the item
                    // is there until cache invalidated.
                    var fileWarppers = Files[keyFolderParent];
                    if (fileWarppers != null && fileWarppers.Any(item => item.ListItem.File.ServerRelativeUrl.Equals(serverRelativeUrl, StringComparison.OrdinalIgnoreCase)))
                    {
                        _logger.Information("Found file item '{0}' in cache with its folder cache key: {1}.", serverRelativeUrl, keyFolderParent);
                        fileWarppers.RemoveAll(item => item.ListItem.File.ServerRelativeUrl.Equals(serverRelativeUrl, StringComparison.OrdinalIgnoreCase));

                        // re-add to cache after removing file
                        var newFiles = new Dictionary<string, List<IModelWrapper>>();
                        newFiles.Add(keyFolderParent, fileWarppers.ToList());

                        var newFolders = new Dictionary<string, List<IModelWrapper>>();
                        newFolders.Add(keyFolderParent, Folders.ContainsKey(keyFolderParent) ?
                            Folders[keyFolderParent] : new List<IModelWrapper>());

                        AddToCache(keyFolderParent, new DataMapping()
                        {
                            Files = newFiles,
                            Folders = newFolders
                        });
                    }
                }

                RemovePageCommonCache(contentLink);
            }
        }

        #endregion

        /// <summary>
        /// Adds files/folders to mapping.
        /// </summary>
        /// <param name="files">The files.</param>
        /// <param name="folders">The folders.</param>
        private void AddToMapping(Dictionary<string, List<IModelWrapper>> files, Dictionary<string, List<IModelWrapper>> folders)
        {
            var sourceFiles = GetDataMapping().Files;
            foreach (var file in files)
            {
                // Remove from Files that files does not contain
                if (file.Value.Count < 1)
                {
                    sourceFiles.Remove(file.Key);
                }
                else
                {
                    if (sourceFiles.ContainsKey(file.Key))
                    {
                        sourceFiles[file.Key] = files[file.Key];
                    }
                    else
                    {
                        sourceFiles.Add(file.Key, files[file.Key]);
                    }
                }
            }

            var sourceFolders = GetDataMapping().Folders;
            foreach (var folder in folders)
            {
                if (sourceFolders.ContainsKey(folder.Key))
                {
                    sourceFolders[folder.Key] = folders[folder.Key];
                }
                else
                {
                    sourceFolders.Add(folder.Key, folders[folder.Key]);
                }
            }
        }

        private void GetFilesAndFoldersBelongToContent(ContentReference parentLink, out Dictionary<string, List<IModelWrapper>> fileDictionary, out Dictionary<string, List<IModelWrapper>> folderDictionary, bool ignoreCache = false)
        {
            _logger.Information("GetFilesAndFoldersBelongToContent - parentLink: {0}.", parentLink);

            var key = parentLink.ToString();
            if (!ignoreCache)
            {
                var cacheData = GetCachedItem(key) as DataMapping;
                if (cacheData != null)
                {
                    _logger.Information("Found mapping in cache, return.");

                    fileDictionary = cacheData.Files;
                    folderDictionary = cacheData.Folders;

                    return;
                }
            }

            fileDictionary = new Dictionary<string, List<IModelWrapper>>();
            folderDictionary = new Dictionary<string, List<IModelWrapper>>();

            _logger.Information("Not found mapping in cache, try to get items from SharePoint.");

            if (parentLink.Equals(EntryPoint))
            {
                var listCollection = GetSupportedMediaLists(true);
                var selectedListViewIds = SharePointHelper.GetSelectedListViewIds(listCollection);
                string[] idListAndView;

                foreach (var item in selectedListViewIds)
                {
                    // use lock to make sure that cache and view query is kept seperately
                    lock (_lockObject)
                    {
                        idListAndView = SharePointHelper.GetListAndViewIds(item.Key);
                        if (idListAndView == null)
                        {
                            continue;
                        }

                        var selectedList = listCollection.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));
                        if (selectedList == null)
                        {
                            continue;
                        }

                        var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                        AddRootFilesAndFolders(selectedList, selectedView, ref fileDictionary, ref folderDictionary);
                    }
                }
            }
            else
            {
                // Find folder modelWrapper conresponds to parentLink and build its files
                foreach (var item in Folders)
                {
                    var folder = item.Value.FirstOrDefault(v => v.ContentLink.Equals(parentLink));
                    if (folder != null && folder.ListItem != null)
                    {
                        // The View maybe has been outdate since last loaded
                        SyncWithCachedViewQuery(folder.View);
                        AddItems(folder.List, folder.View, folder.ListItem.Folder, parentLink.ToString(), ref fileDictionary, ref folderDictionary);
                        break;
                    }
                }

                if (!Folders.ContainsKey(key))
                {
                    return;
                }
                foreach (var folderItem in Folders[key])
                {
                    // The View maybe has been outdate since last loaded
                    SyncWithCachedViewQuery(folderItem.View);
                    AddItems(folderItem.List, folderItem.View, folderItem.ListItem.Folder, folderItem.ContentLink.ToString(), ref fileDictionary, ref folderDictionary);
                }
            }

            _logger.Information("Add mapping into cache for contentLink: {0}.", parentLink);
            AddToCache(key, new DataMapping() { Files = fileDictionary, Folders = folderDictionary });
        }

        private void AddRootFilesAndFolders(SPList selectedList, SPListView selectedView, ref Dictionary<string, List<IModelWrapper>> fileDictionary, ref Dictionary<string, List<IModelWrapper>> folderDictionary)
        {
            var parentFolder = selectedList.RootFolder;

            // Follow the folder structure, RootFolder is first
            var contentLink = AddRootFolder(selectedList, selectedView, folderDictionary);
            // Next are ROOT's files and sub folders recursively one level
            AddItems(selectedList, selectedView, parentFolder, contentLink.ToString(), ref fileDictionary, ref folderDictionary);

            string viewCacheKey;    // = GetViewCacheKey(selectedView); // string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, selectedView.Id);
            var viewQueryCacheKey = GetViewQueryCacheKey(selectedView, out viewCacheKey); // viewCacheKey + "_ViewQuery";            
            var lastTimeLoaded = GetLastTimeLoaded(viewCacheKey);   // If there was NO cache of prior time loading then DateTime.MinValue 

            if ((lastTimeLoaded > DateTime.MinValue))
            {
                bool queryChanged = IsViewQueryChanged(selectedView, viewQueryCacheKey);

                // Store the time that view has changed for sub levels know
                if (queryChanged)
                {
                    AddToCache(GetViewQueryTimeChangedCacheKey(viewCacheKey), DateTime.Now, GetAbsExpirePolicy());
                }

                if (selectedList.LastItemDeletedDate > lastTimeLoaded ||
                selectedList.LastItemModifiedDate > lastTimeLoaded || queryChanged)
                {
                    RemoveChildrenCache(contentLink);
                }
            }

            // Store this view loaded time (root Folder of the List) equals to the first level under the EntryPoint
            AddToCache(viewCacheKey, DateTime.Now, GetAbsExpirePolicy());

            // Store ths view's query, because there is no way to know the time this view's query has changed
            AddToCache(viewQueryCacheKey, selectedView.ListViewXml, GetAbsExpirePolicy());
        }

        private ContentReference AddRootFolder(SPList selectedList, SPListView selectedView, Dictionary<string, List<IModelWrapper>> folderDictionary)
        {
            // Create model wrapper and get/set content link for it
            var modelWrapper = new ModelWrapper(selectedList, selectedView, null);

            AddModelWrapper(EntryPoint.ToString(), modelWrapper, ref folderDictionary);

            RemovePageCommonCache(modelWrapper.ContentLink);

            return modelWrapper.ContentLink;
        }

        private void AddItems(SPList selectedList, SPListView selectedView, SPFolder parentFolder, string parentContentLink, ref Dictionary<string, List<IModelWrapper>> fileDictionary, ref Dictionary<string, List<IModelWrapper>> folderDictionary)
        {
            // Could not be here because the AddToCache called after, see AddRootFilesAndFolders
            //// The View maybe has been outdate since last loaded
            //if (!parentFolder.IsRootFolder)
            //{
            //    SyncCachedViewXml(selectedView);
            //}

            // Get items in a folder filter by view
            var items = SharePointHelper.SPServicesHub.SPListsViewsService.GetItems(selectedList, selectedView, parentFolder);
            var iterator = items.GetEnumerator();

            // Assign an empty list to this parentContentLink for later (AddToMapping) 
            // knowing that which files/folder should be removed from Files/Folders
            fileDictionary.Add(parentContentLink, new List<IModelWrapper>());
            folderDictionary.Add(parentContentLink, new List<IModelWrapper>());

            while (iterator.MoveNext())
            {
                var item = iterator.Current;
                if (item.Folder != null)
                {
                    #region Video item in SharePoint365 Asset list
                    // If folder was created automattically after uploading video files to SharePoint site
                    if (item.VideoSetShowDownloadLink.HasValue || item.VideoSetShowEmbedLink.HasValue)
                    {
                        // When we upload a video file to SharePoint, it will create a folder structure for that video,
                        // ex: if we upload "Alloy Track Video.mp4" to SharePoint, the folder structure will be created as below (it's invisible):
                        //      - Alloy Track Video
                        //          Alloy Track Video.mp4
                        //          - Additional Content
                        //          - Preview Images
                        //              Alloy Track Video_mp4.png
                        // We don't want to show that structure in our add-on, we would like to show folder structure like it displayed on SharePoint UI,
                        // so we hide auto generated folder and add the video file into its actual parent node.
                        // NOTE: - you couldn't upload the video files have same name (with different extensions like Video.mp4 and Video.wmv) to the same SharePoint list's folder.
                        //       - you couldn't create new folder has same name as the invisible one e.g. the above "Alloy Track Video" folder
                        // more info about asset library: http://technet.microsoft.com/en-us/library/ee414276(v=office.15).aspx

                        // Fix bug#121536 - ListsViewsService.GetItems returns nothing from the folder (Alloy Track Video) 
                        // because items in this folder is NOT "Video" type anymore
                        var videoFiles = SharePointHelper.SPServicesHub.SPFilesFoldersService.GetFiles(item.Folder);
                        // With the NOTE above, we expect there is one actual video file in the folder and show only the corresponding video item
                        var videoFileItem = videoFiles.FirstOrDefault(f => f.ListItem != null);
                        if (videoFileItem != null)
                        {
                            AddFile(videoFileItem.ListItem, selectedList, selectedView, parentContentLink, ref fileDictionary);
                        }
                    }
                    #endregion
                    // Normal folder
                    else
                    {
                        // Set reference for later use
                        item.Folder.ParentFolder = parentFolder;
                        // Add folder
                        AddFolder(item, selectedList, selectedView, parentContentLink, ref folderDictionary);
                    }
                }
                else if (item.File != null)
                {
                    // Add file
                    AddFile(item, selectedList, selectedView, parentContentLink, ref fileDictionary);
                }
            }

            var timeLoadedCacheKey = string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, parentContentLink);
            var lastTimeLoaded = GetLastTimeLoaded(timeLoadedCacheKey);

            // Need to remove CMS content cache for this folder' children after the view query changed
            if (GetLastTimeLoaded(string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, selectedView.Id) + "_ViewQueryChangedTime") > lastTimeLoaded)
            {
                RemoveChildrenCache(ContentReference.Parse(parentContentLink));
            }

            // Store loaded time for this level (in the folder tree)
            AddToCache(timeLoadedCacheKey, DateTime.Now, GetAbsExpirePolicy());
        }

        private void AddFolder(SPListItem itemFolder, SPList selectedList, SPListView selectedView, string parentContentLink, ref Dictionary<string, List<IModelWrapper>> folderDictionary)
        {
            // Create model wrapper and get/set content link for it
            var modelWrapper = new ModelWrapper(selectedList, selectedView, itemFolder, false);
            AddModelWrapper(parentContentLink, modelWrapper, ref folderDictionary);

            // Remove CMS cache for parent content link if time created or time last modified greater than time loaded
            var timeLoadedCacheKey = string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, parentContentLink);
            
            var lastTimeLoaded = GetLastTimeLoaded(timeLoadedCacheKey);
            if (lastTimeLoaded > DateTime.MinValue)
            {
                lock (_lockObject)
                {
                    // NOTE: the premise SharePoint (addon Team virtual machine) returns no data for SPFolder TimeCreated, TimeLastModified
                    // then using FilesFoldersService to load the Fodler properties 
                    if (itemFolder.Folder.TimeLastModified <= DateTime.MinValue)
                    {
                        SharePointHelper.SPServicesHub.SPFilesFoldersService.LoadFolderProperties(itemFolder.Folder);
                    }

                    // The Folder's name was changed or item was added to the folder
                    if (itemFolder.Folder.TimeLastModified > lastTimeLoaded)
                    {
                        // Remove cache of CMS
                        RemovePageCommonCache(modelWrapper.ContentLink);
                    }
                }
            }
        }

        private void AddFile(SPListItem itemFile, SPList selectedList, SPListView selectedView, string parentContentLink, ref Dictionary<string, List<IModelWrapper>> fileDictionary)
        {
            var file = itemFile.File;

            // Create model wrapper and get/set content link for it
            var modelWrapper = new ModelWrapper(selectedList, selectedView, itemFile, true);
            AddModelWrapper(parentContentLink, modelWrapper, ref fileDictionary);

            // Remove CMS cache for parent content link if time created or time last modified greater than time loaded
            var timeLoadedCacheKey = string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, parentContentLink);
            var previousLoaded = GetCachedItem(timeLoadedCacheKey);

            if (previousLoaded != null)
            {
                var timeLoaded = (DateTime)previousLoaded;
                // the file is new created or modified after prior loaded time
                if (file.TimeCreated > timeLoaded || file.TimeLastModified > timeLoaded)
                {
                    // Remove cache of CMS
                    RemoveChildrenCache(new ContentReference(parentContentLink));

                    if (file.TimeCreated < timeLoaded)
                    {
                        // File has been modified, remove it from CMS cache to update metadata of the file or thumbnail if it is an image
                        RemovePageCommonCache(modelWrapper.ContentLink);
                    }
                }
            }
        }

        private void AddModelWrapper(string groupKey, IModelWrapper modelWrapper, ref Dictionary<string, List<IModelWrapper>> dictionary)
        {
            if (!dictionary.ContainsKey(groupKey))
            {
                dictionary.Add(groupKey, new List<IModelWrapper>());
            }
            dictionary[groupKey].Add(modelWrapper);
        }

        private IEnumerable<SPList> GetSupportedMediaLists(bool refresh)
        {
            _logger.Information("GetSupportedMediaLists, refresh: {0}.", refresh);

            IEnumerable<SPList> listCollection;
            if (refresh)
            {
                return GetAndCacheMediaLists();
            }

            listCollection = GetCachedItem(_supportMediaListCacheKey) as IEnumerable<SPList>;
            if (listCollection == null)
            {
                listCollection = GetAndCacheMediaLists();
            }

            return listCollection;
        }

        private IEnumerable<SPList> GetAndCacheMediaLists()
        {
            _logger.Information("Get supported media lists from SharePoint.");

            var listCollection = SharePointHelper.SPServicesHub.SPListsViewsService.GetLists((int)ListBaseType.DocumentLibrary);
            if (listCollection != null && listCollection.Count() > 0)
            {
                _logger.Information("Add media lists into cache, cache key: {0} - items count: {1}.", _supportMediaListCacheKey, listCollection.Count());
                AddToCache(_supportMediaListCacheKey, listCollection);
            }
            return listCollection;
        }

        /// <summary>
        /// Add the given value to the cache with the registered cache key
        /// </summary>
        /// <param name="cacheKey">The given cache key</param>
        /// <param name="cacheItem">The given cache value</param>
        private void AddToCache(string cacheKey, object cacheItem)
        {
            AddToCache(cacheKey, cacheItem, GetAbsExpireUpdateCallbackPolicy());
        }

        /// <summary>
        /// Add an object to cache with the given cache key and policy
        /// </summary>
        private void AddToCache(string cacheKey, object cacheItem, CacheItemPolicy cacheItemPolicy)
        {
            _logger.Debug("Now = '{0}' adding to cache key = '{1}', minutes = '{2}'.", DateTime.Now, cacheKey, cacheItemPolicy.AbsoluteExpiration.Minute);
            _cache.Set(cacheKey, cacheItem, cacheItemPolicy);
        }

        /// <summary>
        /// Gets the cached value by the given cache key
        /// </summary>
        /// <param name="cacheKey">The given cache key</param>
        private object GetCachedItem(string cacheKey)
        {
            _logger.Debug("Get item from cache with key = '{0}'.", cacheKey);
            return _cache[cacheKey];
        }

        /// <summary>
        /// Removes the cached item from the cache by the given cache key
        /// </summary>
        /// <param name="cacheKey">The given cache key</param>
        private void RemoveCachedItem(string cacheKey)
        {
            _logger.Debug("Remove cache for key = '{0}'.", cacheKey);
            if (_cache.Contains(cacheKey))
            {
                _cache.Remove(cacheKey);
            }
        }

        /// <summary>
        /// Callback method which is invoked before a cached item is removed from the cache.
        /// </summary>
        /// <param name="arguments">The arguments.</param>
        /// <remarks>
        /// Everytime the DataMapping cache item is about expired, build the new value and assign it to the UpdatedCacheItem
        /// variable which is then used to update the cache object.
        /// </remarks>
        private void CachedItemUpdateCallback(CacheEntryUpdateArguments arguments)
        {
            _logger.Debug("CacheEntryUpdate for key = '{0}' at '{1}'.", arguments.Key, DateTime.Now);
            if (arguments.RemovedReason == CacheEntryRemovedReason.Expired)
            {
                ContentReference contentLink;
                if (ContentReference.TryParse(arguments.Key, out contentLink))
                {
                    lock (_lockObject)
                    {
                        Dictionary<string, List<IModelWrapper>> files;
                        Dictionary<string, List<IModelWrapper>> folders;

                        GetFilesAndFoldersBelongToContent(contentLink, out files, out folders, true);

                        arguments.UpdatedCacheItem = new CacheItem(arguments.Key, new DataMapping(files, folders));
                        arguments.UpdatedCacheItemPolicy = GetAbsExpireUpdateCallbackPolicy();

                        AddToMapping(files, folders);
                    }
                }
            }
            _logger.Debug("CacheEntryUpdate done key = '{0}' at '{1}'.", arguments.Key, DateTime.Now);
        }

        /// <summary>
        /// Clear children cache from CMS of content link
        /// </summary>
        /// <param name="contentLink"></param>
        private void RemoveChildrenCache(ContentReference contentLink)
        {
            if (ContentReference.IsNullOrEmpty(contentLink))
            {
                return;
            }

            // Remove cache of CMS with key in form of "EPChildrenData:ContentReference.ID"
            string childrenCacheKey = DataFactoryCache.ChildrenCacheKey(contentLink);

            CacheManager.Remove(childrenCacheKey);
        }

        /// <summary>
        /// Remove page master language cache from CMS of content link
        /// </summary>
        /// <param name="contentLink"></param>
        private void RemovePageCommonCache(ContentReference contentLink)
        {
            if (ContentReference.IsNullOrEmpty(contentLink))
            {
                return;
            }

            // Remove cache of CMS with key in form of "EPLanguageData:ContentReference.ID"
            string pageCommonCacheKey = DataFactoryCache.PageCommonCacheKey(contentLink);

            CacheManager.Remove(pageCommonCacheKey);
        }

        private DataMapping GetDataMapping()
        {
            var mapping = CacheManager.Get(_dataMappingCacheKey) as DataMapping;
            if (mapping == null)
            {
                mapping = new DataMapping();
                CacheManager.Insert(_dataMappingCacheKey, mapping);
            }
            return mapping;
        }

        /// <summary>
        /// Return a absolute <see cref="CacheItemPolicy"/> of 30 minutes
        /// </summary>
        /// <param name="multiplyNumber">Number to multiply the Expiration Time.</param>
        private CacheItemPolicy GetAbsExpirePolicy()
        {
            return new CacheItemPolicy()
            {
                AbsoluteExpiration = DateTime.Now.AddMinutes(30)
            };
        }

        private CacheItemPolicy GetAbsExpireUpdateCallbackPolicy()
        {
            return new CacheItemPolicy
            {
                UpdateCallback = new CacheEntryUpdateCallback(CachedItemUpdateCallback),
                // AbsoluteExpiration = DateTime.Now.AddMinutes(2)   // use small time when debugging
                AbsoluteExpiration = DateTime.Now.AddMinutes(5)
            };
        }

        /// <summary>
        /// Compare the current view query has change since last cached time
        /// This is only used when loading ROOT foders
        /// </summary>
        private bool IsViewQueryChanged(SPListView currentView, string viewQueryCacheKey)
        {
            var cachedViewQuery = GetCachedItem(viewQueryCacheKey) as string;

            // There was NO cache of prior view query
            if (string.IsNullOrEmpty(cachedViewQuery))
            {
                return false;
            }

            return !cachedViewQuery.Equals(currentView.ListViewXml, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Update the current view's query from cached query.
        /// This is only used when loading SUB foders, after ROOT loading.
        /// </summary>
        /// <param name="currentView">The view need to update its query.</param>
        private void SyncWithCachedViewQuery(SPListView currentView)
        {
            lock (_lockObject)
            {
                var viewQueryCacheKey = GetViewQueryCacheKey(currentView);
                var cachedViewQuery = GetCachedItem(viewQueryCacheKey) as string;

                if (!string.IsNullOrEmpty(cachedViewQuery) &&
                    !cachedViewQuery.Equals(currentView.ListViewXml, StringComparison.OrdinalIgnoreCase))
                {
                    currentView.ListViewXml = cachedViewQuery;
                    _logger.Debug("View '{0}' query was changed and has been sync.", currentView.Id);
                }
            }
        }

        private DateTime GetLastTimeLoaded(string timeCacheKey)
        {
            var lastTime = GetCachedItem(timeCacheKey);

            return (lastTime == null) ? DateTime.MinValue : (DateTime)lastTime;
        }

        private void ResetCache()
        {
            try
            {
                if (_cache != null)
                {
                    _logger.Debug("Dispose '{0}' memory cache.", _cache.Name);
                    _cache.Dispose();
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Error while dispose cache: {0}", ex);
            }
            finally
            {
                _cache = new MemoryCache(Consts.SHAREPOINT_CONTENT_CACHE_NAME);
                _logger.Debug("Recreate '{0}' memory cache.", _cache.Name);
            }
        }

        private string GetViewQueryTimeChangedCacheKey(string viewCacheKey)
        {
            return string.Format("{0}_ViewQueryChangedTime", viewCacheKey);
        }

        private string GetViewQueryCacheKey(SPListView currentView)
        {
            string tempVar = null;
            return GetViewQueryCacheKey(currentView, out tempVar);
            //return string.Format("{0}_ViewQuery", GetViewCacheKey(currentView));
        }

        private string GetViewQueryCacheKey(SPListView currentView, out string viewCacheKey)
        {
            viewCacheKey = GetViewCacheKey(currentView);
            return string.Format("{0}_ViewQuery", viewCacheKey);
        }

        private static string GetViewCacheKey(SPListView currentView)
        {
            return string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, currentView.Id);
        }
    }
}