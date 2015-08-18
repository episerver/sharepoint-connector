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

        #region public methods

        #endregion

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
    }
}