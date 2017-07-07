using EPiServer.Data.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.Data;
using System.Runtime.Caching;

namespace EPiServer.ConnectForSharePoint.Models
{
    [EPiServerDataStore(AutomaticallyRemapStore = true,
        AutomaticallyCreateStore = true, StoreName = "SPFolderMappings")]
    ///
    /// Dynamic data store class that represents the mapping between folder Guids and cms entry points 
    ///
    public class FolderData : IDynamicData
    {
        public Identity Id { get; set; }
        public int RefId { get; set; }
        public Guid FolderGuid { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        static private DynamicDataStoreFactory ddsf;

        protected void Initialize()
        {
            this.Id = Identity.NewIdentity(Guid.NewGuid());
            this.RefId = -1;
            this.FolderGuid = Guid.Empty;
            this.Created = DateTime.Now;
            this.Modified = DateTime.Now;
            if (ddsf == null)
            {
                ddsf = DynamicDataStoreFactory.Instance;
            }
        }

        internal FolderData(DynamicDataStoreFactory factory)
        {
            ddsf = factory;
        }

        public FolderData()
        {
            Initialize();
        }

        public FolderData(int refId, Guid folderGuid, DateTime created, DateTime modified)
        {
            Initialize();
            this.RefId = refId;
            this.FolderGuid = folderGuid;
            this.Created = created;
            this.Modified = modified;
        }

        public void Save()
        {
            // Create a data store (but only if one doesn't exist, we won't overwrite an existing one)
            var store = ddsf.CreateStore(typeof(FolderData));

            // Save the current FolderData
            store.Save(this);

            var data = (List<FolderData>)_fdCache.Get(CACHE_KEY);
            if (data != null)
            {
                data.Add(this);
            }
        }

        public static void Delete(FolderData folderData)
        {
            // Create a data store (but only if one doesn't exist, we won't overwrite an existing one)
            var store = ddsf.CreateStore(typeof(FolderData));

            // Delete the specified FolderData
            store.Delete(folderData.Id);

            var data = (List<FolderData>)_fdCache.Get(CACHE_KEY);
            if (data != null)
            {
                FolderData f = data.FirstOrDefault(x => x.RefId == folderData.RefId);
                if (f != null)
                {
                    data.Remove(f);
                }
            }
        }

        public virtual FolderData GetFolder(int refId)
        {
            var folders = GetFolders().Where(x => x.RefId == refId);

            if (folders == null || folders.Count() == 0) return new FolderData();

            //return first item as FolderData;
            return folders.ToList().First<FolderData>();
        }

        public FolderData GetFolder(Guid folderGuid)
        {
            // Retrieve comments
            var folders = GetFolders().Where(x => x.FolderGuid == folderGuid);

            if (folders == null || folders.Count<FolderData>() == 0) return new FolderData();

            //return first item as FolderData;
            return folders.ToList().First<FolderData>();
        }

        private static ObjectCache _fdCache = MemoryCache.Default;
        private static string CACHE_KEY = "EPi_Sharepint_FDList";

        public IEnumerable<FolderData> GetFolders()
        {
            IEnumerable<FolderData> items;

            if (_fdCache.Contains(CACHE_KEY))
            {
                items = (IEnumerable<FolderData>)_fdCache.Get(CACHE_KEY);
            }
            else
            {
                // Get data store for FolderData
                var store = ddsf.CreateStore(typeof(FolderData));
                List<FolderData> data = new List<FolderData>();
                data.AddRange(store.Items<FolderData>());

                CacheItemPolicy policy = new CacheItemPolicy();
                policy.SlidingExpiration = TimeSpan.FromSeconds(5);
                _fdCache.Add(CACHE_KEY, data, policy);
                items = data;
            }

            return items;
        }

        internal static void clearcache()
        {
            if (_fdCache != null && _fdCache.Contains(CACHE_KEY))
            {
                _fdCache.Remove(CACHE_KEY);
            }
        }
    }
}