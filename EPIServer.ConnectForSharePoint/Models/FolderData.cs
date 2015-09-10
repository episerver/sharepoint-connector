using EPiServer.Data.Dynamic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.Data;

namespace EPiServer.ConnectForSharePoint.Models
{
    [EPiServerDataStore(AutomaticallyRemapStore = true, 
        AutomaticallyCreateStore = true, StoreName="SPFolderMappings")]
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
        }

        public static void Delete(FolderData folderData)
        {
            // Create a data store (but only if one doesn't exist, we won't overwrite an existing one)
            var store = ddsf.CreateStore(typeof(FolderData));

            // Delete the specified FolderData
            store.Delete(folderData.Id);
        }

        public virtual FolderData GetFolder(int refId)
        {
            // Get data store for FolderData
            var store = ddsf.CreateStore(typeof(FolderData));

            // Retrieve comments
            var items = store.Items<FolderData>();
            var folders = items.Where(x => x.RefId == refId);

            if (folders == null || folders.Count() == 0) return new FolderData();

            //return first item as FolderData;
            return folders.ToList().First<FolderData>();
        }

        public FolderData GetFolder(Guid folderGuid)
        {
            // Get data store for FolderData
            var store = ddsf.CreateStore(typeof(FolderData));

            // Retrieve comments
            var folders = store.Items<FolderData>().Where(x => x.FolderGuid == folderGuid);

            if (folders == null || folders.Count<FolderData>() == 0) return new FolderData();

            //return first item as FolderData;
            return folders.ToList().First<FolderData>();
        }
    }
}