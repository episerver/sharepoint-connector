using EPiServer.ConnectForSharePoint.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    interface ISharePointMessageHelper
    {
        void QueueFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null);
        void QueueItemsForAdd(string listids, IEnumerable<SPList> splists);
        void AddSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, string ParentFolderName = null);
    }
}
