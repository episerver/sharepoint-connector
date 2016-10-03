using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface ISharePointMessageHelper
    {
        void QueueRootFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null);
        void BeginProcessing(string listids, IEnumerable<SPList> splists);
        List<SPListItem> ProcessSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, SPListItem ParentFolder = null);
        void RemoveSharePointAddedFolders(ContentReference folderRef, IEnumerable<SPListItem> sharepointListItems);
    }
}
