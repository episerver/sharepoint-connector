using EPiServer.ConnectForSharePoint.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface ISharePointMessageHelper
    {
        void queueRootFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null);
        void beginProcessing(string listids, IEnumerable<SPList> splists);
        void processSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, string ParentFolderName = null);
    }
}
