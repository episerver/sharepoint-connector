using EPiServer.ConnectForSharePoint.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface ISharePointHelper
    {
        string[] GetListAndViewIds(string mappingId);
        IEnumerable<KeyValuePair<string, string>> GetSelectedListViewIds(IEnumerable<SPList> lists);
        IEnumerable<SPListItem> GetPagedItems(SPList list, SPListView view, SPFolder folder, out string nextUri);
        IEnumerable<SPListItem> GetNextPagedItems(SPList list, string Uri, out string nextUri);
    }
}
