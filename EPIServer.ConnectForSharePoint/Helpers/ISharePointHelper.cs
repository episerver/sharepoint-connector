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
        IEnumerable<SPListItem> GetItems(SPList list, SPListView view, SPFolder folder);
    }
}
