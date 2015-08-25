using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    interface ExtraMediaData
    {
        string SharePointGUID { get; set; }
        string SharePointParentListID { get; set; }
        DateTime SharePointTimeCreated { get; set; }
        DateTime SharePointTimeLastModified { get; set; }
        int SharePointMajorVersion { get; set; }
        int SharePointMinorVersion { get; set; }
        string SharePointRelativeUrl { get; set; }
    }
}
