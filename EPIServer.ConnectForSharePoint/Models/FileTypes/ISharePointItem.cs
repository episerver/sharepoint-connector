using System;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    public enum SharePointItemType
    {
        LibraryItem,
        ListItem
    }

    interface ISharePointItem
    {
        string SharePointGUID { get; set; }
        DateTime SharePointTimeLastModified { get; set; }
        SharePointItemType Type { get; }
    }
}
