using System;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    public enum SharePointItemType
    {
        LibraryItem,
        ListItem
    }

    public interface ISharePointItem
    {
        /// <summary>
        /// Sharepoint item Guid for this object.
        /// </summary>
        string SharePointGUID { get; set; }
        /// <summary>
        /// When the item was added to cms, we set the last time modified so we can detect changes
        /// </summary>
        DateTime SharePointTimeLastModified { get; set; }
        /// <summary>
        /// One of the supported types that this connector supports
        /// </summary>
        SharePointItemType Type { get; }
        /// <summary>
        /// Sharepoint list Guid
        /// </summary>
        string ListID { get; set; } 
    }
}
