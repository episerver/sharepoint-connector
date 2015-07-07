using System;
using EPiServer.Core;

namespace EPiServer.ConnectForSharePoint.Models
{
    public interface IModelWrapper
    {
        /// <summary>
        /// The blob uri corresponding to current item.
        /// </summary>
        Uri BlobUri { get; set; }

        /// <summary>
        /// Contentlink corresponding to current item.
        /// </summary>
        ContentReference ContentLink { get; set; }

        bool IsFile { get; }

        /// <summary>
        /// The SharePoint list the current item belongs to.
        /// </summary>
        SPList List { get; set; }

        /// <summary>
        /// The SharePoint view the current item belongs to.
        /// </summary>
        SPListView View { get; set; }

        /// <summary>
        /// SharePoint ListItem object (file/folder)
        /// </summary>
        SPListItem ListItem { get; set; }

        /// <summary>
        /// Mapping on fly for list item
        /// </summary>
        string ToMappingUrl();

        /// <summary>
        /// Return a string in the concatenation of the list, view identification and the listItem identification format
        /// </summary>
        string UniqueKey();
    }
}
