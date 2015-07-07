using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Rest;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class SPFolder
    {
        public string Name { get; set; }
        public string ServerRelativeUrl { get; set; }
        public int ItemCount { get; set; }
        public DateTime TimeCreated { get; set; }
        public DateTime TimeLastModified { get; set; }

        /// <summary>
        /// This is empty until call LoadParentFolder
        /// </summary>
        /// <remarks> This couldn't be lazy load because the serializer has known about SPFolder type 
        /// and trying to access this at deserializing SPFolder time causes exception
        /// </remarks>
        public SPFolder ParentFolder { get; set; }

        /// <summary>
        /// Is the top most folder which its parent points to itself.
        /// The top most folder has empty Name.
        /// </summary>
        public bool IsRootFolder
        {
            get {
                return ServerRelativeUrl.Equals("/");
            }
        }

        /// <summary>
        /// Reference back to its ListItem
        /// </summary>
        public SPListItem ListItem { get; set; }
    }
}