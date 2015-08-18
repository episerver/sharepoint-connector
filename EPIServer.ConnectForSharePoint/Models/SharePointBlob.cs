using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using EPiServer.Framework.Blobs;
using EPiServer.Logging;
using System;
using System.Linq;

namespace EPiServer.ConnectForSharePoint
{
    /// <summary>
    /// A blob of SharePoint media.
    /// </summary>
    public class SharePointBlob : Blob
    {
        private static readonly ILogger _logger = LogManager.GetLogger();

        /// <summary>
        /// Keep the reference to SharePoint item, so we can lookup and read its file stream
        /// </summary>
        internal SPListItem FileItem { get; set; }

        /// <summary>
        /// Keep the content reference of share point file item
        /// </summary>
        internal ContentReference ContentLink { get; set; }

        public SharePointBlob(Uri id) : base(id) { }
    }
}