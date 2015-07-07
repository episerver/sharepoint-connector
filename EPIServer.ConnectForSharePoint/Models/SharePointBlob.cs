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

        internal bool IsThumbnail { get; set; }

        public SharePointBlob(Uri id) : base(id) { }

        /// <summary>
        /// Read the FileStream from the SharePoint item
        /// </summary>
        /// <returns></returns>
        public override System.IO.Stream OpenRead()
        {
            if (FileItem == null)
            {
                _logger.Debug("The sharepoint blob of content link {0} has file item null", ContentLink);
                var spContentCacher = ContentProviderHelper.SharePointMediaProvider.ContentCache;
                var item = spContentCacher.Files.Where(f => f.Value.Exists(v => v.BlobUri != null && v.BlobUri.Equals(this.ID))).FirstOrDefault();

                if (item.Value != null)
                {
                    var modelWrapper = item.Value.Where(v => v.BlobUri != null && v.BlobUri.Equals(ID)).FirstOrDefault();
                    if (modelWrapper != null)
                    {
                        FileItem = modelWrapper.ListItem;
                        ContentLink = modelWrapper.ContentLink;
                    }
                }
                else
                {
                    _logger.Information("Get file stream of Uri: {0}", ID.OriginalString);
                    return SharePointHelper.SPServicesHub.SPFileStreamService.GetFileStream(ID, IsThumbnail);
                }
            }

            return SharePointHelper.SPServicesHub.SPFileStreamService.GetFileStream(FileItem, ContentLink, IsThumbnail);
        }
    }
}