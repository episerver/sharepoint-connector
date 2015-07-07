using System;
using EPiServer.ConnectForSharePoint.Rest;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class SPListItem
    {
        /// <summary>
        /// The GUID identifier of the item.
        /// </summary>
        public Guid GUID { get; set; }

        /// <summary>
        /// The Id identifier of the item in its parent list.
        /// </summary>
        public int Id { get; set; }

        public int FileSystemObjectType { get; set; }

        // NOTE: the DocumentLibrary's item has Title property always null string
        public string Title { get; set; }

        public string EncodedAbsUrl { get; set; }

        public DateTime Created { get; set; }

        public DateTime Modified { get; set; }

        public SPList ParentList { get; set; }

        public SPFile File { get; set; }

        public SPFolder Folder { get; set; }

        public string FileUrl { get; set; }

        public string FolderUrl { get; set; }

        // NOTE: the DocumentLibrary's item has no DisplayName property
        public string DisplayName
        {
            get
            {
                return Title ?? FileLeafRef;
            }
        }

        /// <summary>
        /// Return the associate File/Folder name
        /// </summary>
        protected virtual string FileLeafRef
        {
            get
            {
                return Folder != null ? Folder.Name :
                    File != null ? File.Name : string.Empty;
            }
        }

        public bool? VideoSetShowDownloadLink { get; set; }

        public bool? VideoSetShowEmbedLink { get; set; }
    }
}