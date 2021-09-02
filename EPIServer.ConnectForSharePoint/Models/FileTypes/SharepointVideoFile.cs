using EPiServer.Core;
using EPiServer.DataAnnotations;
using EPiServer.Framework.DataAnnotations;
using System;
using System.ComponentModel.DataAnnotations;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    [ContentType(GUID = "92700CD2-BCAB-4C3B-BD5E-14BBB9DFFD4F")]
    [MediaDescriptor(ExtensionString = "flv,mp4,webm,swf")]
    public class SharepointVideoFile : VideoData, ISharePointItem
    {
        [Editable(false, AllowInitialValue = true)]
        public virtual string SharePointGUID { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public virtual DateTime SharePointTimeLastModified { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public SharePointItemType Type
        {
            get { return SharePointItemType.LibraryItem; }
        }
        [Editable(false, AllowInitialValue = true)]
        public virtual string ListID { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public virtual string ViewID { get; set; }
    }
}