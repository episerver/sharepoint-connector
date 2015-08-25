using EPiServer.Core;
using EPiServer.DataAnnotations;
using EPiServer.Framework.DataAnnotations;
using System;
using System.ComponentModel.DataAnnotations;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    [ContentType(GUID = "92700CD2-BCAB-4C3B-BD5E-14BBB9DFFD4F")]
    [MediaDescriptor(ExtensionString = "flv,mp4,webm")]
    public class SharepointVideoFile : VideoData, ExtraMediaData
    {
        [Editable(false, AllowInitialValue = true)]
        public virtual string SharePointGUID { get; set; }
        [Editable(false, AllowInitialValue = true)]
        [System.ComponentModel.DataAnnotations.ScaffoldColumn(false)]
        public virtual string SharePointParentListID { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public virtual DateTime SharePointTimeCreated { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public virtual DateTime SharePointTimeLastModified { get; set; }
        [Editable(false, AllowInitialValue = true)]
        [System.ComponentModel.DataAnnotations.ScaffoldColumn(false)]
        public virtual int SharePointMajorVersion { get; set; }
        [Editable(false, AllowInitialValue = true)]
        [System.ComponentModel.DataAnnotations.ScaffoldColumn(false)]
        public virtual int SharePointMinorVersion { get; set; }
        [Editable(false, AllowInitialValue = true)]
        [System.ComponentModel.DataAnnotations.ScaffoldColumn(false)]
        public virtual string SharePointRelativeUrl { get; set; }
    }
}