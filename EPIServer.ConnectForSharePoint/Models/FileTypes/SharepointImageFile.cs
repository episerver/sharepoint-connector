using EPiServer.Core;
using EPiServer.DataAnnotations;
using EPiServer.Framework.DataAnnotations;
using System;
using System.ComponentModel.DataAnnotations;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    [ContentType(GUID = "52C6DAC6-A321-4F21-A535-12DB6CD4062A")]
    [MediaDescriptor(ExtensionString = "jpg,jpeg,jpe,ico,gif,bmp,png")]
    public class SharepointImageFile : ImageData, ExtraMediaData
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