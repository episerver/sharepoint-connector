using System;
using EPiServer.Core;
using EPiServer.DataAnnotations;
using System.ComponentModel.DataAnnotations;
using EPiServer.Framework.DataAnnotations;
using EPiServer.Web;
using EPiServer.DataAbstraction;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    [ContentType(DisplayName = "SharePointBlockData", Description = "")]
    public class SharePointBlockData : BlockData, ExtraMediaData
    {
        public virtual DateTime Created { get; set; }

        public virtual XhtmlString Body { get; set; }

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