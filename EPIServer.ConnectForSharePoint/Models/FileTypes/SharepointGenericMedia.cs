using EPiServer.Core;
using EPiServer.DataAnnotations;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
    [ContentType(GUID = "BF0967F4-27BE-4EE4-BA40-EBAE307264B2")]
    public class SharepointGenericMedia : MediaData, ExtraMediaData
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