using EPiServer.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models.FileTypes
{
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