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
    public class SharepointGenericMedia : MediaData, ISharePointItem
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
    }
}