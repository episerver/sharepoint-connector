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
    public class SharePointBlockData : BlockData, ISharePointItem
    {
        public virtual DateTime Created { get; set; }

        public virtual string Json { get; set; }

        [Editable(false, AllowInitialValue = true)]
        public virtual string SharePointGUID { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public virtual DateTime SharePointTimeLastModified { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public SharePointItemType Type
        {
            get { return SharePointItemType.ListItem; }
        }
        [Editable(false, AllowInitialValue = true)]
        public virtual string ListID { get; set; }
        [Editable(false, AllowInitialValue = true)]
        public virtual string ViewID { get; set; }
    }
}