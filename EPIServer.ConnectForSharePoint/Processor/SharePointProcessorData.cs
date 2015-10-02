using EPiServer.ConnectForSharePoint.Models;
using System;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessorData
    {
        public SPListItem ItemData { get; set; }
        public int CmsParentId { get; set; }
        public Guid SharePointParentId { get; set; }
    }
}