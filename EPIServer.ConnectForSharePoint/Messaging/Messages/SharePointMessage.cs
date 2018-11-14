using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Messaging.Messages
{
    public class SharePointMessage : ISharePointMessage
    {
        public SPListItem Item { get; set; }
    }
}