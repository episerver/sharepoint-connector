using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Messaging.Messages
{
    public interface ISharePointMessage
    {
        SPListItem Item { get; set; }
    }
}
