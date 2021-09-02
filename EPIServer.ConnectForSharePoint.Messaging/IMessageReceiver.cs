using System.Threading;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public interface IMessageReceiver
    {
        void Start(CancellationToken cancellationToken);
    }
}
