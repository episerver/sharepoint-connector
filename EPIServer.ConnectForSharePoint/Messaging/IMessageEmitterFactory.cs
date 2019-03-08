using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public interface IMessageEmitterFactory
    {
        /// <summary>
        /// Gets the current number of items in the queue
        /// </summary>
        int QueueCount { get; }

        /// <summary>
        /// Initializes the message message hander
        /// </summary>
        /// <param name="registry"></param>
        void Initialize(MessageHandlerRegistry registry);

        /// <summary>
        /// Clears all items in the queue.
        /// </summary>
        void Clear();

        /// <summary>
        /// Helper method that can be used to emit messages of type T
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="item"></param>
        void Emit<T>(SPListItem item) where T : ISharePointMessage, new();
    }
}
