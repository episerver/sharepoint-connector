using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using System;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    /// <summary>
    /// The SharePointMessageEmitterFactory ecapuusates the "sharepointmod" queue, it also 
    /// ecapsulates the implementation object type for the queue
    /// message queue.
    /// </summary>
    public class SharePointMessageEmitterFactory : IMessageEmitterFactory
    {
        private readonly InMemoryQueueStore queueStore;
        internal const string QueueName = "sharepointmod";

        public SharePointMessageEmitterFactory()
        {
            this.queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
        }

        public IMessageEmitter Create()
        {
            return new InMemoryMessageEmitter(this.queueStore.Get(QueueName));
        }

        public int QueCount
        {
            get
            {
                 return queueStore.Get(SharePointMessageEmitterFactory.QueueName).Count;
            }
        }

        public void Initialize(MessageHandlerRegistry registry)
        {
            var queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
            var queue = queueStore.Get(SharePointMessageEmitterFactory.QueueName);
            var messageDispatcher = new FanOutMessageDispatcher(registry);
            var messageReceiver = new InMemoryMessageReceiver(messageDispatcher, queue);

            var appBuilder = new MessagingApplicationBuilder();
            appBuilder.Add(messageReceiver);
            appBuilder.App.Start();
        }
    }
}