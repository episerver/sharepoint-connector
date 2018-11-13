using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
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
        internal const string QueueName = "sharepointmod";

        private readonly InMemoryQueueStore queueStore;
        private InMemoryMessageEmitter messageEmmiter;

        public SharePointMessageEmitterFactory()
        {
            this.queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
        }

        public int QueueCount
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

        public void Clear()
        {
            var queue = queueStore.Get(SharePointMessageEmitterFactory.QueueName);
            object item;
            while (queue.TryTake(out item, 100)) ;
        }

        public void Emit<T>(SPListItem item) where T : ISharePointMessage, new()
        {
            var message = new T();
            message.Item = item;
            if( messageEmmiter == null )
            {
                messageEmmiter = new InMemoryMessageEmitter(this.queueStore.Get(QueueName));
            }
            messageEmmiter.Emit(message);
        }
    }
}