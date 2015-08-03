using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

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

        public SharePointMessageEmitterFactory()
        {
            this.queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
        }

        public IMessageEmitter Create()
        {
            return new InMemoryMessageEmitter(this.queueStore.Get("sharepointmod"));
        }
    }
}