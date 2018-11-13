using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Test.Messaging.SP_Specific
{
    [TestClass]
    public class SharePointMessageEmitterFactoryTests
    {
        static public object _MessagingLock = new Object();

        [TestMethod]
        public void Create_ReturnsAnInMemoryEmitterTargetingTheIndexingQueue()
        {
            // Get the que based on the name that should match in SharePointMessageEmitterFactory
            var queue = new InMemoryQueueStore(AppDomain.CurrentDomain).Get(SharePointMessageEmitterFactory.QueueName);

            // Get our SharePointMessageEmitterFactory factory and create an emmitter
            var factory = new SharePointMessageEmitterFactory();

            // emit a message.
            var item = new SPListItem() { Title = "SomeGreatItem!" };

            factory.Emit<SharePointFolderDeleteMessage>(item);

            // check the queue instance we have and verify that hte message we emmited is there
            Assert.AreEqual(item.Title, ((SharePointFolderDeleteMessage)queue.Take()).Item.Title, "Indexing queue contains the expected message");
        }

        [TestMethod]
        public void QueueCount_Returns_Expected_Count()
        {
            lock (_MessagingLock)
            {
                var factory = new SharePointMessageEmitterFactory();

                factory.Clear();

                // emit 3 messages
                factory.Emit<SharePointFolderDeleteMessage>(new SPListItem());
                factory.Emit<SharePointFolderDeleteMessage>(new SPListItem());
                factory.Emit<SharePointFolderDeleteMessage>(new SPListItem());

                Assert.AreEqual(factory.QueueCount, 3);
            }
        }

        [TestMethod]
        public void Clear_EmptiesQueue()
        {
            lock (_MessagingLock)
            {
                var factory = new SharePointMessageEmitterFactory();

                factory.Clear();

                // emit 3 messages
                factory.Emit<SharePointFolderDeleteMessage>(new SPListItem());
                factory.Emit<SharePointFolderDeleteMessage>(new SPListItem());
                factory.Emit<SharePointFolderDeleteMessage>(new SPListItem());

                factory.Clear();

                Assert.AreEqual(factory.QueueCount, 0);
            }
        }
    }
}
