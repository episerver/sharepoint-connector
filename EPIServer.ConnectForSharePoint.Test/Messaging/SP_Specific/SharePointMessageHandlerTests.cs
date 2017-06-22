using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Test.Messaging.SP_Specific
{
    [TestClass]
    public class SharePointMessageHandlerTests
    {
        Mock<ILogger> logger;
        Mock<ISharePointProcessor> processor;

        private SharePointMessageHandler GetUnitUnderTest()
        {
            logger = new Mock<ILogger>();
            processor = new Mock<ISharePointProcessor>();

            return new SharePointMessageHandler(logger.Object, processor.Object);
        }
         

        [TestMethod]
        public void Handle_For_AddDocumentMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.AddDocument(It.IsAny<SPListItem>()));

            var message = new SharePointContentAddMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);
            handler._threadManager.WaitAll();

            processor.Verify(p => p.AddDocument(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_UpdateDocumentMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.UpdateDocument(It.IsAny<SPListItem>()));

            var message = new SharePointContentUpdateMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.UpdateDocument(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_DeleteDocumentMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.DeleteDocument(It.IsAny<SPListItem>()));

            var message = new SharePointContentDeleteMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.DeleteDocument(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_AddFolderMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.AddFolder(It.IsAny<SPListItem>()));

            var message = new SharePointFolderAddMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.AddFolder(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_UpdateFolderMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.AddFolder(It.IsAny<SPListItem>()));

            var message = new SharePointFolderUpdateMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.UpdateFolder(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_DeleteFolderMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.DeleteFolder(It.IsAny<SPListItem>()));

            var message = new SharePointFolderDeleteMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.DeleteFolder(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_AddListItem_Message_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.AddFolder(It.IsAny<SPListItem>()));

            var message = new SharePointListItemAddMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);
            handler._threadManager.WaitAll();

            processor.Verify(p => p.AddListItem(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_UpdateListItemMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.UpdateListItem(It.IsAny<SPListItem>()));

            var message = new SharePointListItemUpdateMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.UpdateListItem(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_DeleteListItemMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.DeleteListItem(It.IsAny<SPListItem>()));

            var message = new SharePointListItemDeleteMessage()
            {
                Item = new SPListItem()
                {
                    Title = "Some Title",
                }
            };

            handler.Handle(message);

            processor.Verify(p => p.DeleteListItem(It.Is<SPListItem>(item => item == message.Item)), Times.Once, "Did not call into the processor with the expected string");
        }
    }
}
