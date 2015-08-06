using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;

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
        public void Handle_For_AddContentMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.AddDocument(It.IsAny<string>()));

            var message = new SharePointContentAddMessage()
            {
                JSON = "Test"
            };

            handler.Handle(message);

            processor.Verify(p => p.AddDocument(It.Is<string>(json => json == message.JSON)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_DeleteContentMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.DeleteDocument(It.IsAny<string>()));

            var message = new SharePointContentDeleteMessage()
            {
                JSON = "Test"
            };

            handler.Handle(message);

            processor.Verify(p => p.DeleteDocument(It.Is<string>(json => json == message.JSON)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_AddFolderMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.AddFolder(It.IsAny<string>()));

            var message = new SharePointFolderAddMessage()
            {
                JSON = "Test"
            };

            handler.Handle(message);

            processor.Verify(p => p.AddFolder(It.Is<string>(json => json == message.JSON)), Times.Once, "Did not call into the processor with the expected string");
        }

        [TestMethod]
        public void Handle_For_DeleteFolderMessage_Passes_Message_To_SharePointProcessor()
        {
            var handler = GetUnitUnderTest();
            processor.Setup(p => p.DeleteFolder(It.IsAny<string>()));

            var message = new SharePointFolderDeleteMessage()
            {
                JSON = "Test"
            };

            handler.Handle(message);

            processor.Verify(p => p.DeleteFolder(It.Is<string>(json => json == message.JSON)), Times.Once, "Did not call into the processor with the expected string");
        }
    }
}
