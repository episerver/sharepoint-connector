using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Messaging.Handlers
{
    public class SharePointMessageHandler :
        IMessageHandler<SharePointContentAddMessage>,
        IMessageHandler<SharePointContentUpdateMessage>,
        IMessageHandler<SharePointContentDeleteMessage>,
        IMessageHandler<SharePointFolderAddMessage>,
        IMessageHandler<SharePointFolderUpdateMessage>,
        IMessageHandler<SharePointFolderDeleteMessage>
    {
        private ILogger _logger;
        private ISharePointProcessor _documentProcessor;

        public SharePointMessageHandler()
        {
            var aFactory = new SharePointProcessorFactory();
            _documentProcessor = aFactory.GetSharePointProcessor();
            _logger = LogManager.GetLogger();
        }

        internal SharePointMessageHandler(ILogger theLogger, ISharePointProcessor theProcessor)
        {
            _logger = theLogger;
            _documentProcessor = theProcessor;
        }

        public void Handle(SharePointFolderAddMessage message)
        {
            _documentProcessor.AddFolder(message.Item);
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointFolderUpdateMessage message)
        {
            _documentProcessor.UpdateFolder(message.Item);
        }

        public void Handle(SharePointFolderDeleteMessage message)
        {
            _documentProcessor.DeleteFolder(message.Item);
        }

        /// <summary>
        /// Handles a SharePointContentAddMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentAddMessage message)
        {
            _documentProcessor.AddDocument(message.Item);
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentUpdateMessage message)
        {
            _documentProcessor.UpdateDocument(message.Item);
        }

        /// <summary>
        /// Handles a ContentDeletedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentDeleteMessage message)
        {
            _documentProcessor.DeleteDocument(message.Item);
        }
    }
}