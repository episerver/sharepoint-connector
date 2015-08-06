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
            throw new NotImplementedException();
        }

        public void Handle(SharePointFolderDeleteMessage message)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Handles a SharePointContentAddMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentAddMessage message)
        {
            _documentProcessor.AddDocument(message.JSON);
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentUpdateMessage message)
        {
            _logger.Warning("SharePointContentUpdateMessage notimplemented");
        }

        /// <summary>
        /// Handles a ContentDeletedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentDeleteMessage message)
        {
            _logger.Warning("SharePointContentDeleteMessage notimplemented");
        }
    }
}