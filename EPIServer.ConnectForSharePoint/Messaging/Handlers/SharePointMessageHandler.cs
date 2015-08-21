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

        /// <summary>
        /// Default Constructor
        /// </summary>
        public SharePointMessageHandler()
        {
            _logger = LogManager.GetLogger();
        }

        internal SharePointMessageHandler(ILogger theLogger, ISharePointProcessor theProcessor)
        {
            _logger = theLogger;
            _documentProcessor = theProcessor;
        }

        /// <summary>
        /// Used internally to instantiate a share point processor only when needed. 
        /// </summary>
        private ISharePointProcessor processor
        {
            get
            {
                if (_documentProcessor != null)
                    return _documentProcessor;
                else
                {
                    var aFactory = new SharePointProcessorFactory();
                    _documentProcessor = aFactory.GetSharePointProcessor();
                    return _documentProcessor;
                }
            }
        }

        public void Handle(SharePointFolderAddMessage message)
        {
            processor.AddFolder(message.Item);
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointFolderUpdateMessage message)
        {
            processor.UpdateFolder(message.Item);
        }

        public void Handle(SharePointFolderDeleteMessage message)
        {
            processor.DeleteFolder(message.Item);
        }

        /// <summary>
        /// Handles a SharePointContentAddMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentAddMessage message)
        {
            processor.AddDocument(message.Item);
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentUpdateMessage message)
        {
            processor.UpdateDocument(message.Item);
        }

        /// <summary>
        /// Handles a ContentDeletedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentDeleteMessage message)
        {
            processor.DeleteDocument(message.Item);
        }
    }
}