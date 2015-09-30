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
        IMessageHandler<SharePointFolderDeleteMessage>,
        IMessageHandler<SharePointListItemAddMessage>,
        IMessageHandler<SharePointListItemUpdateMessage>,
        IMessageHandler<SharePointListItemDeleteMessage>
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
            try
            {
                processor.AddFolder(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddFolder", e);
            }
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointFolderUpdateMessage message)
        {
            try
            { 
                processor.UpdateFolder(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error UpdateFolder", e);
            }
        }

        public void Handle(SharePointFolderDeleteMessage message)
        {
            try
            { 
                processor.DeleteFolder(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error DeleteFolder", e);
            }
        }

        /// <summary>
        /// Handles a SharePointContentAddMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentAddMessage message)
        {
            try
            {
                processor.AddDocument(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddDocument", e);
            }
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentUpdateMessage message)
        {
            try
            {
                processor.UpdateDocument(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddDocument", e);
            }
        }

        /// <summary>
        /// Handles a ContentDeletedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentDeleteMessage message)
        {
            try
            {
                processor.DeleteDocument(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddDocument", e);
            }
        }

        public void Handle(SharePointListItemAddMessage message)
        {
            try
            {
                processor.AddListItem(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddDocument", e);
            }
        }

        public void Handle(SharePointListItemUpdateMessage message)
        {
            try
            {
                processor.UpdateListItem(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddDocument", e);
            }
        }

        public void Handle(SharePointListItemDeleteMessage message)
        {
            try
            {
                processor.DeleteListItem(message.Item);
            }
            catch (Exception e)
            {
                _logger.Error("Error AddDocument", e);
            }
        }
    }
}