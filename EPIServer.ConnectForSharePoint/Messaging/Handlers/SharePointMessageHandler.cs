using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.Logging;
using EPiServer.Security;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Principal;
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
        private string error1 = "* Verify that CmsAdmins virtual role exists in the web.config file and that it maps " +
            "to the \"Administrators\" user group.";
        private string error2 = "* Verify that CmsAdmins role or the role that the CmsAdmins role maps to has full access to " +
            "the folder specified in the \"Connect for Sharepoint\" configuration.";


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

        private void showErrorHelperMessages()
        {
            _logger.Error(error1);
            _logger.Error(error2);
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
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.AddFolder(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error AddFolder", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointFolderUpdateMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.UpdateFolder(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error UpdateFolder", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }

        public void Handle(SharePointFolderDeleteMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.DeleteFolder(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error DeleteFolder", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }

        /// <summary>
        /// Handles a SharePointContentAddMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentAddMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.AddDocument(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error AddDocument", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }

        /// <summary>
        /// Handles a ContentUpdatedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentUpdateMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.UpdateDocument(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error UpdateDocument", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }

        /// <summary>
        /// Handles a ContentDeletedMessage
        /// </summary>
        /// <param name="message">Message describing the operation</param>
        public void Handle(SharePointContentDeleteMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.DeleteDocument(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error DeleteDocument", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }

        public void Handle(SharePointListItemAddMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.AddListItem(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error AddListItem", e);
                    if( e is EPiServer.Core.AccessDeniedException )
                        showErrorHelperMessages();
                }
            }
        }

        public void Handle(SharePointListItemUpdateMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.UpdateListItem(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error UpdateListItem", e);
                    if( e is EPiServer.Core.AccessDeniedException )
                        showErrorHelperMessages();
                }
            }
        }

        public void Handle(SharePointListItemDeleteMessage message)
        {
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    processor.DeleteListItem(message.Item);
                }
                catch (Exception e)
                {
                    _logger.Error("Error DeleteListItem", e);
                    if (e is EPiServer.Core.AccessDeniedException)
                        showErrorHelperMessages();
                }
            }
        }
    }

    /// <summary>
    /// Class that sets the current principal to be a generic identity that is part of the CmsAdmins group,
    /// this allows us to perform updates in a folder without having to toggle permissions etc, the only requirement
    /// is that the folder has the CmsAdmins group assigned to it with full administrative access.
    /// </summary>
    public class PrincipalPreserver : IDisposable
    {
        IPrincipal p;

        public PrincipalPreserver()
        {
            p = PrincipalInfo.CurrentPrincipal;
            PrincipalInfo.CurrentPrincipal = new GenericPrincipal(
                            new GenericIdentity("Sharepoint Processor"),
                            new[] { "CmsAdmins" });
        }

        public void Dispose()
        {
            PrincipalInfo.CurrentPrincipal = p;
        }
    }
}