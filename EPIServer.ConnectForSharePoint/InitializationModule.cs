using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using EPiServer.Framework;
using EPiServer.Logging;
using EPiServer.Packaging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Concurrent;

namespace EPiServer.ConnectForSharePoint
{
    /// <summary>
    /// Hook into context init event, to mount the SharePoint root node under Global Asset root
    /// </summary>
    [InitializableModule]
    [ModuleDependency(typeof(EPiServer.Web.InitializationModule))]
    public class InitializationModule : PackageInitializer, IConfigurableModule
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        MessageHandlerRegistry registry;
        MessagingApplicationBuilder appBuilder;

        #region IInitializableModule implementation

        /// <summary>
        /// Event listener so that we can detect if the user deletes one of our sharepoint folders, 
        /// this allows us to clean up the folder id to sharepoint quid table.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="contentEventArgs"></param>
        public void MovedContent(object sender, ContentEventArgs contentEventArgs)
        {
            ContentReference cf = contentEventArgs.ContentLink;
            var repo = ServiceLocator.Current.GetInstance<IContentRepository>();
            IContent f = repo.Get<IContent>(cf);
            if (f != null && f is ContentFolder )
            {
                FolderData fd = new FolderData();
                fd = fd.GetFolder(f.ContentLink.ID);
                if (fd.RefId != -1) // folder ref not found in table, do nothing...
                {
                    // cleanup all the sub folders properly, ie user is deleting a top level
                    // folder that we created * bad idea...
                    SharePointProcessor processor = new SharePointProcessor();
                    processor.DeleteSubFolders(f as ContentFolder);

                    // now delete this folder from the table.
                    FolderData.Delete(fd);
                }
            }
        }

        /// <inheritdoc /> 
        public override void Initialize(Framework.Initialization.InitializationEngine context)
        {
            _logger.Debug("ConnectForSharePoint AddOn Initialize");

            context.InitComplete += InitializationEngine_InitComplete;

            // add an event listener so that we get a message if the user moves something to the trash
            // it could be one of our folders so in that case we want to remove the entry from the 
            // sharepoint quid to cms content folder mapping table
            DataFactory.Instance.MovedContent += MovedContent;
    }

        #endregion

        #region Implementation of IPackageNotification

        /// <inheritdoc /> 
        public override void AfterInstall()
        {
        }

        /// <inheritdoc /> 
        public override void AfterUpdate()
        {
        }

        /// <inheritdoc /> 
        public override void BeforeUninstall()
        {
            ContentProviderHelper.DeleteSharePointMediaProvider();
        }

        #endregion

        private void InitializationEngine_InitComplete(object sender, EventArgs e)
        {
            InitializeMessagingSystem();
        }

        private void InitializeMessagingSystem()
        {
            _logger.Information("InitializeMessagingSystem");
            var handler = new SharePointMessageHandler();

            // RegisterMessageHandlers
            registry = new MessageHandlerRegistry();

            // document add/update/delete messages
            registry.Register<SharePointContentAddMessage>(handler);
            registry.Register<SharePointContentUpdateMessage>(handler);
            registry.Register<SharePointContentDeleteMessage>(handler);
            
            // folder add/update/delete messages
            registry.Register<SharePointFolderAddMessage>(handler);
            registry.Register<SharePointFolderUpdateMessage>(handler);
            registry.Register<SharePointFolderDeleteMessage>(handler);

            // Initialize message queue and the message dispatcher
            var messageDispatcher = new FanOutMessageDispatcher(registry);
            var queue = GetQueue("sharepointmod"); // Queue name has to match what is in SharePointMessageEmitterFactory
            var messageReceiver = new InMemoryMessageReceiver(messageDispatcher, queue);

            // Initialize the message application builder
            appBuilder = new MessagingApplicationBuilder();
            appBuilder.Add(messageReceiver);
            appBuilder.App.Start();
        }

        private BlockingCollection<object> GetQueue(string queueName)
        {
            _logger.Information("GetQueue [" + queueName + "]");

            var queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
            var queue = queueStore.Get(queueName);

            return queue;
        }

        /// <summary>
        /// Override the ContentMediaResolver and use our own SPContentMediaResolver so that 
        /// we can use our media types
        /// </summary>
        /// <param name="context"></param>
        public void ConfigureContainer(ServiceConfigurationContext context)
        {
        }
    }
}