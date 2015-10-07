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
            var cf = contentEventArgs.ContentLink;
            var repo = ServiceLocator.Current.GetInstance<IContentRepository>();
            var f = repo.Get<IContent>(cf);

            // We need to check that the folder is marked for delete, as this is Moved Event includes moving a folder,
            // as well as moving it to the trash
            if (f is ContentFolder && f.IsDeleted)
            {
                var fd = new FolderData();
                fd = fd.GetFolder(f.ContentLink.ID);

                // The folder passed in may not be one we created, it could be a root folder that a user made.  
                // If so, we still want to check to see if it has any children and delete them before deleting the current folder.
                var processor = new SharePointProcessor();
                processor.DeleteSubFolders(repo, f as ContentFolder);

                // Now delete this folder from the table.
                FolderData.Delete(fd);
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

            // list item add/update/delete messages
            registry.Register<SharePointListItemAddMessage>(handler);
            registry.Register<SharePointListItemUpdateMessage>(handler);
            registry.Register<SharePointListItemDeleteMessage>(handler);

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