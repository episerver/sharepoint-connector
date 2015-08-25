using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Core;
using EPiServer.Framework;
using EPiServer.Framework.Blobs;
using EPiServer.Logging;
using EPiServer.Packaging;
using EPiServer.ServiceLocation;
using EPiServer.Web.Hosting;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Reflection;
using System.Threading;
using System.Web.Configuration;
using System.Web.Hosting;

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

        /// <inheritdoc /> 
        public override void Initialize(Framework.Initialization.InitializationEngine context)
        {
            _logger.Debug("ConnectForSharePoint AddOn Initialize");

            context.InitComplete += InitializationEngine_InitComplete;
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
            Thread.Sleep(5000);
            // Replace the MediaResolver with our own
            context.Container.Configure(ce =>
            {
                ce.For<ContentMediaResolver>().Use<SPContentMediaResolver>();
            });
        }
    }
}