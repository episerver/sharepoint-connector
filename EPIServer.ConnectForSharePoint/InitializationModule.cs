using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Framework;
using EPiServer.Framework.Blobs;
using EPiServer.Logging;
using EPiServer.Packaging;
using EPiServer.Web.Hosting;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Reflection;
using System.Web.Configuration;
using System.Web.Hosting;

namespace EPiServer.ConnectForSharePoint
{
    /// <summary>
    /// Hook into context init event, to mount the SharePoint root node under Global Asset root
    /// </summary>
    [InitializableModule]
    [ModuleDependency(typeof(EPiServer.Web.InitializationModule))]
    public class InitializationModule : PackageInitializer
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
            AddSharePointBlobProvider();
            AddSharePointMediaProvider();
            SetAccessToClientResources();
            InitializeMessagingSystem();
        }

        private void InitializeMessagingSystem()
        {
            _logger.Information("InitializeMessagingSystem");
            var handler = new SharePointContentModificationHandler();

            // RegisterMessageHandlers
            registry = new MessageHandlerRegistry();
            registry.Register<SharePointContentAddMessage>(handler);
            registry.Register<SharePointContentUpdateMessage>(handler);
            registry.Register<SharePointContentDeleteMessage>(handler);

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

        private void AddSharePointBlobProvider()
        {
            var sharePointBlobProvider = (SharePointBlobProvider)ProvidersHelper
                .InstantiateProvider(new ProviderSettings("sharepoint", typeof(SharePointBlobProvider).AssemblyQualifiedName), typeof(SharePointBlobProvider));

            // Inject it into the blob factory providers
            var field = BlobFactory.Instance.GetType().GetField("_providers", BindingFlags.NonPublic | BindingFlags.Instance);
            var providers = field.GetValue(BlobFactory.Instance) as Dictionary<string, BlobProvider>;
            providers.Add("sharepoint", sharePointBlobProvider);
            field.SetValue(BlobFactory.Instance, providers);
        }

        private void AddSharePointMediaProvider()
        {
            var contentProviderType = typeof(SharePointMediaProvider);

            try
            {
                ContentProviderHelper.CreateSharePointMediaProvider();
            }
            catch (Exception ex)
            {
                _logger.Error(String.Format("Cannot initialize {0} content provider in initialization phase", contentProviderType), ex);
                return;
            }

            _logger.Debug("Initialize {0} content provider successfully", contentProviderType);
        }

        private void SetAccessToClientResources()
        {
            // TECHNOTE: allow access to protected module's client resources
            // because when viewing Video's elements in viewmode, we need to inject .js file in which anonymous user can access
            var publicVpp = new System.Collections.Specialized.NameValueCollection();
            publicVpp.Add("virtualPath", "~" + ModuleHelper.PublicVirtualPath);
            publicVpp.Add("physicalPath", ModuleHelper.ToPhysicalVPPClientResource("ClientResources"));
            publicVpp.Add("bypassAccessCheck", "true");

            HostingEnvironment.RegisterVirtualPathProvider(new VirtualPathNonUnifiedProvider("PublicSharePointClientResourceVpp", publicVpp));
        }
    }
}