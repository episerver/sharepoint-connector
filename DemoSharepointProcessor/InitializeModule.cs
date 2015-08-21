using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.Framework;
using EPiServer.Framework.Initialization;
using EPiServer.Packaging;
using EPiServer.ServiceLocation;
using System;
using System.Threading;

namespace DemoSharepointProcessor
{
    /// <summary>
    /// Episerver Module initializer for this plugin. 
    /// </summary>
    [InitializableModule]
    [ModuleDependency(typeof(EPiServer.Web.InitializationModule))]
    public class InitializeModule : PackageInitializer
    {
        IServiceLocator _serviceLocator;
        ISharePointProcessor _processor;

        public InitializeModule()
        {
            // Get the service locator from the environment
            _serviceLocator = ServiceLocator.Current;
            _processor = new DemoProcessor();
        }

        /// <summary>
        /// For unit testing
        /// </summary>
        /// <param name="serviceLocator"></param>
        internal InitializeModule(IServiceLocator serviceLocator, ISharePointProcessor processor)
        {
            // Use the Mocked service locator passed in
            _serviceLocator = serviceLocator;
            _processor = processor;
        }

        public override void Initialize(InitializationEngine context)
        {
            context.InitComplete += InitializationEngine_InitComplete;
        }

        private void InitializationEngine_InitComplete(object sender, EventArgs e)
        {
            _serviceLocator.Buildup(_processor); // Register our processor
        }

        public override void AfterInstall()
        {
        }

        public override void AfterUpdate()
        {
        }

        public override void BeforeUninstall()
        {
        }
    }
}
