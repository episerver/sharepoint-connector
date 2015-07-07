using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Framework;
using EPiServer.Framework.Blobs;
using EPiServer.Logging;
using EPiServer.Packaging;
using EPiServer.Web.Hosting;
using System;
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