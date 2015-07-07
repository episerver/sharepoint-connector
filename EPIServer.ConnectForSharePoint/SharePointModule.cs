using EPiServer.Framework.Web.Resources;
using EPiServer.Shell.Modules;

namespace EPiServer.ConnectForSharePoint
{
    public class SharePointModule : ShellModule
    {
        public SharePointModule(string name, string routeBasePath, string resourceBasePath)
            : base(name, routeBasePath, resourceBasePath) { }

        public override ModuleViewModel CreateViewModel(
            ModuleTable moduleTable, IClientResourceService clientResourceService)
        {
            return new SharePointModuleViewModel(this, clientResourceService);
        }
    }
}