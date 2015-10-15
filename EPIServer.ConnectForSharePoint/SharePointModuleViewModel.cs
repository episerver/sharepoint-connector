using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.Framework.Web.Resources;
using EPiServer.Shell.Modules;

namespace EPiServer.ConnectForSharePoint
{
    public class SharePointModuleViewModel : ModuleViewModel
    {
        public SharePointModuleViewModel(ShellModule shellModule, IClientResourceService clientResourceService)
            : base(shellModule, clientResourceService) { }
    }
}