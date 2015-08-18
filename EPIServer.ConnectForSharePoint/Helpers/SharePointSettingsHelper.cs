using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.Core;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public class SharePointSettingsHelper : ISharePointSettingsHelper
    {
        public ContentReference GetEntryPoint()
        {
            return SharePointSettings.Current.EntryPoint;
        }
    }
}