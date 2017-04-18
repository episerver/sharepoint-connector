using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.PlugIn;
using System.Linq;

namespace EPiServer.ConnectForSharePoint
{
    /// 
    /// CMS Update Job thread - retrieves data from sharepoint when task is triggered from the scheduler.
    ///
    [ScheduledPlugIn(DisplayName = "SharePoint Update", Description = "", LanguagePath = "/sharepoint/scheduledjob")]
    public class UpdateJob
    {
        private static IServicesHub _hub;
        private static ISharePointMessageHelper _helper;
        protected static string _ListIds;

        /// <summary>
        /// Helper method for unit testsing, allows us to pass in mocked objects so we can verify the sequence of events.
        /// </summary>
        /// <param name="hub"></param>
        /// <param name="helper"></param>
        /// <param name="ListIds"></param>
        internal UpdateJob(IServicesHub hub, ISharePointMessageHelper helper, string ListIds)
        {
            _hub = hub;
            _helper = helper;
            _ListIds = ListIds;
        }

        public static string Execute()
        {
            if (_hub == null)
                _hub = SharePointHelper.SPServicesHub;

            if (_helper == null)
                _helper = new SharePointMessageHelper();

            // Concatinate the document and custom lists.
            var documentLists = _hub.SPListsViewsService.GetLists((int)ListBaseType.DocumentLibrary);
            var customLists = _hub.SPListsViewsService.GetLists((int)ListBaseType.GenericList);
            var supportedLists = documentLists.Concat(customLists);

            // process the lists, we want to make sure we are always getting the current lists
            // when not doing unit test.
            if( _ListIds == null )
                _helper.BeginProcessing(SharePointSettings.Current.ListIds, supportedLists);
            else
                _helper.BeginProcessing(_ListIds, supportedLists);

            return "Started";
        }
    }
}