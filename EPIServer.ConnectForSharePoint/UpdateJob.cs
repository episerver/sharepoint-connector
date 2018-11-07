using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.Web;
using EPiServer.Core;
using EPiServer.Framework.Localization;
using EPiServer.PlugIn;
using EPiServer.Scheduler;
using EPiServer.ServiceLocation;
using System;
using System.Linq;
using System.Threading;

namespace EPiServer.ConnectForSharePoint
{
    /// 
    /// CMS Update Job thread - retrieves data from sharepoint when task is triggered from the scheduler.
    ///
    [ScheduledPlugIn(DisplayName = "SharePoint Update", Description = "", LanguagePath = "/sharepoint/scheduledjob")]
    public class UpdateJob : ScheduledJobBase
    {
        private static IServicesHub _hub;
        private static IMainProcessor _helper;
        protected static string _ListIds;
        protected static LocalizationService _localization;

        public UpdateJob()
        {
        }

        /// <summary>
        /// Helper method for unit testsing, allows us to pass in mocked objects so we can verify the sequence of events.
        /// </summary>
        /// <param name="hub"></param>
        /// <param name="helper"></param>
        /// <param name="ListIds"></param>
        internal UpdateJob(IServicesHub hub, IMainProcessor helper, string ListIds )
        {
            _hub = hub;
            _helper = helper;
            _ListIds = ListIds;
        }

        public override string Execute()
        {
            if (MainProcessor.IsRunning())
            {
                string msg = "Process is currently running.";
                OnStatusChanged(msg);
                throw new Exception(msg);
            }

            if (_hub == null)
                _hub = SharePointHelper.SPServicesHub;

            if (_helper == null)
                _helper = new MainProcessor();

            if (_localization == null)
                _localization = LocalizationService.Current;
            else
                _localization = ServiceLocator.Current.GetInstance<LocalizationService>();

            ContentReference destfolderRef;
            string foldername = "";

            destfolderRef = SharePointSettings.Current.EntryPoint;
            foldername = SharePointSettings.Current.EntryPointFolderName;

            // Concatinate the document and custom lists.
            var documentLists = _hub.SPListsViewsService.GetLists((int)ListBaseType.DocumentLibrary);
            var customLists = _hub.SPListsViewsService.GetLists((int)ListBaseType.GenericList);
            var supportedLists = documentLists.Concat(customLists);

            // show error if destination folder has been deleted in cms
            ContentFolder f;
            var repo = ServiceLocator.Current.GetInstance<IContentRepository>();
            bool loaded = repo.TryGet<ContentFolder>(destfolderRef, out f);
            if ( (ContentReference.EmptyReference != destfolderRef && loaded == true && !f.IsDeleted) )
            {
                // process the lists, we want to make sure we are always getting the current lists
                // when not doing unit test.
                if (_ListIds == null)
                    _helper.BeginProcessing(SharePointSettings.Current.ListIds, supportedLists);
                else
                    _helper.BeginProcessing(_ListIds, supportedLists);

                // Join the proccesing thread with the job thread so that the 
                // status of the job stays busy while its running.
                if (MainProcessor._processingThread != null) // fix unit test
                {
                    MainProcessor._processingThread.Join();
                }

                // As long as items are in the queue make sure we show we are still running
                while(  MainProcessor.GetQueCount() > 0 )
                {
                    Thread.Sleep(2000);
                }

                if (MainProcessor._processingResult != true)
                {
                    return ("Failed : " + MainProcessor._processingException.Message);
                }
                else
                {
                    return "Completed.";
                }
            }
            else
            {
                string msg = string.Format(_localization.GetString("/sharepoint/admin/message/folderdeleted"), foldername);
                OnStatusChanged(msg);
                throw new Exception(msg);
            }
        }
    }
}