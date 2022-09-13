using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Processor;
using EPiServer.ConnectForSharePoint.Rest;
using EPiServer.ConnectForSharePoint.SPServices.REST;
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
        private static IMetaDataProcessor _metaDataProcessor;
        protected static string _listIds;
        protected static LocalizationService _localization;
        private IServiceLocator _serviceLocator;
        private bool _stopSignaled;

        public UpdateJob()
        {
            _serviceLocator = ServiceLocator.Current;
            IsStoppable = true;
        }

        /// <summary>
        /// Helper method for unit testsing, allows us to pass in mocked objects so we can verify the sequence of events.
        /// </summary>
        /// <param name="hub"></param>
        /// <param name="processor"></param>
        /// <param name="ListIds"></param>
        internal UpdateJob(IServiceLocator locator, IMetaDataProcessor processor, string ListIds )
        {
            _metaDataProcessor = processor;
            _listIds = ListIds;
            _serviceLocator = locator;
            IsStoppable = true;
        }

        /// <summary>
        /// Called when a user clicks on Stop for a manually started job, or when ASP.NET shuts down.
        /// </summary>
        public override void Stop()
        {
            _stopSignaled = true;
        }

        public override string Execute()
        {
            if (_metaDataProcessor == null)
            {
                _metaDataProcessor = _serviceLocator.GetInstance<IMetaDataProcessor>();
            }

            if (_metaDataProcessor.IsRunning())
            {
                return WaitForDone();
            }

            if (_localization == null)
                _localization = LocalizationService.Current;
            else
                _localization = ServiceLocator.Current.GetInstance<LocalizationService>();

            ContentReference destfolderRef;
            string foldername = "";

            destfolderRef = SharePointSettings.Current.EntryPoint;
            foldername = SharePointSettings.Current.EntryPointFolderName;

            // Concatenate the document and custom lists.
            var contextService = SharePointSettings.Current.IsModernAuth ?
            new EndpointContextService
            {
                EndpointContext = new SPRestClient(SharePointSettings.Current)
            } :
            new EndpointContextService
            {
                EndpointContext = new SPRestClient(SharePointSettings.Current, SharePointHelper.GetLatestCookies)
            };

            var listViewsService = new ListsViewsService { EndpointContextService = contextService };
            var documentLists = listViewsService.GetLists((int)ListBaseType.DocumentLibrary);
            var customLists = listViewsService.GetLists((int)ListBaseType.GenericList);
            var supportedLists = documentLists.Concat(customLists);

            // show error if destination folder has been deleted in cms
            ContentFolder f;
            var repo = ServiceLocator.Current.GetInstance<IContentRepository>();
            bool loaded = repo.TryGet<ContentFolder>(destfolderRef, out f);
            if ((ContentReference.EmptyReference != destfolderRef && loaded == true && !f.IsDeleted))
            {
                // process the lists, we want to make sure we are always getting the current lists
                // when not doing unit test.
                if (_listIds == null)
                {
                    _metaDataProcessor.BeginProcessing(SharePointSettings.Current.ListIds, supportedLists);
                }
                else
                {
                    _metaDataProcessor.BeginProcessing(_listIds, supportedLists);
                }

                return WaitForDone();
            }
            else
            {
                string msg = string.Format(_localization.GetString("/sharepoint/admin/message/folderdeleted"), foldername);
                OnStatusChanged(msg);
                throw new Exception(msg);
            }
        }

        public string WaitForDone()
        {
            var emitter = new SharePointMessageEmitterFactory();
            while (_metaDataProcessor.IsRunning())
            {
                OnStatusChanged( String.Format( "Processing {0} items.", emitter.QueueCount) );
                if (_stopSignaled)
                {
                    _metaDataProcessor.Stop();
                }
                Thread.Sleep(250);
            }

            // As long as items are in the queue make sure we show we are still running
            while (emitter.QueueCount > 0)
            {
                OnStatusChanged(String.Format("Processing {0} items.", emitter.QueueCount));
                if (_stopSignaled)
                {
                    emitter.Clear();
                }
                Thread.Sleep(250);
            }

            if (MetaDataProcessor._processingResult != true)
            {
                return ("Failed : " + MetaDataProcessor._processingException.Message);
            }
            else
            {
                return "Completed.";
            }
        }
    }
}