using EPiServer.BaseLibrary.Scheduling;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Framework.Localization;
using EPiServer.Logging;
using EPiServer.PlugIn;
using EPiServer.ServiceLocation;
using System;

namespace EPiServer.ConnectForSharePoint
{
    [ScheduledPlugIn(
        LanguagePath = SPContentIndexJob.LanguagePath,
        DefaultEnabled = true,
        InitialTime = "0",
        IntervalLength = 1,
        IntervalType = ScheduledIntervalType.Weeks)]
    public class SPContentIndexJob : JobBase
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        private readonly LocalizationService _localizationService = LocalizationService.Current;
        public const string LanguagePath = "/sharepoint/plugin/spcontentindexjob";

        /// <summary>
        /// Index the content from SharePoint
        /// </summary>
        /// <returns>this string will be use as history log</returns>
        public override string Execute()
        {
            // if SharePoint settings of configuration not set
            if (ContentReference.IsNullOrEmpty(SharePointSettings.Current.EntryPoint))
            {
                return string.Empty;
            }

            string message = _localizationService.GetString(LanguagePath + "/indexcompleted");
            try
            {
                var virtualPathNodes = ContentSearchHandler.GetVirtualPathNodes(SharePointSettings.Current.EntryPoint, ServiceLocator.Current.GetInstance<IContentRepository>());
                var reIndex = ServiceLocation.ServiceLocator.Current.GetInstance<SPContentReindex>();

                reIndex.RemoveItemsByVirtualPath(virtualPathNodes);
                reIndex.ReIndex();
            }
            
            catch (Exception ex)
            {
                message = string.Format("{0} {1}", _localizationService.GetString(LanguagePath + "/indexfailed"), ex.Message);
                _logger.Error("Failed to index SharePoint content", ex);
            }

            return message;
        }
    }
}