using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.PlugIn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint
{
    [ScheduledPlugIn(DisplayName = "SharePoint Update", Description = "")]
    public class UpdateJob
    {
        public static string Execute()
        {
            // Concatinate the document and custom lists.
            var documentLists = SharePointHelper.SPServicesHub.SPListsViewsService
                .GetLists((int)ListBaseType.DocumentLibrary);
            var customLists = SharePointHelper.SPServicesHub.SPListsViewsService
                .GetLists((int)ListBaseType.GenericList);
            var supportedLists = documentLists.Concat(customLists);

            // Put all the add folder messages on the queue
            var messageHelper = new SharePointMessageHelper();
            messageHelper.BeginProcessing(SharePointSettings.Current.ListIds, supportedLists);
            return "Started";
        }
    }
}