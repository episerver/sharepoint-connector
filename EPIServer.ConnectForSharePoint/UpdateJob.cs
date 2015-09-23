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
            var lists = SharePointHelper.SPServicesHub.SPListsViewsService.GetLists((int)ListBaseType.DocumentLibrary);
            var messageHelper = new SharePointMessageHelper();
            // Put all the add folder messages on the queue
            messageHelper.QueueItemsForAdd(SharePointSettings.Current.ListIds, lists);
            return "Started";
        }
    }
}