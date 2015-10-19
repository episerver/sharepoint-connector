using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using EPiServer.Web.Mvc;
using EPiServer.Web.Routing;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Web.Mvc;

namespace EPiServer.ConnectForSharePoint.Controllers
{
    public class SharePointBlockDataController : BlockController<SharePointBlockData>
    {
        private readonly UrlResolver _urlResolver;
        private IServiceLocator _serviceLocator;

        internal SharePointBlockDataController(IServiceLocator locator) : base()
        {
            _serviceLocator = locator;
            _urlResolver = _serviceLocator.GetInstance<UrlResolver>();
        }

        [ExcludeFromCodeCoverage]
        public SharePointBlockDataController() : base()
        {
            _urlResolver = ServiceLocator.Current.GetInstance<UrlResolver>();
        }

        public override ActionResult Index(SharePointBlockData currentContent)
        {
            var content = currentContent as IContent;
            string str = currentContent.Json;

            var model = new ListItemViewModel
            {
                Name = content!= null ? content.Name : "DefaultName",
                JsonData = str,
                ListID = Guid.Parse(currentContent.ListID)
            };
            return PartialView("~/Views/Shared/SharePoint/ListItem.cshtml", model);
        }
    }
}