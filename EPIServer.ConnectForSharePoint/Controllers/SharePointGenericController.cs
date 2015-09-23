using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using EPiServer.Web.Mvc;
using EPiServer.Web.Routing;
using System.Web.Mvc;

namespace EPiServer.ConnectForSharePoint.Controllers
{
    public class SharePointGenericController : PartialContentController<SharepointGenericMedia>
    {
        private readonly UrlResolver _urlResolver;

        public SharePointGenericController(UrlResolver urlResolver)
        {
            _urlResolver = urlResolver;
        }

        /// <summary>
        /// The index action for the image file. Creates the view model and renders the view.
        /// </summary>
        /// <param name="currentContent">The current image file.</param>
        public override ActionResult Index(SharepointGenericMedia currentContent)
        {
            var model = new GenericViewModel
            {
                Url = _urlResolver.GetUrl(currentContent.ContentLink),
                Name = currentContent.Name,
            };

            return PartialView("~/Views/Shared/SharePoint/Generic.cshtml", model);
        }

    }
}