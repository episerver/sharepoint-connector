using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using EPiServer.Web.Mvc;
using EPiServer.Web.Routing;
using System.Web.Mvc;

namespace EPiServer.ConnectForSharePoint.Controllers
{
    public class SharePointVideoController : PartialContentController<SharepointVideoFile>
    {
        private readonly UrlResolver _urlResolver;

        public SharePointVideoController(UrlResolver urlResolver)
        {
            _urlResolver = urlResolver;
        }

        /// <summary>
        /// The index action for the video file. Creates the view model and renders the view.
        /// </summary>
        /// <param name="currentContent">The current image file.</param>
        public override ActionResult Index(SharepointVideoFile currentContent)
        {
            var model = new VideoViewModel
            {
                Url = _urlResolver.GetUrl(currentContent.ContentLink),
                Name = currentContent.Name,
            };

            return PartialView("~/Views/Shared/SharePoint/Video.cshtml", model);
        }
    }
}