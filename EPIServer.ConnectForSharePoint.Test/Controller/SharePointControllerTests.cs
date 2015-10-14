using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using EPiServer.Web.Routing;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using System.Web.Mvc;
using EPiServer.Core;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using EPiServer.ConnectForSharePoint.Controllers;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.Test.Controller
{
    [TestClass]
    public class SharePointControllerTests
    {
        [TestMethod]
        public void  SharePointBlockDataController_Returns_ProperUrl()
        {
            Mock<UrlResolver> urlResolver = new Mock<UrlResolver>();
            Mock<IServiceLocator> serviceLocator = new Mock<IServiceLocator>();
                serviceLocator.Setup(sl => sl.GetInstance<UrlResolver>()).Returns(urlResolver.Object);
            Mock<SharePointBlockData> item = new Mock<SharePointBlockData>();
                item.Setup(f => f.Json).Returns("JsonDAta");
                item.Setup(f => f.ListID).Returns("D299B990-E246-454F-B28B-D34630573DB2");

            SharePointBlockDataController controller = new SharePointBlockDataController(serviceLocator.Object);
            ActionResult ar = controller.Index(item.Object);
            Assert.IsInstanceOfType(ar, typeof(PartialViewResult));

            var vr = ar as PartialViewResult;
            Assert.IsInstanceOfType(vr.Model, typeof(ListItemViewModel), "model was not of correct type");

            var model = vr.Model as ListItemViewModel;
            Assert.AreEqual(model.Name, "DefaultName", false, "model name was unexpected");

            Assert.AreEqual(vr.ViewName, "~/Views/Shared/SharePoint/ListItem.cshtml", false, "view was unexpexted");
        }

        [TestMethod]
        public void SharePointGenericController_Returns_ProperUrl()
        {
            Mock<UrlResolver> urlResolver = new Mock<UrlResolver>();
            SharePointGenericController controller = new SharePointGenericController(urlResolver.Object);
            Mock<SharepointGenericMedia> item = new Mock<SharepointGenericMedia>();
            item.Setup(f => f.ContentLink).Returns(new ContentReference(42));
            item.Setup(f => f.Name).Returns("ThisThing");

            ActionResult ar = controller.Index(item.Object);
            Assert.IsInstanceOfType(ar, typeof(PartialViewResult));

            var vr = ar as PartialViewResult;
            Assert.IsInstanceOfType(vr.Model, typeof(GenericViewModel), "model was not of correct type");

            var model = vr.Model as GenericViewModel;
            Assert.AreEqual(model.Name, "ThisThing", false, "model name was unexpected");
            Assert.AreEqual(vr.ViewName, "~/Views/Shared/SharePoint/Generic.cshtml", false, "view was unexpexted");
        }

        [TestMethod]
        public void SharePointImageController_Returns_Proper_Url()
        {
            Mock<UrlResolver> urlResolver = new Mock<UrlResolver>();
            SharePointImageController controller = new SharePointImageController(urlResolver.Object);
            Mock<SharepointImageFile> item = new Mock<SharepointImageFile>();
                item.Setup(f => f.ContentLink).Returns(new ContentReference(42));
                item.Setup(f => f.Name).Returns("ThisThing");

            ActionResult ar = controller.Index(item.Object);
            Assert.IsInstanceOfType(ar, typeof(PartialViewResult));

            var vr = ar as PartialViewResult;
            Assert.IsInstanceOfType(vr.Model, typeof(ImageViewModel), "model was not of correct type");

            var model = vr.Model as ImageViewModel;
            Assert.AreEqual(model.Name, "ThisThing", false, "model name was unexpected");
            Assert.AreEqual(vr.ViewName, "~/Views/Shared/SharePoint/Image.cshtml", false, "view was unexpexted");
        }

        [TestMethod]
        public void SharePointVideoController_Returns_ProperUrl()
        {
            Mock<UrlResolver> urlResolver = new Mock<UrlResolver>();
            SharePointVideoController controller = new SharePointVideoController(urlResolver.Object);
            Mock<SharepointVideoFile> item = new Mock<SharepointVideoFile>();
            item.Setup(f => f.ContentLink).Returns(new ContentReference(42));
            item.Setup(f => f.Name).Returns("ThisThing");

            ActionResult ar = controller.Index(item.Object);
            Assert.IsInstanceOfType(ar, typeof(PartialViewResult));

            var vr = ar as PartialViewResult;
            Assert.IsInstanceOfType(vr.Model, typeof(VideoViewModel), "model was not of correct type");

            var model = vr.Model as VideoViewModel;
            Assert.AreEqual(model.Name, "ThisThing", false, "model name was unexpected");
            Assert.AreEqual(vr.ViewName, "~/Views/Shared/SharePoint/Video.cshtml", false, "view was unexpexted");
        }
    }
}
