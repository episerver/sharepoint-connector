using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.SPServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class SharePointHelperTest
    {

        private Mock<IServicesHub> serviceHub;
        private SharePointHelper GetUnitUnderTest()
        {
            serviceHub = new Mock<IServicesHub>();
            return new SharePointHelper(serviceHub.Object);
        }
        [TestMethod]
        public void GetListId_Should_Return_Correct_Value()
        {
            var mappingId = "B54915F8-793E-425A-A2DA-602CCE00495A~CEE5CCEF-2880-4EC2-A48C-D34E31761373";
            var idListAndView = SharePointHelper.GetListAndViewIds(mappingId);

            Assert.AreEqual("B54915F8-793E-425A-A2DA-602CCE00495A", idListAndView[0]);
        }

        [TestMethod]
        public void GetViewId_Should_Return_Correct_Value()
        {
            var mappingId = "B54915F8-793E-425A-A2DA-602CCE00495A~CEE5CCEF-2880-4EC2-A48C-D34E31761373";
            var idListAndView = SharePointHelper.GetListAndViewIds(mappingId);

            Assert.AreEqual("CEE5CCEF-2880-4EC2-A48C-D34E31761373", idListAndView[1]);
        }

        [TestMethod]
        public void GetItems_Returns_The_Items_From_The_ServiceHub()
        {
            var helper = GetUnitUnderTest();
            var listViewsService = new Mock<IListsViewsService>();
            listViewsService.Setup(lvs => lvs.GetItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>()));
            serviceHub.SetupGet(sh => sh.SPListsViewsService).Returns(listViewsService.Object);

            var list = new SPList();
            var listView = new SPListView();
            var folder = new SPFolder();

            helper.GetItems(list, listView, folder);

            listViewsService.Verify(lvs => lvs.GetItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>()), Times.Once, "failed to call to the ListViewsService to get the Items");

        }
    }
}