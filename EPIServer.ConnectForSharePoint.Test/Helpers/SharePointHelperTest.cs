using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ServiceLocation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class SharePointHelperTest
    {
        Mock<IServiceLocator> _locator;

        private SharePointHelper GetUnitUnderTest()
        {
            _locator = new Mock<IServiceLocator>();
            return new SharePointHelper(_locator.Object);
        }
        [TestMethod]
        public void GetListId_Should_Return_Correct_Value()
        {
            var helper = GetUnitUnderTest();
            var mappingId = "B54915F8-793E-425A-A2DA-602CCE00495A~CEE5CCEF-2880-4EC2-A48C-D34E31761373";
            var idListAndView = helper.GetListAndViewIds(mappingId);

            Assert.AreEqual("B54915F8-793E-425A-A2DA-602CCE00495A", idListAndView[0]);
        }

        [TestMethod]
        public void GetViewId_Should_Return_Correct_Value()
        {
            var helper = GetUnitUnderTest();
            var mappingId = "B54915F8-793E-425A-A2DA-602CCE00495A~CEE5CCEF-2880-4EC2-A48C-D34E31761373";
            var idListAndView = helper.GetListAndViewIds(mappingId);

            Assert.AreEqual("CEE5CCEF-2880-4EC2-A48C-D34E31761373", idListAndView[1]);
        }
    }
}