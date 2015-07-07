using EPiServer.ConnectForSharePoint.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class SharePointHelperTest
    {
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
    }
}