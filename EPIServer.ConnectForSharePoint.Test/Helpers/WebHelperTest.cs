using EPiServer.ConnectForSharePoint.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class WebHelperTest
    {
        [TestMethod]
        public void StripUrlSpecialCharacters_Should_Return_Correct_Value()
        {
            var originalName = "l@$%o&^!g~o.jpg";
            var strippedName = WebHelper.StripUrlSpecialCharacters(originalName);

            Assert.AreEqual("logo.jpg", strippedName);
        }
    }
}