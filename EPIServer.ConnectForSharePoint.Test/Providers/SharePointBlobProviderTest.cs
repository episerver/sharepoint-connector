using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Test.Providers
{
    [TestClass]
    public class SharePointBlobProviderTest
    {
        [TestMethod]
        public void CreateBlob_ShouldReturn_SharePointBlob_Type()
        {
            var provider = new SharePointBlobProvider();
            var blob = provider.CreateBlob(new Uri("http://sharepoint.com"), ".png");
            Assert.IsInstanceOfType(blob, typeof(SharePointBlob));
        }

        [TestMethod]
        public void GetBlob_ShouldReturn_SharePointBlob_Object()
        {
            var provider = new SharePointBlobProvider();
            var id = new Uri("http://sharepoint.com");
            var blob = provider.GetBlob(id);
            Assert.IsInstanceOfType(blob, typeof(SharePointBlob));
            Assert.AreEqual(blob.ID, id);
        }
    }
}
