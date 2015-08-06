using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Messaging;

namespace EPiServer.ConnectForSharePoint.Test.Messaging
{
    [TestClass]
    public class SharePointProcessFactoryTests
    {
        private SharePointProcessorFactory GetUnitUnderTest()
        {
            return new SharePointProcessorFactory();
        }

        [TestMethod]
        public void GetSharePointProcessor_ReturnsASharePointProcessor()
        {
            var myFactory = GetUnitUnderTest();
            var aProcessor = myFactory.GetSharePointProcessor();

            Assert.IsInstanceOfType(aProcessor, typeof(ISharePointProcessor));
        }
    }
}
