using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Messaging;
using Moq;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.Test.Messaging
{
    [TestClass]
    public class SharePointProcessFactoryTests
    {
        Mock<IContentRepository> repo = new Mock<IContentRepository>();

        private SharePointProcessorFactory GetUnitUnderTest()
        {
            return new SharePointProcessorFactory();
        }

        [TestMethod]
        public void GetSharePointProcessor_ReturnsASharePointProcessor()
        {
            var myFactory = GetUnitUnderTest();
            var aProcessor = myFactory.GetSharePointProcessor(repo.Object);

            Assert.IsInstanceOfType(aProcessor, typeof(ISharePointProcessor));
        }
    }
}
