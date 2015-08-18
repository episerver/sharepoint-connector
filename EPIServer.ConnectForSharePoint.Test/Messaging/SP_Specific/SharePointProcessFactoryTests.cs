using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Messaging;
using Moq;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices;

namespace EPiServer.ConnectForSharePoint.Test.Messaging
{
    [TestClass]
    public class SharePointProcessFactoryTests
    {
        Mock<IServiceLocator> serviceLocator = new Mock<IServiceLocator>();
        Mock<IContentProviderHelper> contentHelper = new Mock<IContentProviderHelper>();
        Mock<IFileStreamService> fileStreamService = new Mock<IFileStreamService>();
        private SharePointProcessorFactory GetUnitUnderTest()
        {
            return new SharePointProcessorFactory();
        }

        [TestMethod]
        public void GetSharePointProcessor_ReturnsASharePointProcessor()
        {
            var myFactory = GetUnitUnderTest();
            var aProcessor = myFactory.GetSharePointProcessor(serviceLocator.Object, contentHelper.Object, fileStreamService.Object);

            Assert.IsInstanceOfType(aProcessor, typeof(SharePointProcessor));
        }
    }
}
