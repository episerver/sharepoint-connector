using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Messaging;
using Moq;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Test.Messaging
{
    [TestClass]
    public class SharePointProcessFactoryTests
    {
        Mock<IServiceLocator> serviceLocator = new Mock<IServiceLocator>();
        Mock<IContentProviderHelper> contentHelper = new Mock<IContentProviderHelper>();
        Mock<IFileStreamService> fileStreamService = new Mock<IFileStreamService>();
        Mock<IListsViewsService> listViewService = new Mock<IListsViewsService>();

        private SharePointProcessorFactory GetUnitUnderTest()
        {
            return new SharePointProcessorFactory(serviceLocator.Object);
        }

        [TestMethod]
        public void GetSharePointProcessor_Verifys_Returns_A_SharePointProcessor()
        {
            var myFactory = GetUnitUnderTest();
            serviceLocator.Setup(sl => sl.GetInstance<ISharePointProcessor>()).Verifiable(); // Make sure it checks for the service
            serviceLocator.Setup(sl => sl.GetInstance<ISharePointProcessor>())
                .Returns((ISharePointProcessor)null);

            var aProcessor = myFactory.GetSharePointProcessor(contentHelper.Object, fileStreamService.Object, listViewService.Object);
            Assert.IsInstanceOfType(aProcessor, typeof(SharePointProcessor));

            serviceLocator.Verify(sl => sl.GetInstance<ISharePointProcessor>(), Times.Once, "ServiceLocator needs to be called at least once looking for ISharePointProcessor");
        }

        [TestMethod]
        public void GetSharePointProcessor_Verifys_Returns_A_Plugin()
        {
            var myFactory = GetUnitUnderTest();
            DemoProcessor processor = new DemoProcessor();

            serviceLocator.Setup(sl => sl.GetInstance<ISharePointProcessor>()).Verifiable(); // Make sure it checks for the service
            serviceLocator.Setup(sl => sl.GetInstance<ISharePointProcessor>())
                .Returns(processor);

            var aProcessor = myFactory.GetSharePointProcessor(contentHelper.Object, fileStreamService.Object, listViewService.Object);

            // Verify get instance was called.
            serviceLocator.Verify(sl => sl.GetInstance<ISharePointProcessor>(), Times.Once, "ServiceLocator needs to be called at least once looking for ISharePointProcessor");
            // Verify that we got an instance of DemoProcessor and not the default processor
            Assert.IsInstanceOfType(aProcessor, typeof(DemoProcessor));

        }
    }

    /// <summary>
    /// Temp class used for testing factory and detection of plugins.
    /// </summary>
    class DemoProcessor : ISharePointProcessor
    {
        public void AddDocument(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void AddFolder(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void AddListItem(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void DeleteDocument(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void DeleteFolder(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void DeleteListItem(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void UpdateDocument(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void UpdateFolder(SPListItem item)
        {
            throw new NotImplementedException();
        }

        public void UpdateListItem(SPListItem item)
        {
            throw new NotImplementedException();
        }
    }
}
