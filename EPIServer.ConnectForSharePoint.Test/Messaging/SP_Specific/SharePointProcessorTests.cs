using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ServiceLocation;
using EPiServer.DataAbstraction;
using EPiServer.Core;
using EPiServer.Security;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.DataAccess;

namespace EPiServer.ConnectForSharePoint.Test.Messaging.SP_Specific
{
    [TestClass]
    public class SharePointProcessorTests
    {
        private Mock<IServiceLocator> serviceLocator;
        private Mock<IContentSecurityRepository> contentSecurityRepo;
        private Mock<IContentProviderHelper> contentHelper;
        private Mock<IContentRepository> contentRepo;

        private SharePointProcessor GetUnitUnderTest()
        {
            serviceLocator = new Mock<IServiceLocator>();
            contentSecurityRepo = new Mock<IContentSecurityRepository>();
            contentRepo = new Mock<IContentRepository>();
            serviceLocator.Setup(sl => sl.GetInstance<IContentSecurityRepository>()).Returns(contentSecurityRepo.Object);
            serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);

            contentHelper = new Mock<IContentProviderHelper>();
            return new SharePointProcessor(serviceLocator.Object, contentHelper.Object);
        }

        [TestMethod]
        public void AddFolder_Processes_AddFolder_Message_And_Calls_CMS_To_Do_Add()
        {
            var processor = GetUnitUnderTest();

            contentHelper.Setup(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.GetDefault<ContentFolder>(It.IsAny<ContentReference>())).Returns(new ContentFolder());
            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();

            var item = new SPListItem()
            {
                Folder = new SPFolder() { Name = "My Folder Name" },
                entryPointId = 123
            };
            processor.AddFolder(item);

            contentHelper.Verify(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>()), Times.Once, "SetAccessRights needs to be called once");
            contentRepo.Verify(cr => cr.Save(It.Is<IContent>(con => con.Name == item.Folder.Name && con.ParentLink != null),
                It.Is<SaveAction>(sa => sa == SaveAction.Publish), It.Is<AccessLevel>(al => al == AccessLevel.FullAccess)),
                Times.Once,
                "Save was called with unexpected values");
        }

        


    }
}
