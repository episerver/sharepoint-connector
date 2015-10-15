using Moq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ServiceLocation;
using EPiServer.Core;
using EPiServer.Web;
using System.Collections.Generic;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using System;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class ContentProviderHelperTests
    {
        private Mock<IServiceLocator> serviceLocator;
        private Mock<IContentRepository> contentRepo;

        private ContentProviderHelper GetUnitUnderTest()
        {
            serviceLocator = new Mock<IServiceLocator>();
            contentRepo = new Mock<IContentRepository>();
            serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);

            return new ContentProviderHelper(serviceLocator.Object);
        }

        [TestMethod]
        public void FindRecursiveFolders_Uses_GlobalAssetsRoot()
        {
            var helper = GetUnitUnderTest();

            var folders = new List<ContentFolder>();
            folders.Add(new ContentFolder() { Name = "Test 1", ContentLink = new ContentReference(1) });
            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>())).Verifiable();
            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>())).Returns(folders);

            ContentReference cf = helper.FindRecursiveFolders("Test 1");

            contentRepo.Verify(cr => cr.GetChildren<ContentFolder>(It.Is<ContentReference>
                (f => f.ID == SiteDefinition.Current.GlobalAssetsRoot.ID)),
                    "SiteDefinition.Current.GlobalAssetsRoot was not used");
            contentRepo.Verify(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()), Times.Exactly(1),
                    "The method was not called just once");
        }

        [TestMethod]
        public void FindRecursiveFolders_CallsRecursively()
        {
            var helper = GetUnitUnderTest();

            var folders = new List<ContentFolder>();
            folders.Add(new ContentFolder() { Name = "Test 1", ContentLink = new ContentReference(1) });
            folders.Add(new ContentFolder() { Name = "Test 2", ContentLink = new ContentReference(2) });

            var folders2 = new List<ContentFolder>();
            folders2.Add(new ContentFolder() { Name = "Test 3", ContentLink = new ContentReference(3) });
            folders2.Add(new ContentFolder() { Name = "Test 4", ContentLink = new ContentReference(4) });
            folders2.Add(new ContentFolder() { Name = "Test 5", ContentLink = new ContentReference(5) });

            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>())).Verifiable();
            contentRepo.SetupSequence(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()))
                .Returns(folders)
                .Returns(folders2);

            ContentReference cf = helper.FindRecursiveFolders("Test 3");

            contentRepo.Verify(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()), Times.Exactly(2),
                    "The method must be called 2 times");
            Assert.AreEqual(cf.ID, 3, "Unexpected ID returned");
        }

        [TestMethod]
        public void FindRecursiveFolders_Returns_Null()
        {
            var helper = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>())).Verifiable();
            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()))
                .Returns(new List<ContentFolder>());

            var contentreference = helper.FindRecursiveFolders("Test 7");

            contentRepo.Verify(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()), Times.Exactly(1),
                    "The method must be called 1 times, the list is empty");
            Assert.IsNull(contentreference, "Expected null return value because item is not found");
        }

        [TestMethod]
        public void FindMediaInFolder_Returns_ContentLink()
        {
            var helper = GetUnitUnderTest();
            contentRepo.Setup(cr => cr.GetChildren<IContent>(It.IsAny<ContentReference>())).Verifiable();

            var itemGuid = new Guid("76B3BC47-01E8-4F6C-A07D-7F85976F5BE8");

            var folder_ref = new ContentReference() { ID = 42 };
            Mock<IItemForMoq> item = new Mock<IItemForMoq>();
            item.Setup(i => i.SharePointGUID).Returns(itemGuid.ToString());
            item.Setup(i => i.ContentLink).Returns(new ContentReference(4242));

            var items = new List<IItemForMoq>();
            items.Add(item.Object);
            contentRepo.Setup(cr => cr.GetChildren<IContent>(It.IsAny<ContentReference>())).Returns(items);

            var contentReference = helper.FindMediaInFolder(itemGuid, folder_ref);

            contentRepo.Verify(cr => cr.GetChildren<IContent>(It.Is<ContentReference>(c => c.ID == folder_ref.ID)),
                Times.Exactly(1),
                "The method must be called with the content reference passed in");
            Assert.AreEqual(contentReference.ID, 4242, "Item was not found");
        }

        [TestMethod]
        public void FindMediaInFolder_Returns_Null_IfNotFound()
        {
            var helper = GetUnitUnderTest();

            var itemGuid = new Guid("76B3BC47-01E8-4F6C-A07D-7F85976F5BE8");
            var folder_ref = new ContentReference() { ID = 42 };

            Mock<IItemForMoq> item = new Mock<IItemForMoq>();
            item.Setup(i => i.SharePointGUID).Returns(Guid.Empty.ToString()); // not the Guid your looking for
            item.Setup(i => i.ContentLink).Returns(new ContentReference(4242));

            var items = new List<IItemForMoq>();
            items.Add(item.Object);
            contentRepo.Setup(cr => cr.GetChildren<IContent>(It.IsAny<ContentReference>())).Returns(items);

            var contentReference = helper.FindMediaInFolder(itemGuid, folder_ref);
            Assert.IsNull(contentReference, "ContentReference should be null since the item could not be found");

        }
    }

    public interface IItemForMoq : IContent, ISharePointItem { }
}