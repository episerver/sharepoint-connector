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
using System.Collections.Generic;
using System;

namespace EPiServer.ConnectForSharePoint.Test.Messaging.SP_Specific
{

    class MyContentMediaResolver : ContentMediaResolver
    {
        public override Type GetFirstMatching(string extension)
        {
            return new MediaData().GetType();
        }
    }

    [TestClass]
    public class SharePointProcessorTests
    {
        private Mock<IServiceLocator> serviceLocator;
        private Mock<IContentSecurityRepository> contentSecurityRepo;
        private Mock<IContentProviderHelper> contentHelper;
        private Mock<IContentRepository> contentRepo;
        private Mock<IContentTypeRepository> contentTypeRepo;

        private SharePointProcessor GetUnitUnderTest()
        {
            contentSecurityRepo = new Mock<IContentSecurityRepository>();
            contentRepo = new Mock<IContentRepository>();
            contentTypeRepo = new Mock<IContentTypeRepository>();

            // Initialize our service locator instance.
            serviceLocator = new Mock<IServiceLocator>();
            serviceLocator.Setup(sl => sl.GetInstance<IContentSecurityRepository>()).Returns(contentSecurityRepo.Object);
            serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);
            serviceLocator.Setup(sl => sl.GetInstance<IContentTypeRepository>()).Returns(contentTypeRepo.Object);
            ContentMediaResolver cmr = new ContentMediaResolver();
            serviceLocator.Setup(sl => sl.GetInstance<ContentMediaResolver>()).Returns(new MyContentMediaResolver());

            contentHelper = new Mock<IContentProviderHelper>();
            return new SharePointProcessor(serviceLocator.Object, contentHelper.Object);
        }

        [TestMethod]
        public void AddFolder_Processes_AddFolder_Message_And_Calls_CMS_To_Do_Add_With_EntryPointID()
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


        [TestMethod]
        public void AddFolder_Processes_AddFolder_Message_And_Calls_CMS_To_Do_Add_With_ParentFolderName()
        {
            var processor = GetUnitUnderTest();

            contentHelper.Setup(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.GetDefault<ContentFolder>(It.IsAny<ContentReference>())).Returns(new ContentFolder());
            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();

            contentHelper.Setup(ch => ch.FindRecursiveFolders(It.IsAny<string>(),
                It.IsAny<ContentReference>())).Returns(new ContentReference()
                {
                    ID = 10,
                    ProviderName = "Some Provider"
                });

            var item = new SPListItem()
            {
                Folder = new SPFolder() { Name = "My Folder Name" },
                ParentFolderName = "Parent Folder Name",
                entryPointId = -1
            };
            processor.AddFolder(item);

            contentHelper.Verify(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>()), Times.Once, "SetAccessRights needs to be called once");

            // Make sure the find recursive was called and the parent link id was set, and that the proper access control level was set.
            contentRepo.Verify(cr => cr.Save(It.Is<IContent>(con => con.Name == item.Folder.Name && con.ParentLink != null &&
                con.ParentLink.ID == 10),
                It.Is<SaveAction>(sa => sa == SaveAction.Publish), It.Is<AccessLevel>(al => al == AccessLevel.FullAccess)),
                Times.Once,
                "Save was called with unexpected values");
        }

        [TestMethod]
        public void DeleteFolder_Calls_SetAccessRight()
        {
            var processor = GetUnitUnderTest();
            contentHelper.Setup(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>())).Verifiable();

            var item = new SPListItem()
            {
                Folder = new SPFolder() { Name = "My Folder Name" },
                entryPointId = 42
            };

            processor.DeleteFolder(item);

            contentHelper.Verify(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>()), Times.Once, "SetAccessRights needs to be called once");
        }

        [TestMethod]
        public void DeleteFolder_Calls_Repository_Delete_When_Item_Is_Found()
        {
            var processor = GetUnitUnderTest();
            contentRepo.Setup(cr => cr.Delete(It.IsAny<ContentReference>(), It.IsAny<bool>(), It.IsAny<AccessLevel>())).Verifiable();
            var item = new SPListItem()
            {
                Folder = new SPFolder() { Name = "FolderToBeDeleted" },
                entryPointId = 42
            };

            // Return contentfolders list for getchildren
            var folders = new List<ContentFolder>();
            folders.Add(new ContentFolder() { Name = "Test 1", ContentLink = new ContentReference(1) });
            folders.Add(new ContentFolder() { Name = "Test 2", ContentLink = new ContentReference(2) });
            folders.Add(new ContentFolder() { Name = "FolderToBeDeleted", ContentLink = new ContentReference(42) });
            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()))
                .Returns(folders);

            processor.DeleteFolder(item);

            // Verify that the code found the specified contentreference that matches the entry point in the entrypoint id
            // specified in the splistitem
            contentRepo.Verify(ch => ch.Delete(It.Is<ContentReference>(cf => cf.ID == 42),
                It.IsAny<bool>(), It.IsAny<AccessLevel>()), Times.Once,
                "Delete needs to be called once with proper Content reference");
        }

        [TestMethod]
        public void AddDocument_Calls_ContentTypeRepository_Load()
        {
            var processor = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Returns(new ContentType() { ID = 4 });

            var item = new SPListItem()
            {
                File = new SPFile() { Name = "Filename.jpg" },
                entryPointId = 42
            };

            var data = new List<MediaData>();
            data.Add(new MediaData() { Name = "Test 1", ContentLink = new ContentReference(1) });
            data.Add(new MediaData() { Name = "Test 2", ContentLink = new ContentReference(2) });
            data.Add(new MediaData() { Name = "Filename.jpg", ContentLink = new ContentReference(42) });
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(data);

            contentRepo.Setup(cr => cr.GetDefault<MediaData>(It.IsAny<ContentReference>(), It.IsAny<int>()))
                .Returns(new MediaData());

            processor.AddDocument(item);

            contentTypeRepo.Verify(ctr => ctr.Load(It.IsAny<Type>()), Times.Once, "Load Media type must be called once");
            contentRepo.Verify(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>()), Times.Once, "Save needs to be called once");
        }

        [TestMethod]
        public void AddDocument_Calls_SearchForParentFolder()
        {
            var processor = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Returns(new ContentType() { ID = 4 });
            contentHelper.Setup(ch => ch.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Verifiable();

            var item = new SPListItem()
            {
                File = new SPFile() { Name = "Filename.jpg" },
                entryPointId = -1,
                ParentFolderName = "SomeParentFolder",
                rootCMSFolderEntryPointID = 4000
            };

            var data = new List<MediaData>();
            data.Add(new MediaData() { Name = "Test 1", ContentLink = new ContentReference(1) });
            data.Add(new MediaData() { Name = "Test 2", ContentLink = new ContentReference(2) });
            data.Add(new MediaData() { Name = "Filename.jpg", ContentLink = new ContentReference(42) });
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(data);

            // Return contentfolders list for getchildren
            var folders = new List<ContentFolder>();
            folders.Add(new ContentFolder() { Name = "Test 1", ContentLink = new ContentReference(1) });
            folders.Add(new ContentFolder() { Name = "Test 2", ContentLink = new ContentReference(2) });
            folders.Add(new ContentFolder() { Name = "SomeParentFolder", ContentLink = new ContentReference(42) });
            contentRepo.Setup(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()))
                .Returns(folders);

            contentRepo.Setup(cr => cr.GetDefault<MediaData>(It.IsAny<ContentReference>(), It.IsAny<int>()))
                .Returns(new MediaData());

            processor.AddDocument(item);

            contentTypeRepo.Verify(ctr => ctr.Load(It.IsAny<Type>()), Times.Once, "Load Media type must be called once");
            contentRepo.Verify(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>()), Times.Once, "Save needs to be called once");
            contentHelper.Verify(ch => ch.FindRecursiveFolders(It.IsAny<string>(),
                It.IsAny<ContentReference>()), Times.Once, "Find recersive Folders failed to be called.");
            contentHelper.Verify(ch => ch.FindRecursiveFolders(It.Is<string>(str => str == item.ParentFolderName),
                It.IsAny<ContentReference>()), Times.Once, "Find recersive Folders failed to be called with proper arguments.");
        }

        [TestMethod]
        [ExpectedException(typeof(NotImplementedException))]
        public void UpdateDocument_NotImplemented()
        {
            var processor = GetUnitUnderTest();
            processor.UpdateDocument(null);
        }

        [TestMethod]
        [ExpectedException(typeof(NotImplementedException))]
        public void UpdateFolder_NotImplemented()
        {
            var processor = GetUnitUnderTest();
            processor.UpdateFolder(null);
        }

        [TestMethod]
        [ExpectedException(typeof(NotImplementedException))]
        public void DeleteDocument_NotImplemented()
        {
            var processor = GetUnitUnderTest();
            processor.UpdateFolder(null);
        }
    }
}
