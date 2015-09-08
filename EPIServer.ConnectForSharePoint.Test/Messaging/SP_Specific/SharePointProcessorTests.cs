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
using EPiServer.ConnectForSharePoint.SPServices;
using System.IO;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.Data.Dynamic;
using System.Linq;
using System.Linq.Expressions;
using EPiServer.Data;

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
        private Mock<IFileStreamService> fileStreamService;

        private SharePointProcessor GetUnitUnderTest()
        {
            contentSecurityRepo = new Mock<IContentSecurityRepository>();
            contentRepo = new Mock<IContentRepository>();
            contentTypeRepo = new Mock<IContentTypeRepository>();
            fileStreamService = new Mock<IFileStreamService>();

            // Initialize our service locator instance.
            serviceLocator = new Mock<IServiceLocator>();
            serviceLocator.Setup(sl => sl.GetInstance<IContentSecurityRepository>()).Returns(contentSecurityRepo.Object);
            serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);
            serviceLocator.Setup(sl => sl.GetInstance<IContentTypeRepository>()).Returns(contentTypeRepo.Object);
            ContentMediaResolver cmr = new ContentMediaResolver();
            serviceLocator.Setup(sl => sl.GetInstance<ContentMediaResolver>()).Returns(new MyContentMediaResolver());

            contentHelper = new Mock<IContentProviderHelper>();
            return new SharePointProcessor(serviceLocator.Object, contentHelper.Object, fileStreamService.Object);
        }

        [TestMethod]
        public void AddFolder_Processes_AddFolder_Message_And_Calls_CMS_To_Do_Add_With_EntryPointID()
        {
            var processor = GetUnitUnderTest();

            contentHelper.Setup(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.GetDefault<ContentFolder>(It.IsAny<ContentReference>()))
                .Returns(new ContentFolder());
            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Returns(new ContentReference(42));

            Mock<DynamicDataStoreFactory> ddsf = new Mock<DynamicDataStoreFactory>();
            Mock<MyDataStore> dds = new Mock<MyDataStore>();
            ddsf.Setup(fact => fact.CreateStore(It.IsAny<Type>())).Returns(dds.Object);
            dds.Setup(dd => dd.Save(It.IsAny<FolderData>())).Verifiable();
            FolderData folderData = new FolderData(ddsf.Object);

            var item = new SPListItem()
            {
                Folder = new SPFolder() { Name = "My Folder Name" },
                entryPointId = 123,
                GUID = new Guid("6178D0E1-4F19-4231-BC58-A433821A941B")
            };
            processor.AddFolder(item);

            ddsf.Verify(fact => fact.CreateStore(It.IsAny<Type>()), Times.Once, "Failed to create data store");
            dds.Verify(dd => dd.Save(It.IsAny<FolderData>()), Times.Once, "Failed to save the folder data");
            contentHelper.Verify(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>()), Times.Once, "SetAccessRights needs to be called once");
            contentRepo.Verify(cr => cr.Save(  It.Is<ContentFolder> (con => con.Name == item.Folder.Name && 
                con.ParentLink != null ),
                It.Is<SaveAction>(sa => sa == SaveAction.Publish), It.Is<AccessLevel>(al => al == AccessLevel.FullAccess)),
                Times.Once, "Save was called with unexpected values");
        }


        [TestMethod]
        public void AddFolder_Processes_AddFolder_Message_And_Calls_CMS_To_Do_Add_With_ParentFolderName()
        {
            var processor = GetUnitUnderTest();

            contentHelper.Setup(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.GetDefault<ContentFolder>(It.IsAny<ContentReference>())).Returns(new ContentFolder());
            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Returns(new ContentReference(42));

            Mock<DynamicDataStoreFactory> ddsf = new Mock<DynamicDataStoreFactory>();
            Mock<MyDataStore> dds = new Mock<MyDataStore>();
            ddsf.Setup(fact => fact.CreateStore(It.IsAny<Type>())).Returns(dds.Object);
            dds.Setup(dd => dd.Save(It.IsAny<FolderData>())).Verifiable();
            FolderData folderData = new FolderData(ddsf.Object);

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
                entryPointId = -1,
                GUID = new Guid("6178D0E1-4F19-4231-BC58-A433821A941B")
            };
            processor.AddFolder(item);

            ddsf.Verify(fact => fact.CreateStore(It.IsAny<Type>()), Times.Once, "Failed to create data store");
            dds.Verify(dd => dd.Save(It.IsAny<FolderData>()), Times.Once, "Failed to save the folder data");

            contentHelper.Verify(ch => ch.SetAccessRight(It.IsAny<ContentReference>(), It.IsAny<AccessLevel>()), Times.Once, "SetAccessRights needs to be called once");

            // Make sure the find recursive was called and the parent link id was set, and that the proper access control level was set.
            contentRepo.Verify(cr => cr.Save(It.Is<ContentFolder>(con => con.Name == item.Folder.Name &&
             con.ParentLink != null),
                It.Is<SaveAction>(sa => sa == SaveAction.Publish), It.Is<AccessLevel>(al => al == AccessLevel.FullAccess)),
                Times.Once, "Save was called with unexpected values");
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
                entryPointId = 42,
                GUID = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D")
            };

            // Return contentfolders list for getchildren
            var folders = new List<ContentFolder>();
            folders.Add(new ContentFolder() { Name = "Test 1", ContentLink = new ContentReference(1) });
            folders.Add(new ContentFolder() { Name = "Test 2", ContentLink = new ContentReference(2) });
            folders.Add(new ContentFolder() { Name = "FolderToBeDeleted", ContentLink = new ContentReference(42) });
            contentRepo.SetupSequence(cr => cr.GetChildren<ContentFolder>(It.IsAny<ContentReference>()))
                .Returns(folders)
                .Returns(new List<ContentFolder>());

            Mock<DynamicDataStoreFactory> ddsf = new Mock<DynamicDataStoreFactory>();
            Mock<MyDataStore> dds = new Mock<MyDataStore>();
            ddsf.Setup(fact => fact.CreateStore(It.IsAny<Type>())).Returns(dds.Object);
            dds.Setup(dd => dd.Delete(It.IsAny<FolderData>())).Verifiable();
            FolderData folderData = new FolderData(ddsf.Object);

            // Crazy but we need to return an ordered list of nothing...
            IList<FolderData> list = new List<FolderData>();
            dds.Setup(dd => dd.Items<FolderData>()).Returns(list.AsQueryable().OrderBy(fd => fd.Id));

            processor.DeleteFolder(item);

            // Verify that the code found the specified contentreference that matches the entry point in the entrypoint id
            // specified in the splistitem
            contentRepo.Verify(ch => ch.Delete(It.Is<ContentReference>(cf => cf.ID == 42),
                It.IsAny<bool>(), It.IsAny<AccessLevel>()), Times.Once,
                "Repository Delete needs to be called once with proper Content reference");
            dds.Verify(dd => dd.Delete(It.IsAny<Identity>()), Times.Once,
                "Delete dynamic data folder data must be called once.");
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
        public void AddDocument_Calls_FileStreamMethods()
        {
            var processor = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Returns(new ContentType() { ID = 4 });
            fileStreamService.Setup(fss => fss.GetFileRelativeUrl(It.IsAny<SPListItem>())).Verifiable();
            fileStreamService.Setup(fss => fss.GetFileRelativeUrl(It.IsAny<SPListItem>())).Returns("hereissomepath");

            MemoryStream mstream = new MemoryStream();
            StreamWriter myStreamWriter = new StreamWriter(mstream);
            myStreamWriter.Write("Some stream data");

            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Verifiable();
            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Returns(mstream);

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
                .Returns(new SharepointGenericMedia());

            processor.AddDocument(item);

            fileStreamService.Verify(fss => fss.GetFileSteam(It.IsAny<string>()), Times.Once, "GetFileSteam not called or thumbnail arg not false");
            fileStreamService.Verify(fss => fss.GetFileRelativeUrl(It.IsAny<SPListItem>()), Times.Once, "GetFileRelativeUrl or thumbnail arg not false");
            contentRepo.Verify(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>()), Times.Once, "Save needs to be called once");
        }

        public void UpdateDocument_Calls_ContentTypeRepository_Load()
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

            processor.UpdateDocument(item);

            contentTypeRepo.Verify(ctr => ctr.Load(It.IsAny<Type>()), Times.Once, "Load Media type must be called once");
            contentRepo.Verify(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>()), Times.Once, "Save needs to be called once");
        }

        [TestMethod]
        public void UpdateDocument_Calls_SearchForParentFolder()
        {
            var processor = GetUnitUnderTest();

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

            processor.UpdateDocument(item);

            contentTypeRepo.Verify(ctr => ctr.Load(It.IsAny<Type>()), Times.Once, "Load Media type must be called once");
            contentHelper.Verify(ch => ch.FindRecursiveFolders(It.IsAny<string>(),
                It.IsAny<ContentReference>()), Times.Once, "Find recersive Folders failed to be called.");
            contentHelper.Verify(ch => ch.FindRecursiveFolders(It.Is<string>(str => str == item.ParentFolderName),
                It.IsAny<ContentReference>()), Times.Once, "Find recersive Folders failed to be called with proper arguments.");
        }

        [TestMethod]
        public void UpdateDocument_Calls_FileStreamMethods()
        {
            var processor = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Returns(new ContentType() { ID = 4 });
            fileStreamService.Setup(fss => fss.GetFileRelativeUrl(It.IsAny<SPListItem>())).Verifiable();
            fileStreamService.Setup(fss => fss.GetFileRelativeUrl(It.IsAny<SPListItem>())).Returns("hereissomepath");
            contentHelper.Setup(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>())).Verifiable();
            contentHelper.Setup(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>())).Returns(new ContentReference(42));
            contentRepo.Setup(cr => cr.Get<MediaData>(It.IsAny<ContentReference>())).Returns( 
                new SharepointGenericMedia()
                {
                    Name = "Filename.jpg",
                    ContentLink = new ContentReference(42),
                    SharePointTimeLastModified = new DateTime(2014, 1, 1)
                }
            );

            MemoryStream mstream = new MemoryStream();
            StreamWriter myStreamWriter = new StreamWriter(mstream);
            myStreamWriter.Write("Some stream data");

            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Verifiable();
            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Returns(mstream);

            var item = new SPListItem()
            {
                File = new SPFile() { Name = "Filename.jpg", TimeLastModified = DateTime.Now },
                entryPointId = 42, GUID = new Guid("8906E84E-DCA7-45CC-A41B-B63BC49943A1")
            };

            var data = new List<MediaData>();
            data.Add(new MediaData() { Name = "Test 1", ContentLink = new ContentReference(1) });
            data.Add(new MediaData() { Name = "Test 2", ContentLink = new ContentReference(2) });
            data.Add(new SharepointGenericMedia()
            {
                Name = "Filename.jpg",
                ContentLink = new ContentReference(42),
                SharePointTimeLastModified = new DateTime(2014,1,1)
            });
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(data);

            contentRepo.Setup(cr => cr.GetDefault<MediaData>(It.IsAny<ContentReference>(), It.IsAny<int>()))
                .Returns(new MediaData());

            processor.UpdateDocument(item);

            contentHelper.Verify(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>()), Times.Once, "FindMediaInFolder needs to be called once");
            fileStreamService.Verify(fss => fss.GetFileSteam(It.IsAny<string>()), Times.Once, "GetFileSteam not called or thumbnail arg not false");
            fileStreamService.Verify(fss => fss.GetFileRelativeUrl(It.IsAny<SPListItem>()), Times.Once, "GetFileRelativeUrl or thumbnail arg not false");
            contentRepo.Verify(cr => cr.Save(It.IsAny<IContent>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>()), Times.Once, "Save needs to be called once");
        }

        [TestMethod]
        public void UpdateFolder_Calls_RepositorySave()
        {
            var processor = GetUnitUnderTest();
            contentRepo.Setup(cr => cr.Save(It.IsAny<ContentFolder>(), It.IsAny<SaveAction>(), It.IsAny<AccessLevel>())).Verifiable();
            contentRepo.Setup(cr => cr.Get<ContentFolder>(It.IsAny<ContentReference>())).Returns(
                new ContentFolder()
                {
                    Name = "SomeFolder",
                    ContentLink = new ContentReference(42),
                }
            );

            var item = new SPListItem()
            {
                Folder = new SPFolder() { Name = "FolderToBeUpdated" },
                entryPointId = 42,
                GUID = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D"),
                Modified = DateTime.Now
            };

            Mock<DynamicDataStoreFactory> ddsf = new Mock<DynamicDataStoreFactory>();
            Mock<MyDataStore> dds = new Mock<MyDataStore>();
            ddsf.Setup(fact => fact.CreateStore(It.IsAny<Type>())).Returns(dds.Object);
            dds.Setup(dd => dd.Save(It.IsAny<FolderData>())).Verifiable();
            FolderData folderData = new FolderData(ddsf.Object);

            // make sure our dynamic data store returns a ordered list with the folder data required.
            IList<FolderData> list = new List<FolderData>();
            dds.Setup(dd => dd.Items<FolderData>()).Returns(list.AsQueryable().OrderBy(fd => fd.Id));

            processor.UpdateFolder(item);

            // Verify that the code found the specified contentreference that matches the entry point in the entrypoint id
            // specified in the splistitem
            contentRepo.Verify(ch => ch.Save(It.Is<ContentFolder>(cf => cf.ContentLink.ID == 42),
                It.IsAny<SaveAction>(), It.IsAny<AccessLevel>()), Times.Once,
                "Repository Save needs to be called once with proper Content reference when updating");
            dds.Verify(dd => dd.Save(It.Is<FolderData>(fd => fd.Modified == item.Modified )), Times.Once,
                "Delete dynamic data folder data must be called once and the date must be updated");
        }

        [TestMethod]
        public void DeleteDocument_Calls_Delete()
        {
            var processor = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.Delete(It.IsAny<ContentReference>(), It.IsAny<Boolean>(), It.IsAny<AccessLevel>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Returns(new ContentType() { ID = 4 });
            contentHelper.Setup(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>())).Verifiable();
            contentHelper.Setup(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>())).Returns(
                new ContentReference(42));
            contentRepo.Setup(cr => cr.Get<MediaData>(It.IsAny<ContentReference>())).Returns(
                new SharepointGenericMedia()
                {
                    Name = "Filename.jpg",
                    ContentLink = new ContentReference(42),
                    SharePointTimeLastModified = new DateTime(2014, 1, 1)
                }
            );

            MemoryStream mstream = new MemoryStream();
            StreamWriter myStreamWriter = new StreamWriter(mstream);
            myStreamWriter.Write("Some stream data");

            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Verifiable();
            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Returns(mstream);

            var item = new SPListItem()
            {
                File = new SPFile() { Name = "Filename.jpg", TimeLastModified = DateTime.Now },
                entryPointId = 42,
                GUID = new Guid("8906E84E-DCA7-45CC-A41B-B63BC49943A1")
            };

            var data = new List<MediaData>();
            data.Add(new MediaData() { Name = "Test 1", ContentLink = new ContentReference(1) });
            data.Add(new MediaData() { Name = "Test 2", ContentLink = new ContentReference(2) });
            data.Add(new SharepointGenericMedia()
            {
                Name = "Filename.jpg",
                ContentLink = new ContentReference(42),
                SharePointTimeLastModified = new DateTime(2014, 1, 1)
            });
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(data);

            contentRepo.Setup(cr => cr.GetDefault<MediaData>(It.IsAny<ContentReference>(), It.IsAny<int>()))
                .Returns(new MediaData());

            processor.DeleteDocument(item);

            contentHelper.Verify(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>()), Times.Once, "FindMediaInFolder needs to be called once");
            contentRepo.Verify(cr => cr.Delete(It.IsAny<ContentReference>(), It.Is<Boolean>(force => force == true), 
                It.Is<AccessLevel>(al => al == AccessLevel.Delete)), Times.Once, "Save needs to be called once");
        }

        [TestMethod]
        public void DeleteDocument_Calls_Delete_AFter_FindRecursiveFolders()
        {
            var processor = GetUnitUnderTest();

            contentRepo.Setup(cr => cr.Delete(It.IsAny<ContentReference>(), It.IsAny<Boolean>(), It.IsAny<AccessLevel>())).Verifiable();
            contentTypeRepo.Setup(ct => ct.Load(It.IsAny<Type>())).Returns(new ContentType() { ID = 4 });
            contentHelper.Setup(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>())).Verifiable();
            contentHelper.Setup(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>())).Returns(
                new ContentReference(42));
            contentRepo.Setup(cr => cr.Get<MediaData>(It.IsAny<ContentReference>())).Returns(
                new SharepointGenericMedia()
                {
                    Name = "Filename.jpg",
                    ContentLink = new ContentReference(42),
                    SharePointTimeLastModified = new DateTime(2014, 1, 1)
                }
            );
            contentHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Verifiable();
            contentHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Returns(
                new ContentReference());


            MemoryStream mstream = new MemoryStream();
            StreamWriter myStreamWriter = new StreamWriter(mstream);
            myStreamWriter.Write("Some stream data");

            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Verifiable();
            fileStreamService.Setup(fss => fss.GetFileSteam(It.IsAny<string>())).Returns(mstream);

            var item = new SPListItem()
            {
                File = new SPFile() { Name = "Filename.jpg", TimeLastModified = DateTime.Now },
                entryPointId = -1,
                ParentFolderName = "ParentFolderName",
                GUID = new Guid("8906E84E-DCA7-45CC-A41B-B63BC49943A1") 
            };

            var data = new List<MediaData>();
            data.Add(new MediaData() { Name = "Test 1", ContentLink = new ContentReference(1) });
            data.Add(new MediaData() { Name = "Test 2", ContentLink = new ContentReference(2) });
            data.Add(new SharepointGenericMedia()
            {
                Name = "Filename.jpg",
                ContentLink = new ContentReference(42),
                SharePointTimeLastModified = new DateTime(2014, 1, 1)
            });
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(data);

            contentRepo.Setup(cr => cr.GetDefault<MediaData>(It.IsAny<ContentReference>(), It.IsAny<int>()))
                .Returns(new MediaData());

            processor.DeleteDocument(item);

            contentHelper.Verify(ch => ch.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>()), Times.Once, "Failed to call find recursive folders");
            contentHelper.Verify(ch => ch.FindMediaInFolder(It.IsAny<Guid>(), It.IsAny<ContentReference>()), Times.Once, "FindMediaInFolder needs to be called once");
            contentRepo.Verify(cr => cr.Delete(It.IsAny<ContentReference>(), It.Is<Boolean>(force => force == true),
                It.Is<AccessLevel>(al => al == AccessLevel.Delete)), Times.Once, "Save needs to be called once");
        }
    }
    public abstract class MyDataStore : DynamicDataStore
    {
        private static StoreDefinition storeDefiniton;

        public MyDataStore() : base(storeDefiniton)
        {

        }
        public MyDataStore(StoreDefinition storeDefiniton) : base(storeDefiniton)
        {
        }

        override public IOrderedQueryable<TResult> Items<TResult>()
        {
            Mock<IOrderedQueryable<TResult>> q = new Mock<IOrderedQueryable<TResult>>();
            IList<TResult> list = new List<TResult>();
            q.Setup(x => x.Where(It.IsAny<Expression<Func<TResult, bool>>>())).Returns(list.AsQueryable);

            return q.Object;
        }
    }
}
