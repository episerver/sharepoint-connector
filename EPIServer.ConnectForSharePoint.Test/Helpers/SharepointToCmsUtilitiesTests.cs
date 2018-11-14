using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.Test.Messaging.SP_Specific;
using EPiServer.Core;
using EPiServer.Data.Dynamic;
using EPiServer.ServiceLocation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [ExcludeFromCodeCoverage]
    [TestClass]
    public class SharepointToCmsUtilitiesTests
    {
        private Mock<IServiceLocator> _serviceLocator;

        private Mock<IMessageEmitterFactory> _messageEmitterFactory;
        private Mock<MyDataStore> _dds;

        private SharepointToCmsUtilities GetUnitUnderTest()
        {
            _serviceLocator = new Mock<IServiceLocator>();
            _messageEmitterFactory = new Mock<IMessageEmitterFactory>();

            Mock<DynamicDataStoreFactory> ddsf = new Mock<DynamicDataStoreFactory>();
            _dds = new Mock<MyDataStore>();
            ddsf.Setup(fact => fact.CreateStore(It.IsAny<Type>())).Returns(_dds.Object);
            _dds.Setup(dd => dd.Save(It.IsAny<FolderData>())).Verifiable();
            var folderList = new List<FolderData>() { };
            _dds.Setup(dd => dd.Items<FolderData>()).Returns(folderList.AsQueryable().OrderBy(folder => folder.RefId));
            FolderData folderData = new FolderData(ddsf.Object);

            return new SharepointToCmsUtilities(_serviceLocator.Object);
        }

        [TestMethod]
        public void FindSupportedLists_Returns_Expected_Lists()
        {
            var messageHelper = GetUnitUnderTest();

            Mock<IListsViewsService> service = new Mock<IListsViewsService>();
            service.Setup(m => m.GetLists(It.Is<int>(v => v.Equals((int)ListBaseType.DocumentLibrary))))
                .Returns(new List<SPList>() { new SPList(), new SPList() });
            service.Setup(m => m.GetLists(It.Is<int>(v => v.Equals((int)ListBaseType.GenericList))))
                .Returns(new List<SPList>() { new SPList(), new SPList() });
            _serviceLocator.Setup(m => m.GetInstance<IListsViewsService>()).Returns(service.Object);

            var lists = messageHelper.FindSupportedLists();

            Assert.AreEqual(4, lists.Count());
            service.Verify(m => m.GetLists(It.Is<int>(v => v.Equals((int)ListBaseType.DocumentLibrary))));
            service.Verify(m => m.GetLists(It.Is<int>(v => v.Equals((int)ListBaseType.GenericList))));
            service.Verify(m => m.GetLists(It.IsAny<int>()), Times.Exactly(2), "We only support the specified types of lists at this time.");
        }

        [TestMethod]
        public void AddOrUpdateItem_DoesntProcess_FolderItem()
        {
            var messageHelper = GetUnitUnderTest();
            Mock<IEnumerable<IContent>> cmsItems = new Mock<IEnumerable<IContent>>();
            cmsItems.Setup(e => e.GetEnumerator()).Verifiable();

            SPListItem spItem = new SPListItem() { Folder = new SPFolder() };
            messageHelper.AddOrUpdateItem(_messageEmitterFactory.Object, spItem, cmsItems.Object);

            cmsItems.Verify(e => e.GetEnumerator(), Times.Never, "Item is not a folder item, we should not be iterating");

        }

        [TestMethod]
        public void AddOrUpdateItem_Finds_Items_UpdateItemMessage()
        {
            var messageHelper = GetUnitUnderTest();

            Guid itemGuid = new Guid("9A7961C7-4823-4B6B-9B37-829F85DC95CE");

            List<IContent> cmsItems = new List<IContent>();
            cmsItems.Add(new SharepointGenericMedia() { Name = "1" });
            cmsItems.Add(new SharepointGenericMedia()
            {
                Name = "FindMe",
                SharePointGUID = itemGuid.ToString(),
                ContentLink = new ContentReference(4242)
            });
            cmsItems.Add(new SharepointGenericMedia() { Name = "3" });

            SPListItem spItem = new SPListItem() { GUID = itemGuid };
            messageHelper.AddOrUpdateItem(_messageEmitterFactory.Object, spItem, cmsItems);

            _messageEmitterFactory.Verify( mf => mf.Emit<SharePointListItemUpdateMessage>(It.IsAny<SPListItem>()), Times.Exactly(1),
                "The item we where looking for should have been found AND emited update message");
        }

        [TestMethod]
        public void AddOrUpdateItem_Finds_Items_AddItemMessage()
        {
            var messageHelper = GetUnitUnderTest();

            Guid itemGuid = new Guid("9A7961C7-4823-4B6B-9B37-829F85DC95CE");

            List<IContent> cmsItems = new List<IContent>();
            cmsItems.Add(new SharepointGenericMedia() { Name = "1" });
            cmsItems.Add(new SharepointGenericMedia() { Name = "2" });
            cmsItems.Add(new SharepointGenericMedia() { Name = "3" });

            SPListItem spItem = new SPListItem() { GUID = itemGuid };
            messageHelper.AddOrUpdateItem(_messageEmitterFactory.Object, spItem, cmsItems);
            
            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointListItemAddMessage>(It.IsAny<SPListItem>()), Times.Exactly(1),
                "The item we where looking for should NOT have been found AND emited add message");
        }


        [TestMethod]
        public void AddOrUpdateItem_Finds_Items_UpdateContentMessage()
        {
            var messageHelper = GetUnitUnderTest();

            Guid itemGuid = new Guid("9A7961C7-4823-4B6B-9B37-829F85DC95CE");

            List<IContent> cmsItems = new List<IContent>();
            cmsItems.Add(new SharepointGenericMedia() { Name = "1" });
            cmsItems.Add(new SharepointGenericMedia()
            {
                Name = "FindMe",
                SharePointGUID = itemGuid.ToString(),
                ContentLink = new ContentReference(4242)
            });
            cmsItems.Add(new SharepointGenericMedia() { Name = "3" });

            SPListItem spItem = new SPListItem() { GUID = itemGuid, File = new SPFile() };
            messageHelper.AddOrUpdateItem(_messageEmitterFactory.Object, spItem, cmsItems);

            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointContentUpdateMessage>(It.IsAny<SPListItem>()), Times.Exactly(1),
                "The item we where looking for should have been found AND emited update message");
        }

        [TestMethod]
        public void AddOrUpdateItem_Finds_Items_AddContentMessage()
        {
            var messageHelper = GetUnitUnderTest();

            Guid itemGuid = new Guid("9A7961C7-4823-4B6B-9B37-829F85DC95CE");

            List<IContent> cmsItems = new List<IContent>();
            cmsItems.Add(new SharepointGenericMedia() { Name = "1" });
            cmsItems.Add(new SharepointGenericMedia() { Name = "2" });
            cmsItems.Add(new SharepointGenericMedia() { Name = "3" });

            SPListItem spItem = new SPListItem() { GUID = itemGuid, File = new SPFile() };
            messageHelper.AddOrUpdateItem(_messageEmitterFactory.Object, spItem, cmsItems);

            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointContentAddMessage>(It.IsAny<SPListItem>()), Times.Exactly(1),
                "The item we where looking for should NOT have been found AND emited add message");
        }

        [TestMethod]
        public void QueueFoldersForDelete_Does_Not_Call_Emit_For_Folders_Specified_As_Excluded()
        {
            var messageHelper = GetUnitUnderTest();
            var selectedListIds = new List<KeyValuePair<string, string>>();

            var viewKey = Guid.NewGuid().ToString();
            var listKey = Guid.NewGuid().ToString();
            selectedListIds.Add(new KeyValuePair<string, string>(viewKey, listKey));
            var testSPList = new SPList()
            {
                Id = new Guid(listKey),
                Title = "Test SPList",
                Views = new List<SPListView> { new SPListView() { Id = new Guid(viewKey), Title = "Test SPListView" } }
            };

            var lists = new List<SPList>() { testSPList };
            var excludeList = new List<string>() { listKey + "~" + viewKey };

            SharePointSettings._currentSettings = new SharePointSettings();

            messageHelper.QueueRootFoldersForDelete(_messageEmitterFactory.Object, lists, excludeList);

            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointMessage>(It.IsAny<SPListItem>()), 
                Times.Never, "Emit should not be called when we exclude the list-view combo returned");
        }

        [TestMethod]
        public void CreateParentSPListItem_CreatesExpected_SPListItem()
        {
            var messageHelper = GetUnitUnderTest();
            SharePointSettings._currentSettings = new SharePointSettings() { EntryPoint = new ContentReference(4242) };

            var list = new SPList() { Title = "List", Id = Guid.NewGuid() };
            var view = new SPListView() { Title = "View", Id = Guid.NewGuid() };
            var folderName = list.Title + " - " + view.Title;

            var item = messageHelper.CreateParentSPListItem(list, view);

            Assert.IsTrue(item.entryPointId == SharePointSettings._currentSettings.EntryPoint.ID);
            Assert.IsTrue(item.GUID == list.Id);
            Assert.IsTrue(folderName.Equals(item.Folder.Name));
            Assert.IsTrue(view.Title.Equals(item.Title));
        }

        [TestMethod]
        public void RemoveUserAddedFolders_Emmits_Expected_Message()
        {
            var messageHelper = GetUnitUnderTest();

            var startingFolderReference = new ContentReference(4242);
            var repo = new Mock<IContentRepository>();
            repo.Setup(m => m.GetChildren<ContentFolder>(It.IsAny<ContentReference>())).Returns(new List<ContentFolder>() {
                new ContentFolder() { Name = "Some Folder", ContentLink = new ContentReference(4242) }
            });
            _serviceLocator.Setup(m => m.GetInstance<IContentRepository>()).Returns(repo.Object);

            messageHelper.RemoveUserAddedFolders(_messageEmitterFactory.Object, startingFolderReference);
            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointFolderDeleteMessage>(It.IsAny<SPListItem>()), Times.Exactly(1),
                "The folder did not exist in the cms folder mapping store to sharepoint so it should have been deleted.");
        }

        [TestMethod]
        public void AddOrUpdateFolder_Emmits_Add_Folder()
        {
            var messageHelper = GetUnitUnderTest();

            var itemGuid = Guid.NewGuid();
            messageHelper.AddOrUpdateFolder(_messageEmitterFactory.Object, new SPListItem() { GUID = itemGuid });

            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointFolderAddMessage>(It.IsAny<SPListItem>()), Times.Exactly(1),
                "The folder did not exist in the cms folder mapping store so we will create it.");
            _messageEmitterFactory.Verify(mf => mf.Emit<SharePointFolderAddMessage>(It.Is<SPListItem>( item => item.GUID == itemGuid)),
                "The message did not contain the same item");
        }

    }
}
