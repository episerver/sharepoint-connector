using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using Moq;
using EPiServer.ConnectForSharePoint.Models;
using System.Collections.Generic;
using EPiServer.Core;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using System;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.Data.Dynamic;
using EPiServer.ConnectForSharePoint.Test.Messaging.SP_Specific;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class SharePointMessageHelperTests
    {

        private Mock<ILogger> _logger;
        private Mock<ISharePointHelper> _helper;
        private Mock<IContentProviderHelper> _contentProviderHelper;
        private Mock<ISharePointSettingsHelper> _settingsHelper;
        private Mock<IMessageEmitterFactory> _emitterFactory;
        private Mock<IMessageEmitter> _messageEmitter;
        private Mock<IServiceLocator> _serviceLocator;
        private Mock<MyDataStore> dds;

        private SharePointMessageHelper GetUnitUnderTest()
        {
            _logger = new Mock<ILogger>();
            _helper = new Mock<ISharePointHelper>();
            _contentProviderHelper = new Mock<IContentProviderHelper>();
            _settingsHelper = new Mock<ISharePointSettingsHelper>();
            _settingsHelper.Setup(sh => sh.GetEntryPoint()).Returns(new ContentReference() { ID = 123 });
            _messageEmitter = new Mock<IMessageEmitter>();
            _messageEmitter.Setup(me => me.Emit(It.IsAny<SharePointMessage>()));
            _emitterFactory = new Mock<IMessageEmitterFactory>();
            _emitterFactory.Setup(ef => ef.Create()).Returns(_messageEmitter.Object);

            Mock<DynamicDataStoreFactory> ddsf = new Mock<DynamicDataStoreFactory>();
            dds = new Mock<MyDataStore>();
            ddsf.Setup(fact => fact.CreateStore(It.IsAny<Type>())).Returns(dds.Object);
            dds.Setup(dd => dd.Save(It.IsAny<FolderData>())).Verifiable();
            FolderData folderData = new FolderData(ddsf.Object);

            _serviceLocator = new Mock<IServiceLocator>();
            return new SharePointMessageHelper(_logger.Object, _helper.Object, _contentProviderHelper.Object, _settingsHelper.Object, _emitterFactory.Object, _serviceLocator.Object);
        }
        [TestMethod]
        public void ProcessSubItems_Creates_Queue_Emits_Queue_Messages_For_Folders()
        {
            var messageHelper = GetUnitUnderTest();

            // make sure our dynamic data store returns a ordered list with the folder data required.
            IList<FolderData> ddslist = new List<FolderData>();
            ddslist.Add(new FolderData() { RefId = 42, FolderGuid = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D") });
            dds.Setup(dd => dd.Items<FolderData>()).Returns(ddslist.AsQueryable().OrderBy(fd => fd.Id));

            var itemList = new List<SPListItem>();
            var folderItem = new SPFolder() { Name = "Test Folder" };
            var item = new SPListItem() { Title = "Test Item One", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" };
            itemList.Add(item);

            _helper.SetupSequence(sph => sph.GetItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>()))
                .Returns(itemList)
                .Returns(new List<SPListItem>());

            ContentReference test = null;
            _contentProviderHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Returns(test);

            var pFolder = new SPFolder();
            var list = new SPList();
            var listView = new SPListView();
            var folderName = "Test Folder";
            messageHelper.ProcessSubItems(pFolder, list, listView, folderName);

            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(spm => spm.Item.Title == item.Title && spm.Item.ParentFolderName == folderName)), Times.AtLeastOnce, "We didn't call emit or the folder was not the one we expect");
        }

        [TestMethod]
        public void ProcessSubItems_Returns_NonFolders()
        {
            var messageHelper = GetUnitUnderTest();

            var itemList = new List<SPListItem>();
            var fileItem = new SPFile() { Name = "Test Folder" };
            var item = new SPListItem()
            {
                Title = "Test Item One",
                entryPointId = 123,
                File = fileItem,
                ParentFolderName = string.Empty
            };
            itemList.Add(item);
            itemList.Add(new SPListItem()
            {
                Title = "Test Item One",
                Folder = new SPFolder(),
                entryPointId = 123,
                ParentFolderName = string.Empty
            });
            _helper.SetupSequence(sph => sph.GetItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>()))
                .Returns(itemList)
                .Returns(new List<SPListItem>());

            IList<FolderData> ddslist = new List<FolderData>();
            ddslist.Add(new FolderData() { RefId = 42, FolderGuid = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D") });
            dds.Setup(dd => dd.Items<FolderData>()).Returns(ddslist.AsQueryable().OrderBy(fd => fd.Id));

            ContentReference test = null;
            _contentProviderHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Returns(test);

            var pFolder = new SPFolder() { Name = "Test Folder Name" };
            var list = new SPList();
            var listView = new SPListView();
            var items = messageHelper.ProcessSubItems(pFolder, list, listView);

            Assert.AreEqual(1, items.Count, "Item count does not match expected count.");
        }

        [TestMethod]
        public void QueueFoldersForDelete_Emits_Delete_Messages_For_Folders()
        {
            var messageHelper = GetUnitUnderTest();
            var selectedListIds = new List<KeyValuePair<string, string>>();

            var viewKey = Guid.NewGuid().ToString();
            var listKey = Guid.NewGuid().ToString();
            selectedListIds.Add(new KeyValuePair<string, string>(viewKey, listKey));
            _helper.Setup(h => h.GetSelectedListViewIds(It.IsAny<IEnumerable<SPList>>())).Returns(selectedListIds);
            _helper.Setup(h => h.GetListAndViewIds(It.IsAny<string>())).Returns(new string[2] { listKey, viewKey });

            var testSPList = new SPList()
            {
                Id = new Guid(listKey),
                Title = "Test SPList",
                Views = new List<SPListView> { new SPListView() { Id = new Guid(viewKey), Title = "Test SPListView" } }
            };

            var lists = new List<SPList>() { testSPList };
            messageHelper.QueueRootFoldersForDelete(lists);

            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(mess => mess.GetType() == typeof(SharePointFolderDeleteMessage))), Times.AtLeastOnce, "Emit was not called enough times with the expected message type");
        }

        [TestMethod]
        public void QueueFoldersForDelete_Does_Not_Call_Emit_For_Folders_Specified_As_Excluded()
        {
            var messageHelper = GetUnitUnderTest();
            var selectedListIds = new List<KeyValuePair<string, string>>();

            var viewKey = Guid.NewGuid().ToString();
            var listKey = Guid.NewGuid().ToString();
            selectedListIds.Add(new KeyValuePair<string, string>(viewKey, listKey));
            _helper.Setup(h => h.GetSelectedListViewIds(It.IsAny<IEnumerable<SPList>>())).Returns(selectedListIds);
            _helper.Setup(h => h.GetListAndViewIds(It.IsAny<string>())).Returns(new string[2] { listKey, viewKey });

            var testSPList = new SPList()
            {
                Id = new Guid(listKey),
                Title = "Test SPList",
                Views = new List<SPListView> { new SPListView() { Id = new Guid(viewKey), Title = "Test SPListView" } }
            };

            var lists = new List<SPList>() { testSPList };
            var excludeList = new List<string>() { listKey + "~" + viewKey };
            messageHelper.QueueRootFoldersForDelete(lists, excludeList);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointMessage>()), Times.Never, "Emit should not be called when we exclude the list-view combo returned");
        }

        [TestMethod]
        public void AddOrUpdateItem_DoesntProcess_FolderItem()
        {
            var messageHelper = GetUnitUnderTest();
            Mock<IEnumerable<IContent>> cmsItems = new Mock<IEnumerable<IContent>>();
            cmsItems.Setup(e => e.GetEnumerator()).Verifiable();

            SPListItem spItem = new SPListItem() { Folder = new SPFolder() };
            messageHelper.AddOrUpdateItem(spItem, cmsItems.Object);

            cmsItems.Verify(e => e.GetEnumerator(), Times.Never, "Item is not a folder item, we should not be iterating");

        }

        [TestMethod]
        public void AddOrUpdateItem_Finds_Items_UpdateItemMessage()
        {
            var messageHelper = GetUnitUnderTest();

            Guid itemGuid = new Guid("9A7961C7-4823-4B6B-9B37-829F85DC95CE");

            List<IContent> cmsItems = new List<IContent>();
            cmsItems.Add(new SharepointGenericMedia() { Name = "1" });
            cmsItems.Add(new SharepointGenericMedia() { Name = "FindMe", SharePointGUID = itemGuid.ToString(),
                ContentLink = new ContentReference(4242) });
            cmsItems.Add(new SharepointGenericMedia() { Name = "3" });

            SPListItem spItem = new SPListItem() { GUID = itemGuid };
            messageHelper.AddOrUpdateItem(spItem, cmsItems);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointListItemUpdateMessage>()), Times.Exactly(1),
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
            messageHelper.AddOrUpdateItem(spItem, cmsItems);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointListItemAddMessage>()), Times.Exactly(1),
                "The item we where looking for should NOT have been found AND emited add message");
        }


        [TestMethod]
        public void AddOrUpdateItem_Finds_Items_UpdateContentMessage()
        {
            var messageHelper = GetUnitUnderTest();

            Guid itemGuid = new Guid("9A7961C7-4823-4B6B-9B37-829F85DC95CE");

            List<IContent> cmsItems = new List<IContent>();
            cmsItems.Add(new SharepointGenericMedia() { Name = "1" });
            cmsItems.Add(new SharepointGenericMedia() { Name = "FindMe", SharePointGUID = itemGuid.ToString(),
                ContentLink = new ContentReference(4242) });
            cmsItems.Add(new SharepointGenericMedia() { Name = "3" });

            SPListItem spItem = new SPListItem() { GUID = itemGuid, File = new SPFile() };
            messageHelper.AddOrUpdateItem(spItem, cmsItems);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointContentUpdateMessage>()), Times.Exactly(1),
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
            messageHelper.AddOrUpdateItem(spItem, cmsItems);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointContentAddMessage>()), Times.Exactly(1),
                "The item we where looking for should NOT have been found AND emited add message");
        }
    }
}
