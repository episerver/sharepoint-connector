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
        public void ProcessSubItems_Creates_Queue_Emits_Queue_Messages_For_Items()
        {
            var messageHelper = GetUnitUnderTest();

            var itemList = new List<SPListItem>();
            var fileItem = new SPFile() { Name = "Test Folder" };
            var item = new SPListItem() { Title = "Test Item One", entryPointId = 123, File = fileItem, ParentFolderName = string.Empty };
            itemList.Add(item);

            _helper.SetupSequence(sph => sph.GetItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>()))
                .Returns(itemList)
                .Returns(new List<SPListItem>());

            ContentReference test = null;
            _contentProviderHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Returns(test);

            var pFolder = new SPFolder() { Name = "Test Folder Name" };
            var list = new SPList();
            var listView = new SPListView();
            messageHelper.ProcessSubItems(pFolder, list, listView);

            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointContentAddMessage>(spm => spm.Item.ParentFolderName == pFolder.Name)), Times.AtLeastOnce, "We didn't call emit or the item was not the one we expect");
        }

        [TestMethod]
        public void ProcessSubItems_Emits_Delete_Items()
        {
            var messageHelper = GetUnitUnderTest();

            var sharepointListItems = new List<SPListItem>();
            // The list from sharepoint will have a folder and two items in it.
            sharepointListItems.Add(new SPListItem() { Folder = new SPFolder() { Name = "SomeFolder" } });
            var item = new SPListItem() { Title = "Test Item One", entryPointId = 121,
                File = new SPFile() { Name="test1"}, ParentFolderName = string.Empty,
                GUID = new Guid("4F39D06F-B540-466C-BD97-E3DE220034B1") };
            sharepointListItems.Add(item);
            item = new SPListItem() { Title = "Test Item three", entryPointId = 123,
                File = new SPFile() { Name = "test3" }, ParentFolderName = string.Empty,
                GUID = new Guid("4F39D06F-B540-466C-BD97-E3DE220034B3")
            };
            sharepointListItems.Add(item);
            _helper.SetupSequence(sph => sph.GetItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>()))
                .Returns(sharepointListItems);

            // Now make sure we return something not null indicating the folder was found and we need to check for items to delete
            _contentProviderHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>()))
                .Returns(new ContentReference()); // if its not null that means we found it and we are going to delete any items that are duplicate


            // Now setup the cms list, it will have 3 items that are not folders, the goal is delete the correct item that 
            // was removed from sharepoint
            // 
            Mock<IContentRepository> contentRepo = new Mock<IContentRepository>();
            _serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);
            // Return contentfolders list for getchildren
            var items = new List<MediaData>();
            items.Add(new SharepointGenericMedia() { Name = "test1", ContentLink = new ContentReference(1),
                SharePointGUID = "4F39D06F-B540-466C-BD97-E3DE220034B1" });
            items.Add(new SharepointGenericMedia() { Name = "ItemToBeDeleted", ContentLink = new ContentReference(2),
            SharePointGUID = "4F39D06F-B540-466C-BD97-E3DE220034B2"});
            items.Add(new SharepointGenericMedia() { Name = "test3", ContentLink = new ContentReference(42),
                SharePointGUID = "4F39D06F-B540-466C-BD97-E3DE220034B3" });
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(items);

            // none of the args actually matter because the mocks return test data
            messageHelper.ProcessSubItems(new SPFolder() { Name = "Test Folder Name" },
                new SPList(), new SPListView()); 

            // Comparer the quid and make sure its the correct one.
            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointContentDeleteMessage>(spm => spm.Item.GUID == new Guid("4F39D06F-B540-466C-BD97-E3DE220034B2"))), 
                Times.AtLeastOnce, "We didn't call emit or the item was not the one we expect");
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
            messageHelper.QueueFoldersForDelete(lists);

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
            messageHelper.QueueFoldersForDelete(lists, excludeList);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointMessage>()), Times.Never, "Emit should not be called when we exclude the list-view combo returned");
        }

        [TestMethod]
        public void QueueItemsForAdd_Emits_A_Message_To_Add_The_Root_Folder()
        {
            var messageHelper = GetUnitUnderTest();
            // make sure our dynamic data store returns a ordered list with the folder data required.
            IList<FolderData> ddslist = new List<FolderData>();
            ddslist.Add(new FolderData() { RefId = 42, FolderGuid = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D") });
            dds.Setup(dd => dd.Items<FolderData>()).Returns(ddslist.AsQueryable().OrderBy(fd => fd.Id));

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
                Views = new List<SPListView> { new SPListView() { Id = new Guid(viewKey), Title = "Test SPListView" } },
                RootFolder = new SPFolder()
            };

            var lists = new List<SPList>() { testSPList };

            ContentReference test = null;
            _contentProviderHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Returns(test);

            messageHelper.QueueItemsForAdd(string.Empty, lists);

            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(spm => spm.GetType() == typeof(SharePointFolderAddMessage))), Times.AtLeastOnce, "Emit should be called with a folder add message for the root folder");
        }

        [TestMethod]
        public void QueueItemsForAdd_Does_Not_Emit_A_Message_To_Add_If_The_Root_Exists()
        {
            var messageHelper = GetUnitUnderTest();

            // make sure our dynamic data store returns a ordered list with the folder data required.
            IList<FolderData> ddslist = new List<FolderData>();
            ddslist.Add(new FolderData() { RefId = 42, FolderGuid = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D") });
            dds.Setup(dd => dd.Items<FolderData>()).Returns(ddslist.AsQueryable().OrderBy(fd => fd.Id));

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
                Views = new List<SPListView> { new SPListView() { Id = new Guid(viewKey), Title = "Test SPListView" } },
                RootFolder = new SPFolder()
            };

            var lists = new List<SPList>() { testSPList };

            ContentReference test = new ContentReference();
            _contentProviderHelper.Setup(cph => cph.FindRecursiveFolders(It.IsAny<string>(), It.IsAny<ContentReference>())).Returns(test);

            var data = new List<MediaData>();
            data.Add(new MediaData() { Name = "Test 1", ContentLink = new ContentReference(1) });
            data.Add(new MediaData() { Name = "Test 2", ContentLink = new ContentReference(2) });
            data.Add(new SharepointGenericMedia()
            {
                Name = "Filename.jpg",
                ContentLink = new ContentReference(42),
                SharePointTimeLastModified = new DateTime(2014, 1, 1)
            });
            Mock<IContentRepository> contentRepo = new Mock<IContentRepository>();
            _serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);
            contentRepo.Setup(cr => cr.GetChildren<MediaData>(It.IsAny<ContentReference>()))
                .Returns(data);

            messageHelper.QueueItemsForAdd(string.Empty, lists);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointMessage>()), Times.Never, "Emit should be called with a folder add message for the root folder");
        }
    }
}
