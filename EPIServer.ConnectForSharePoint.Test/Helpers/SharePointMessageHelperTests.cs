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
            return new SharePointMessageHelper(_logger.Object, _helper.Object, _contentProviderHelper.Object, _settingsHelper.Object, _emitterFactory.Object);
        }
        [TestMethod]
        public void AddSubItems_Creates_Queue_Emits_Queue_Messages_For_Folders()
        {
            var messageHelper = GetUnitUnderTest();

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
            messageHelper.AddSubItems(pFolder, list, listView, folderName);

            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(spm => spm.Item.Title == item.Title && spm.Item.ParentFolderName == folderName)), Times.AtLeastOnce, "We didn't call emit or the folder was not the one we expect");
        }

        [TestMethod]
        public void AddSubItems_Creates_Queue_Emits_Queue_Messages_For_Items()
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
            messageHelper.AddSubItems(pFolder, list, listView);

            _messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(spm => spm.Item.ParentFolderName == pFolder.Name)), Times.AtLeastOnce, "We didn't call emit or the item was not the one we expect");
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

            messageHelper.QueueItemsForAdd(string.Empty, lists);

            _messageEmitter.Verify(me => me.Emit(It.IsAny<SharePointMessage>()), Times.Never, "Emit should be called with a folder add message for the root folder");
        }
    }
}
