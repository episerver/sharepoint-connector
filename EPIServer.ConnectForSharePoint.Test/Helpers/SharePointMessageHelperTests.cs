using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.Logging;
using EPiServer.ConnectForSharePoint.Messaging;
using Moq;
using EPiServer.ConnectForSharePoint.Models;
using System.Collections.Generic;
using EPiServer.Core;
using EPiServer.ConnectForSharePoint.Messaging.Messages;

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

        private SharePointMessageHelper GetUnitUnderTest()
        {
            _logger = new Mock<ILogger>();
            _helper = new Mock<ISharePointHelper>();
            _contentProviderHelper = new Mock<IContentProviderHelper>();
            _settingsHelper = new Mock<ISharePointSettingsHelper>();
            _emitterFactory = new Mock<IMessageEmitterFactory>();
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

            var messageEmitter = new Mock<IMessageEmitter>();
            messageEmitter.Setup(me => me.Emit(It.IsAny<SharePointMessage>()));
            _emitterFactory.Setup(ef => ef.Create()).Returns(messageEmitter.Object);

            var pFolder = new SPFolder();
            var list = new SPList();
            var listView = new SPListView();
            var folderName = "Test Folder";
            messageHelper.AddSubItems(pFolder, list, listView, folderName);

            messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(spm => spm.Item.Title == item.Title && spm.Item.ParentFolderName == folderName)), Times.AtLeastOnce, "We didn't call emit or the folder was not the one we expect");
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

            var messageEmitter = new Mock<IMessageEmitter>();
            messageEmitter.Setup(me => me.Emit(It.IsAny<SharePointMessage>()));
            _emitterFactory.Setup(ef => ef.Create()).Returns(messageEmitter.Object);

            var pFolder = new SPFolder() { Name = "Test Folder Name" };
            var list = new SPList();
            var listView = new SPListView();
            messageHelper.AddSubItems(pFolder, list, listView);

            messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(spm => spm.Item.ParentFolderName == pFolder.Name)), Times.AtLeastOnce, "We didn't call emit or the item was not the one we expect");
        }
    }
}
