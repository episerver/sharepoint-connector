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
using EPiServer.DataAbstraction;
using EPiServer.ConnectForSharePoint.SPServices;

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
        private Mock<ISharePointProcessorHelper> _processorHelper;

        private Mock<IContentTypeRepository> _contentTypeRepo;
        private Mock<IContentModelUsage> _modelUsage;
        private Mock<IContentRepository> _contentRepo;
        private Mock<IFilesFoldersService> _fileFolderService;

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
            _processorHelper = new Mock<ISharePointProcessorHelper>();
            _serviceLocator.Setup(sl => sl.GetInstance<ISharePointProcessorHelper>()).Returns(_processorHelper.Object);

            _contentTypeRepo = new Mock<IContentTypeRepository>();
            _contentTypeRepo.Setup(m => m.List()).Returns(new List<ContentType>());
            _serviceLocator.Setup(sl => sl.GetInstance<IContentTypeRepository>()).Returns(_contentTypeRepo.Object);

            _modelUsage = new Mock<IContentModelUsage>();
            _modelUsage.Setup(m => m.ListContentOfContentType(It.IsAny<ContentType>())).Returns(new List<ContentUsage>());
            _serviceLocator.Setup(sl => sl.GetInstance<IContentModelUsage>()).Returns(_modelUsage.Object);

            _contentRepo = new Mock<IContentRepository>();
            _contentRepo.Setup(m => m.GetItems(It.IsAny<IEnumerable<ContentReference>>(), It.IsAny<LoaderOptions>())).Returns(new List<IContent>());
            _serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(_contentRepo.Object);

            _fileFolderService = new Mock<IFilesFoldersService>();
            _fileFolderService.Setup(m => m.LoadParentFolder(It.IsAny<SPFolder>())).Returns(new SPFolder() { Name = "Some Loaded Folder", ServerRelativeUrl = "root/ServerRelativeUrl" });
            _serviceLocator.Setup(sl => sl.GetInstance<IFilesFoldersService>()).Returns(_fileFolderService.Object);

            return new SharePointMessageHelper(_logger.Object, _helper.Object, _contentProviderHelper.Object, _settingsHelper.Object, _emitterFactory.Object, _serviceLocator.Object);

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
            cmsItems.Add(new SharepointGenericMedia()
            {
                Name = "FindMe",
                SharePointGUID = itemGuid.ToString(),
                ContentLink = new ContentReference(4242)
            });
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
            cmsItems.Add(new SharepointGenericMedia()
            {
                Name = "FindMe",
                SharePointGUID = itemGuid.ToString(),
                ContentLink = new ContentReference(4242)
            });
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

        string _listids;
        List<KeyValuePair<string, string>> _selectedListsAndViews;
        List<SPList> _lists;
        private void initBeginProcessingConditions()
        {
            _listids = "edf2847c-cab7-4842-96ee-d104c7da5c9b~16d6e901-152d-414d-aef4-6819abf701d2";
            _lists = new List<SPList>()
            {
                new SPList() { Title = "Test List 1", Id = new Guid("edf2847c-cab7-4842-96ee-d104c7da5c9b"),
                    Views = new List<SPListView>() { new SPListView() { Title = "ViewTitle", Id = new Guid("16d6e901-152d-414d-aef4-6819abf701d2") } } },
                new SPList() { Title = "Test List 2", Id = new Guid("edf2847c-cab7-4842-96ee-d104c7da5c9c"),
                    Views = new List<SPListView>() { new SPListView() { Title = "ViewTitle", Id = new Guid("16d6e901-152d-414d-aef4-6819abf701d2") } } },
                new SPList() { Title = "Test List 3", Id = new Guid("edf2847c-cab7-4842-96ee-d104c7da5c9d"),
                    Views = new List<SPListView>() { new SPListView() { Title = "ViewTitle", Id = new Guid("16d6e901-152d-414d-aef4-6819abf701d2") } } },
                new SPList() { Title = "Test List 4", Id = new Guid("edf2847c-cab7-4842-96ee-d104c7da5c9e"),
                    Views = new List<SPListView>() { new SPListView() { Title = "ViewTitle", Id = new Guid("16d6e901-152d-414d-aef4-6819abf701d2") } } }
            };

            _selectedListsAndViews = new List<KeyValuePair<string, string>>()
            {
                new KeyValuePair<string, string>(_listids, "ddf2847c-cab7-4842-96ee-d104c7da5c9b"),
                new KeyValuePair<string, string>(_listids, "ddf2847c-cab7-4842-96ee-d104c7da5c9b"),
                new KeyValuePair<string, string>(_listids, "ddf2847c-cab7-4842-96ee-d104c7da5c9b")
            };
            _helper.Setup(m => m.GetSelectedListViewIds(It.IsAny<IEnumerable<SPList>>())).Returns(_selectedListsAndViews);
            _helper.Setup(m => m.GetListAndViewIds(It.IsAny<string>())).Returns(_listids.Split('~'));

            IList<FolderData> ddslist = new List<FolderData>();
            ddslist.Add(new FolderData() { RefId = 42, FolderGuid = new Guid("6178D0E1-4F19-4231-BC58-A433821A941D") });
            dds.Setup(dd => dd.Items<FolderData>()).Returns(ddslist.AsQueryable().OrderBy(fd => fd.Id));
        }

        [TestMethod]
        public void BeginProcessing_clears_item_lists_with_each_selected_List()
        {
            var messageHelper = GetUnitUnderTest();

            initBeginProcessingConditions();
            Mock<IList<SPListItem>> folderList = new Mock<IList<SPListItem>>();
            Mock<IList<SPListItem>> itemList = new Mock<IList<SPListItem>>();
            messageHelper._ListSPFolders = folderList.Object;
            messageHelper._ListSPItems = itemList.Object;

            string Path = "";
            _helper.Setup(m => m.GetPagedItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>(), out Path)).Returns(
                new List<SPListItem>());

            messageHelper.BeginProcessing(_listids, _lists);

            // _messageEmitter.Verify(me => me.Emit(It.Is<SharePointMessage>(mess => mess.GetType() == typeof(SharePointFolderDeleteMessage))), Times.AtLeastOnce, "Emit was not called enough times with the expected message type");
            folderList.Verify(m => m.Clear(), Times.Exactly(_selectedListsAndViews.Count()), "Failed to clear folderList at start of new list iteration.");
            itemList.Verify(m => m.Clear(), Times.Exactly(_selectedListsAndViews.Count()), "Failed to clear itemList at start of new list iteration.");
        }

        [TestMethod]
        public void BeginProcessing_Calls_GetPagedItems()
        {
            var messageHelper = GetUnitUnderTest();

            initBeginProcessingConditions();

            string nextUriPath = "";
            _helper.Setup(m => m.GetPagedItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>(), out nextUriPath)).Returns(
                new List<SPListItem>());

            messageHelper.BeginProcessing(_listids, _lists);
            _helper.Verify(m => m.GetPagedItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>(), out nextUriPath),
                Times.Exactly(_selectedListsAndViews.Count()), "Failed to call GetPagedItems " + _selectedListsAndViews.Count() + " times.");
            _helper.Verify(m => m.GetNextPagedItems(It.IsAny<SPList>(), It.IsAny<string>(), out nextUriPath),
                Times.Never, "Unexpected call to GetNextPagedItems, there was no nextUriPath specified.");
        }

        [TestMethod]
        public void BeginProcessing_Calls_GetNextPagedItems()
        {
            var messageHelper = GetUnitUnderTest();

            initBeginProcessingConditions();

            string nextUriPath = " yo yo yo";
            _helper.Setup(m => m.GetPagedItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>(), out nextUriPath))
                .Callback(() => nextUriPath = "Next Path")
                .Returns(new List<SPListItem>());
            string nextUriPath2 = "";
            _helper.Setup(m => m.GetNextPagedItems(It.IsAny<SPList>(), It.IsAny<string>(), out nextUriPath2))
                .Callback(() => nextUriPath = null)
                .Returns(new List<SPListItem>());

            messageHelper.BeginProcessing(_listids, _lists);

            _helper.Verify(m => m.GetPagedItems(It.IsAny<SPList>(), It.IsAny<SPListView>(), It.IsAny<SPFolder>(), out nextUriPath),
                Times.Exactly(_selectedListsAndViews.Count()), "Failed to call GetPagedItems " + _selectedListsAndViews.Count() + " times.");
            _helper.Verify(m => m.GetNextPagedItems(It.IsAny<SPList>(), It.IsAny<string>(), out nextUriPath),
                Times.Exactly(_selectedListsAndViews.Count()), "Failed to call GetNextPagedItems " + _selectedListsAndViews.Count() + " times.");
        }

        [TestMethod]
        public void processItemList_Adds_Folder_To_FolderList()
        {
            var messageHelper = GetUnitUnderTest();
            var rootFolder = new SPListItem()
            {
                Title = "Root Folder",
                GUID = new Guid("eeeeeeee-cab7-4842-96ee-d104c7da5c9b"),
                Folder = new SPFolder() { Name = "Root Folder" }
            };
            var listitems = new List<SPListItem>() { new SPListItem() { Title = "Sub Folder", GUID = new Guid("ffffffff-cab7-4842-96ee-d104c7da5c9b"),
                    Folder = new SPFolder() { Name = "Sub Folder"  } } };

            initBeginProcessingConditions();

            messageHelper.processItemList(new SPList()
            {
                Title = "Test List 1",
                Id = new Guid("edf2847c-cab7-4842-96ee-d104c7da5c9b"),
                Views = new List<SPListView>() { new SPListView() { Title = "ViewTitle", Id = new Guid("16d6e901-152d-414d-aef4-6819abf701d2") } }
            }, rootFolder, listitems);

            _fileFolderService.Verify(m => m.LoadParentFolder(It.IsAny<SPFolder>()), "processItemList did not attempt to load the parent folder");
            Assert.IsTrue(messageHelper._ListSPFolders.Contains(listitems[0]), "Folder was not added to internal folder list.");
        }

        [TestMethod]
        public void processItemList_Adds_Item_To_ItemList()
        {
            var messageHelper = GetUnitUnderTest();
            var rootFolder = new SPListItem()
            {
                Title = "Root Folder",
                GUID = new Guid("eeeeeeee-cab7-4842-96ee-d104c7da5c9b"),
                Folder = new SPFolder() { Name = "Root Folder" }
            };
            var listitems = new List<SPListItem>() {
                    new SPListItem() { Title = "Some Item", GUID = new Guid("ffffffff-cab7-4842-96ee-d104c7da5c9b"),
                    File = new SPFile() { Name = "Some Item", ServerRelativeUrl = "root/asdf/asdf/Some Item"  }} } ;

            initBeginProcessingConditions();

            messageHelper._listSPItemsInCMs = new List<IContent>();

            messageHelper.processItemList(new SPList()
            {
                Title = "Test List 1",
                Id = new Guid("edf2847c-cab7-4842-96ee-d104c7da5c9b"),
                Views = new List<SPListView>() { new SPListView() { Title = "ViewTitle", Id = new Guid("16d6e901-152d-414d-aef4-6819abf701d2") } }
            }, rootFolder, listitems);

            Assert.IsTrue(messageHelper._ListSPItems.Contains(listitems[0]), "Item was not added to internal item list.");
        }
    }
}
