using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.SPServices;
using Moq;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using System.Collections.Generic;
using EPiServer.ServiceLocation;
using EPiServer.Core;
using EPiServer.Framework.Localization;
using System.Globalization;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Processor;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.ConnectForSharePoint.Messaging.InMemory;

namespace EPiServer.ConnectForSharePoint.Test
{
    public class MyLocalizationService : LocalizationService
    {
        public MyLocalizationService(ResourceKeyHandler keyHandler) : base(keyHandler)
        {
        }

        public override IEnumerable<CultureInfo> AvailableLocalizations
        {
            get {
                throw new NotImplementedException();
            }
        }

        protected override IEnumerable<ResourceItem> GetAllStringsByCulture(string originalKey, string[] normalizedKey, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        protected override string LoadString(string[] normalizedKey, string originalKey, CultureInfo culture)
        {
            return originalKey;
        }
    }

    [TestClass]
    public class UpdateJobTest
    {
        [TestMethod]
        public void Verify_ExecuteMethod_GetsAllLists_And_Calls_BeginsProcessing()
        {
            SharePointSettings._currentSettings = new SharePointSettings() {  EntryPoint = new ContentReference(42) };

            ContentFolder someContent = new ContentFolder() { Name = "thisone", IsDeleted = false };
            Mock<IContentRepository> repo = new Mock<IContentRepository>();
            repo.Setup(call => call.TryGet<ContentFolder>(It.IsAny<ContentReference>(), out someContent)).Returns(true);

            Mock<IServiceLocator> sl = new Mock<IServiceLocator>();
            ServiceLocator.SetLocator(sl.Object);
            sl.Setup(m => m.GetInstance<IContentRepository>()).Returns(repo.Object);
            sl.Setup(m => m.GetInstance<LocalizationService>()).Returns( new MyLocalizationService( new ResourceKeyHandler()) );

            Mock<IServiceLocator> locator = new Mock<IServiceLocator>();
            var itemList = new List<SPListItem>();
            var folderItem = new SPFolder() { Name = "Test Folder" };
            itemList.Add(new SPListItem() { Title = "Test Item 1", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList.Add(new SPListItem() { Title = "Test Item 2", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList.Add(new SPListItem() { Title = "Test Item 3", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });

            var itemList2 = new List<SPListItem>();
            itemList2.Add(new SPListItem() { Title = "Test Item 4", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList2.Add(new SPListItem() { Title = "Test Item 5", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList2.Add(new SPListItem() { Title = "Test Item 6", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });

            var listService = new Mock<IListsViewsService>();
            listService.Setup(sph => sph.GetLists(It.IsAny<int>())).Verifiable();
            listService.SetupSequence(sph => sph.GetLists(It.IsAny<int>()))
                .Returns( new List<SPList>() )
                .Returns(new List<SPList>());
            locator.Setup(m => m.GetInstance<IListsViewsService>()).Returns(listService.Object);

            Mock<IMetaDataProcessor> processor = new Mock<IMetaDataProcessor>();
            processor.Setup(h => h.BeginProcessing(It.IsAny<string>(),It.IsAny<IEnumerable<SPList>>())).Verifiable();
           
            UpdateJob job = new UpdateJob(locator.Object, processor.Object, "some list ids");
            var queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
            while(queueStore.Get(SharePointMessageEmitterFactory.QueueName).Count > 0)
            {
                queueStore.Get(SharePointMessageEmitterFactory.QueueName).Take();
            }
            job.Execute();

            listService.Verify(h => h.GetLists(It.IsAny<int>()), Times.Exactly(2), "GetLists should be called 2 with two different libs");
            listService.Verify(h => h.GetLists(It.Is<int>(i => i == (int)ListBaseType.DocumentLibrary )), Times.AtLeastOnce, "Never called with DocumentLibrary");
            listService.Verify(h => h.GetLists(It.Is<int>(i => i == (int)ListBaseType.GenericList)), Times.AtLeastOnce, "Never called with GenericList");
            processor.Verify(h => h.BeginProcessing(It.IsAny<string>(), It.IsAny<IEnumerable<SPList>>()), Times.Exactly(1), "Begin processing should be called once.");
        }
    }
}
