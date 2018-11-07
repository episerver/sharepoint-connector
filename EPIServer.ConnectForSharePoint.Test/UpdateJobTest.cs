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

            Mock<IServicesHub> hub = new Mock<IServicesHub>();
            var itemList = new List<SPListItem>();
            var folderItem = new SPFolder() { Name = "Test Folder" };
            itemList.Add(new SPListItem() { Title = "Test Item 1", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList.Add(new SPListItem() { Title = "Test Item 2", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList.Add(new SPListItem() { Title = "Test Item 3", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });

            var itemList2 = new List<SPListItem>();
            itemList2.Add(new SPListItem() { Title = "Test Item 4", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList2.Add(new SPListItem() { Title = "Test Item 5", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });
            itemList2.Add(new SPListItem() { Title = "Test Item 6", entryPointId = 123, Folder = folderItem, ParentFolderName = "Test Parent" });

            hub.Setup(sph => sph.SPListsViewsService.GetLists(It.IsAny<int>())).Verifiable();
            hub.SetupSequence(sph => sph.SPListsViewsService.GetLists(It.IsAny<int>()))
                .Returns( new List<SPList>() )
                .Returns(new List<SPList>());

            Mock<IMainProcessor> helper = new Mock<IMainProcessor>();
            helper.Setup(h => h.BeginProcessing(It.IsAny<string>(),It.IsAny<IEnumerable<SPList>>())).Verifiable();

            UpdateJob job = new UpdateJob(hub.Object, helper.Object, "some list ids");
            job.Execute();

            hub.Verify(h => h.SPListsViewsService.GetLists(It.IsAny<int>()), Times.Exactly(2), "GetLists should be called 2 with two different libs");

            hub.Verify(h => h.SPListsViewsService.GetLists(It.Is<int>(i => i == (int)ListBaseType.DocumentLibrary )), Times.AtLeastOnce, "Never called with DocumentLibrary");
            hub.Verify(h => h.SPListsViewsService.GetLists(It.Is<int>(i => i == (int)ListBaseType.GenericList)), Times.AtLeastOnce, "Never called with GenericList");
            helper.Verify(h => h.BeginProcessing(It.IsAny<string>(), It.IsAny<IEnumerable<SPList>>()), Times.Exactly(1), "Begin processing should be called once.");
        }
    }
}
