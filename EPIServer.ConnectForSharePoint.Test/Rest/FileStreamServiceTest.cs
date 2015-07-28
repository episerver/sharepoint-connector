using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.ConnectForSharePoint.Rest.Models;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using EPiServer.Core;
using EPiServer.Data.Dynamic;
using EPiServer.DataAbstraction;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;

namespace EPiServer.ConnectForSharePoint.Test
{

    [TestClass]
    public class FileStreamServiceTest
    {
        Mock<IListsViewsService> _spListsService;
        Mock<IFilesFoldersService> _spFilesFoldersService;
        Mock<ISharePointContentMappingRepository> _spContentMappingRepository;
        Mock<ISharePointContentCacheProvider> _spContentCacheProvider;

        Mock<FileStreamService> _serviceTesting;

        SPList _list;
        SPListItem _fileItem;
        MappedIdentity _mappedIdentity;

        [TestInitialize()]
        public void Initialize()
        {
            var ddsMock = new Mock<DynamicDataStore>(null);
            ddsMock.Setup(x => x.LoadAll<SharePointSettings>()).Returns(GetProviderSettings());
            var ddsFactoryMock = new Mock<DynamicDataStoreFactory>();
            ddsFactoryMock.Setup(x => x.GetStore(typeof(SharePointSettings))).Returns(ddsMock.Object);
            DynamicDataStoreFactory.Instance = ddsFactoryMock.Object;

            // Init List
            _list = new SPList
            {
                Id = Guid.NewGuid()
            };

            // Init File
            _fileItem = new SPListItem
            {
                FileUrl = "abc.gif",
                File = new SPFile()
            };

            // Init dependency and mock functions for MappedIdentity
            _mappedIdentity = new MappedIdentity();

            // Init dependency and mock functions for SPListsService
            _spListsService = new Mock<IListsViewsService>();
            _spListsService.Setup(s => s.GetList(It.IsAny<string>())).Returns(_list);
            _spListsService.Setup(s => s.GetListItem(It.IsAny<SPList>(), It.IsAny<string>())).Returns(_fileItem);

            // Init dependency and mock functions for SPFilesFoldersService
            _spFilesFoldersService = new Mock<IFilesFoldersService>();
            _spFilesFoldersService.Setup(s => s.LoadFile(It.IsAny<SPListItem>())).Returns(new SPFile());

            // Init dependency and mock functions for SPContentMappingRepository
            _spContentMappingRepository = new Mock<ISharePointContentMappingRepository>();
            _spContentMappingRepository.Setup(s => s.GetIdentity(It.IsAny<Guid>())).Returns(_mappedIdentity);
            _spContentMappingRepository.Setup(s => s.Delete(It.IsAny<ContentReference>()));

            // Init dependency and mock functions for SPContentCacheProvider 
            _spContentCacheProvider = new Mock<ISharePointContentCacheProvider>();
            _spContentCacheProvider.Setup(s => s.RemoveFileFromCache(It.IsAny<ContentReference>(), It.IsAny<string>()));

            // Mock for SendRequest method of base class            
            var jsonString = "Mock";
            _serviceTesting = new Mock<FileStreamService>(_spListsService.Object, _spFilesFoldersService.Object, _spContentMappingRepository.Object, _spContentCacheProvider.Object, new byte[1]) { CallBase = true };
            _serviceTesting.Setup(x => x.SendRequest(It.IsAny<IRestRequest>(), null, null)).Returns(jsonString);
            _serviceTesting.Setup(x => x.SendRequest(It.IsAny<IRestRequest>(), null, null, false)).Returns(jsonString);
            _serviceTesting.Setup(x => x.GetResponseStream(It.IsAny<IRestRequest>(), It.IsAny<string>(), It.IsAny<NameValueCollection>())).Returns(new MemoryStream());
            _serviceTesting.Setup(x => x.Deserialize<JsonFolder>(It.IsAny<string>())).Returns(new JsonFolder());
            _serviceTesting.Setup(x => x.Deserialize<JsonFile>(It.IsAny<string>())).Returns(MockJsonFile);

        }

        protected IEnumerable<SharePointSettings> GetProviderSettings()
        {
            return new List<SharePointSettings>() {
                new SharePointSettings
                {
                    Id = Data.Identity.NewIdentity(),
                    ThumbnailOption = ThumbnailOptions.GenericThumbnailForAll
                }
            };
        }

        /// <summary>
        /// Mock for Json file
        /// </summary>
        /// <returns></returns>
        private JsonFile MockJsonFile()
        {
            var f = new JsonFile();
            f.d = new JsonFile.D();
            f.d.Author =  new JsonUser().d;
            f.d.ModifiedBy = new JsonUser().d;
            return f;
        }

        [TestMethod]
        public void GetFileStream_with_uri_and_not_thumbnail_expected_not_null_and_length_equal_zero()
        {
            var url = string.Format("http://Test/{0}/abc.gif", Guid.NewGuid());
            var result = _serviceTesting.Object.GetFileStream(new Uri(url), false);
            Assert.IsNotNull(result);
            Assert.AreEqual(result.Length, 0); // Empty memory stream
        }

        [TestMethod]
        public void GetFileStream_with_SPListItem_not_have_url_expected_not_null_and_length_equal_zero()
        {
            _fileItem.File.ServerRelativeUrl = ""; // make sure its not null and length = zero.. not sure y we would have this test.

            var result = _serviceTesting.Object.GetFileStream(_fileItem, new ContentReference(100000), false);
            Assert.IsNotNull(result);
            Assert.AreEqual(result.Length, 0); // Empty memory stream
        }

        [TestMethod]
        public void GetFileStream_with_SPListItem_has_relative_url_expected_not_null_and_length_equal_zero()
        {
            _fileItem.File.ServerRelativeUrl = "http://test/abc.gif";

            var result = _serviceTesting.Object.GetFileStream(_fileItem, new ContentReference(100000), false);
            Assert.IsNotNull(result);
            Assert.AreEqual(result.Length, 0); // Empty memory stream
        }

        [TestMethod]
        public void GetFileStream_with_SPListItem_and_ContentReference_and_thumbnail_expected_not_null_and_length_equal_zero()
        {
            _fileItem.File.ServerRelativeUrl = "http://test/abc.gif";

            var result = _serviceTesting.Object.GetFileStream(_fileItem, new ContentReference(100000), true);
            Assert.IsNotNull(result);
            Assert.AreEqual(result.Length, 1); // Default thumbnail
        }
    }
}
