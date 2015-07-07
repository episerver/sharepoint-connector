using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.ConnectForSharePoint.Providers.Constructions;
using EPiServer.Construction;
using EPiServer.Core;
using EPiServer.Core.Transfer;
using EPiServer.DataAbstraction;
using EPiServer.ServiceLocation;
using EPiServer.Validation;
using EPiServer.Web;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.Test.Providers
{
    [TestClass]
    public class SharePointMediaProviderTest
    {
        [TestMethod]
        public void LoadOrBuildContent_Should_Return_Null_If_There_Are_No_MappedItem_To_The_ContentLink()
        {
            var mockSharePointContentMappingRepository = new Mock<ISharePointContentMappingRepository>();
            mockSharePointContentMappingRepository.Setup(r => r.GetIdentity(It.IsAny<ContentReference>()))
                                                   .Returns<MappedIdentity>(null);
            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        null, //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        null, //ContentRepository,
                                                                        null, //SharePointDataBuilder,
                                                                        mockSharePointContentMappingRepository.Object,
                                                                        null,
                                                                        null, //EntryPoint
                                                                        null); //SharePointContentCacheProvider.Object);
            var result = sharePointMediaProvider.LoadOrBuildContent(new ContentReference(), null);
            Assert.IsNull(result);
        }

        [TestMethod]
        public void LoadOrBuildContent_Should_Invoke_CMSCache_To_Get_Content()
        {
            // Setup
            var contentMock = new Mock<IContent>();
            var mockSharePointContentMappingRepository = new Mock<ISharePointContentMappingRepository>();
            mockSharePointContentMappingRepository.Setup(r => r.GetIdentity(It.IsAny<ContentReference>()))
                                                   .Returns(new MappedIdentity());
            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        null, //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        null, //ContentRepository,
                                                                        null, //SharePointDataBuilder,
                                                                        mockSharePointContentMappingRepository.Object,
                                                                        (contentLink) => { return (object)contentMock.Object; },
                                                                        null, //EntryPoint
                                                                        null); //SharePointContentCacheProvider.Object);

            // Run
            var result = sharePointMediaProvider.LoadOrBuildContent(new ContentReference(), null);

            // Check
            Assert.AreEqual(contentMock.Object, result);
        }

        [TestMethod]
        public void LoadOrBuildContent_Should_Invoke_ContentCache_BuildMapping_Method_If_There_Is_No_Mapping()
        {
            // Setup
            var mockSharePointContentMappingRepository = new Mock<ISharePointContentMappingRepository>();
            mockSharePointContentMappingRepository.Setup(r => r.GetIdentity(It.IsAny<ContentReference>()))
                                                   .Returns(new MappedIdentity());

            var mockSharePointContentCacheProvider = new Mock<ISharePointContentCacheProvider>();
            mockSharePointContentCacheProvider.SetupGet(p => p.Folders).Returns(new Dictionary<string, List<IModelWrapper>>());
            mockSharePointContentCacheProvider.SetupGet(p => p.Files).Returns(new Dictionary<string, List<IModelWrapper>>());
            mockSharePointContentCacheProvider.Setup(p => p.BuildMapping(new ContentReference()));

            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        null, //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        null, //ContentRepository,
                                                                        null, //SharePointDataBuilder,
                                                                        mockSharePointContentMappingRepository.Object,
                                                                        (contentLink) => { return null; },
                                                                        null, //EntryPoint
                                                                        mockSharePointContentCacheProvider.Object); //SharePointContentCacheProvider.Object);

            // Run
            var result = sharePointMediaProvider.LoadOrBuildContent(new ContentReference(), null);

            // Check SharePointContentCacheProvider.BuildMapping is invoked.
            mockSharePointContentCacheProvider.Verify(mock => mock.BuildMapping(new ContentReference()), Times.Once());
            // Return null if listItem has deleted or unmap
            Assert.IsNull(null);
        }

        [TestMethod]
        [ExpectedException(typeof(NotSupportedException))]
        public void LoadOrBuildContent_Should_Throw_Exception_If_There_Are_No_ContentBuilders_For_The_ContentLink()
        {
            // Setup
            var mockSharePointContentMappingRepository = new Mock<ISharePointContentMappingRepository>();
            mockSharePointContentMappingRepository.Setup(r => r.GetIdentity(It.IsAny<ContentReference>()))
                                                   .Returns(new MappedIdentity());

            var mockContentReference = new Mock<ContentReference>();
            mockContentReference.Setup(r => r.Equals(It.IsAny<ContentReference>())).Returns(true);
            var mockModelWrapper = new Mock<IModelWrapper>();
            mockModelWrapper.SetupGet(w => w.ContentLink).Returns(mockContentReference.Object);

            var modelWrapperList = new List<IModelWrapper>() { mockModelWrapper.Object };
            var folders = new Dictionary<string, List<IModelWrapper>>();
            folders.Add("123123", modelWrapperList);

            var mockSharePointContentCacheProvider = new Mock<ISharePointContentCacheProvider>();
            mockSharePointContentCacheProvider.SetupGet(p => p.Folders).Returns(folders);
            mockSharePointContentCacheProvider.SetupGet(p => p.Files).Returns(new Dictionary<string, List<IModelWrapper>>());

            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        null, //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        null, //ContentRepository,
                                                                        Enumerable.Empty<ISharePointDataBuilder>(),
                                                                        mockSharePointContentMappingRepository.Object,
                                                                        (contentLink) => { return null; },
                                                                        null, //EntryPoint
                                                                        mockSharePointContentCacheProvider.Object); //SharePointContentCacheProvider.Object);

            // Run and check
            var result = sharePointMediaProvider.LoadOrBuildContent(mockContentReference.Object, null);
        }

        [TestMethod]
        public void LoadOrBuildContent_Should_Create_And_Return_A_Content()
        {
            // Setup
            var mockSharePointContentMappingRepository = new Mock<ISharePointContentMappingRepository>();
            mockSharePointContentMappingRepository.Setup(r => r.GetIdentity(It.IsAny<ContentReference>()))
                                                   .Returns(new MappedIdentity());

            var mockContentReference = new Mock<ContentReference>();
            mockContentReference.Setup(r => r.Equals(It.IsAny<ContentReference>())).Returns(true);
            var mockModelWrapper = new Mock<IModelWrapper>();
            mockModelWrapper.SetupGet(w => w.ContentLink).Returns(mockContentReference.Object);

            var modelWrapperList = new List<IModelWrapper>() { mockModelWrapper.Object };
            var folders = new Dictionary<string, List<IModelWrapper>>();
            folders.Add("123123", modelWrapperList);

            var mockSharePointContentCacheProvider = new Mock<ISharePointContentCacheProvider>();
            mockSharePointContentCacheProvider.SetupGet(p => p.Folders).Returns(folders);
            mockSharePointContentCacheProvider.SetupGet(p => p.Files).Returns(new Dictionary<string, List<IModelWrapper>>());

            var mockContentTypeRepository = new Mock<IContentTypeRepository>();
            mockContentTypeRepository.Setup(r => r.Load(It.IsAny<Type>())).Returns(new Mock<ContentType>().Object);

            var mockContentRepository = new Mock<IContentRepository>();
            mockContentRepository.Setup(r => r.GetDefault<IContent>(It.IsAny<ContentReference>(), It.IsAny<int>()))
                                                .Returns(new Mock<IContent>().Object);

            var mockFolderBuilder = new Mock<ISharePointDataBuilder>();
            mockFolderBuilder.Setup(b => b.HandleType(It.IsAny<Type>())).Returns(true);
            mockFolderBuilder.Setup(b => b.Build(It.IsAny<IContent>(), It.IsAny<ContentReference>(), It.IsAny<ContentReference>(), It.IsAny<IModelWrapper>()))
                                .Returns((IContent content, ContentReference contentLink, ContentReference parentLink, IModelWrapper modelWrapper) => content);
            var contentBuilders = new List<ISharePointDataBuilder>();
            contentBuilders.Add(mockFolderBuilder.Object);

            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        new ServiceAccessor<IContentTypeRepository>(() => mockContentTypeRepository.Object), //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        mockContentRepository.Object,
                                                                        contentBuilders,
                                                                        mockSharePointContentMappingRepository.Object,
                                                                        (contentLink) => { return null; },
                                                                        null, //EntryPoint
                                                                        mockSharePointContentCacheProvider.Object); //SharePointContentCacheProvider.Object);

            // Run and check
            var result = sharePointMediaProvider.LoadOrBuildContent(mockContentReference.Object, null);
            Assert.IsNotNull(result);
        }

        [TestMethod]
        public void LoadChildrenReferencesAndTypes_Should_Invoke_ContentCache_BuildMapping_Method_If_There_Is_No_Mapping()
        {
            var mockContentRepository = new Mock<IContentRepository>();
            mockContentRepository.Setup(r => r.Get<IContent>(It.IsAny<ContentReference>()));

            var mockSharePointContentCacheProvider = new Mock<ISharePointContentCacheProvider>();
            mockSharePointContentCacheProvider.SetupGet(p => p.Folders).Returns(new Dictionary<string, List<IModelWrapper>>());
            mockSharePointContentCacheProvider.SetupGet(p => p.Files).Returns(new Dictionary<string, List<IModelWrapper>>());
            mockSharePointContentCacheProvider.Setup(p => p.BuildMapping(new ContentReference()));

            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        null, //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        mockContentRepository.Object, //ContentRepository,
                                                                        null, //SharePointDataBuilder,
                                                                        null,
                                                                        null,
                                                                        () => { return ContentReference.EmptyReference; }, //EntryPoint
                                                                        mockSharePointContentCacheProvider.Object); //SharePointContentCacheProvider.Object);

            // Run
            bool languageSpecific;
            var result = sharePointMediaProvider.LoadChildrenReferencesAndTypes(ContentReference.EmptyReference, null, out languageSpecific);

            // Check SharePointContentCacheProvider.BuildMapping is invoked.
            mockSharePointContentCacheProvider.Verify(mock => mock.BuildMapping(new ContentReference()), Times.Once());
        }

        [TestMethod]
        public void LoadChildrenReferencesAndTypes_Should_Return_Correct_Value_For_ChildFolders()
        {
            var mockContentRepository = new Mock<IContentRepository>();
            mockContentRepository.Setup(r => r.Get<IContent>(It.IsAny<ContentReference>()));

            var mockContentReference = new Mock<ContentReference>();
            mockContentReference.Setup(r => r.Equals(It.IsAny<ContentReference>())).Returns(true);
            var mockModelWrapper = new Mock<IModelWrapper>();
            mockModelWrapper.SetupGet(w => w.ContentLink).Returns(mockContentReference.Object);

            var modelWrapperList = new List<IModelWrapper>() { mockModelWrapper.Object };
            var folders = new Dictionary<string, List<IModelWrapper>>();
            folders.Add("123123", modelWrapperList);

            var mockSharePointContentCacheProvider = new Mock<ISharePointContentCacheProvider>();
            mockSharePointContentCacheProvider.SetupGet(p => p.Folders).Returns(folders);
            mockSharePointContentCacheProvider.SetupGet(p => p.Files).Returns(new Dictionary<string, List<IModelWrapper>>());
            mockSharePointContentCacheProvider.Setup(p => p.BuildMapping(new ContentReference()));

            var sharePointMediaProvider = new SharePointMediaProviderWrapper("providerKey",
                                                                        null, //ContentFactory,
                                                                        null, //ServiceAccessor,
                                                                        null, //TemplateResolver,
                                                                        null, //UrlSegmentHandler,
                                                                        null, //LanguageBranchRepository,
                                                                        null, //FolderIDGenerator,
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        mockContentRepository.Object, //ContentRepository,
                                                                        null, //SharePointDataBuilder,
                                                                        null,
                                                                        null,
                                                                        () => { return ContentReference.EmptyReference; }, //EntryPoint
                                                                        mockSharePointContentCacheProvider.Object); //SharePointContentCacheProvider.Object);

            // Run
            bool languageSpecific;
            var result = sharePointMediaProvider.LoadChildrenReferencesAndTypes(new ContentReference(123123), null, out languageSpecific);

            // Check
            Assert.IsNotNull(result);
            Assert.AreEqual(1, result.Count());
        }

    }

    #region SharePointMediaProviderWrapper
    /// <summary>
    /// Wrapper class for SharePointMediaProvider to expose LoadOrBuildContent(private) and LoadChildrenReferencesAndTypes(protected) methods for testing.
    /// </summary>
    public class SharePointMediaProviderWrapper : SharePointMediaProvider
    {
        public SharePointMediaProviderWrapper(
            string providerKey,
            ContentFactory contentFactory,
            ServiceAccessor<IContentTypeRepository> serviceAccessor,
            TemplateResolver templateResolver,
            UrlSegmentHandler urlSegmentHandler,
            ILanguageBranchRepository languageBranchRepository,
            IFolderIDGenerator folderIDGenerator,
            IContentCopyHandler contentCopyHandler,
            IValidationService validationService,
            IContentEvents contentEvents,
            ContentMediaResolver contentMediaResolver,
            IContentRepository contentRepository,
            IEnumerable<ISharePointDataBuilder> sharePointDataBuilders,
            ISharePointContentMappingRepository sharePointRepository,
            Func<ContentReference, object> getCMSCachedContent,
            Func<ContentReference> getEntryPoint,
            ISharePointContentCacheProvider contentCache)
            : base(providerKey,
                contentFactory,
                serviceAccessor,
                templateResolver,
                urlSegmentHandler,
                languageBranchRepository,
                folderIDGenerator,
                contentCopyHandler,
                validationService,
                contentEvents,
                contentMediaResolver,
                contentRepository,
                sharePointDataBuilders,
                sharePointRepository,
                getCMSCachedContent,
                getEntryPoint,
                contentCache)
        { }

        public IContent LoadOrBuildContent(ContentReference contentLink, ILanguageSelector languageSelector)
        {
            return base.LoadContent(contentLink, languageSelector);
        }

        public new IList<GetChildrenReferenceResult> LoadChildrenReferencesAndTypes(ContentReference contentLink, string languageID, out bool languageSpecific)
        {
            return base.LoadChildrenReferencesAndTypes(contentLink, languageID, out languageSpecific);
        }
    }

    #endregion
}