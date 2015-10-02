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
                                                                        null, //ContentCopyHandler,
                                                                        null, //ValidationService,
                                                                        null, //ContentEvents,
                                                                        null, //ContentMediaResolver,
                                                                        null, //ContentRepository,
                                                                        null, //SharePointDataBuilder,
                                                                        mockSharePointContentMappingRepository.Object,
                                                                        null,
                                                                        null); //EntryPoint
            var result = sharePointMediaProvider.LoadOrBuildContent(new ContentReference(), null);
            Assert.IsNull(result);
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
                IContentCopyHandler contentCopyHandler,
                IValidationService validationService,
                IContentEvents contentEvents,
                ContentMediaResolver contentMediaResolver,
                IContentRepository contentRepository,
                IEnumerable<ISharePointDataBuilder> sharePointDataBuilders,
                ISharePointContentMappingRepository sharePointRepository,
                Func<ContentReference, object> getCMSCachedContent,
                Func<ContentReference> getEntryPoint)
                : base(providerKey,
                    contentFactory,
                    serviceAccessor,
                    templateResolver,
                    urlSegmentHandler,
                    languageBranchRepository,
                    contentCopyHandler,
                    validationService,
                    contentEvents,
                    contentMediaResolver,
                    contentRepository,
                    sharePointDataBuilders,
                    sharePointRepository,
                    getCMSCachedContent,
                    getEntryPoint)
            { }

            public IContent LoadOrBuildContent(ContentReference contentLink, ILanguageSelector languageSelector)
            {
                return null; // base.LoadContent(contentLink, languageSelector);
            }
        }
    }
    #endregion
}