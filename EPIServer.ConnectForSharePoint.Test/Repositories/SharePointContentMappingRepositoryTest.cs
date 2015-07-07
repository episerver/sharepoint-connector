using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Providers.Constructions;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using Moq;
using EPiServer.DataAbstraction;

namespace EPiServer.ConnectForSharePoint.Test
{
    [TestClass]
    public class SharePointContentMappingRepositoryTest
    {
        private Mock<IdentityMappingService> _mappingService;
        private SharePointContentMappingRepository _repository;
        private MappedIdentity mappedIdentity;

        private void DeleteMappedIdentity()
        {
            mappedIdentity = null;
        }

        [TestInitialize()]
        public void Initialize()
        {
            mappedIdentity = new MappedIdentity();
            var uri = MappedIdentity.ConstructExternalIdentifier("sharepointcontent", "episerver");
            _mappingService = new Mock<IdentityMappingService>();
            _mappingService.Setup(s => s.Get(uri, true)).Returns(mappedIdentity);
            _mappingService.Setup(s => s.Get(mappedIdentity.ContentGuid)).Returns(mappedIdentity);
            _mappingService.Setup(s => s.Get(mappedIdentity.ContentLink)).Returns(mappedIdentity);
            _mappingService.Setup(s => s.Delete(mappedIdentity.ExternalIdentifier)).Callback(DeleteMappedIdentity);

            _repository = new SharePointContentMappingRepository(_mappingService.Object);
        }

        [TestMethod]
        public void GetIdentity_By_ContentGuid_Return_MappedIdentity()
        {
            var newMapping = _repository.GetIdentity(Guid.NewGuid());
            Assert.IsNull(newMapping);

            var mapped = _repository.GetIdentity(mappedIdentity.ContentGuid);
            Assert.AreEqual(mappedIdentity.ContentGuid, mapped.ContentGuid);
            Assert.AreEqual(mappedIdentity.ContentLink, mapped.ContentLink);
            Assert.AreEqual(mappedIdentity.ExternalIdentifier, mapped.ExternalIdentifier);
        }

        [TestMethod]
        public void GetIdentity_By_ContentLink_Return_MappedIdentity()
        {
            var newMapping = _repository.GetIdentity(new ContentReference(1000));
            Assert.IsNull(newMapping);

            var mapped = _repository.GetIdentity(mappedIdentity.ContentLink);
            Assert.AreEqual(mappedIdentity.ContentGuid, mapped.ContentGuid);
            Assert.AreEqual(mappedIdentity.ContentLink, mapped.ContentLink);
            Assert.AreEqual(mappedIdentity.ExternalIdentifier, mapped.ExternalIdentifier);
        }

        [TestMethod]
        public void GetContentLink_By_MappingKey_Return_MappedIdentity()
        {
            var newMapping = _repository.GetContentLink(DateTime.Today.ToLongTimeString());
            Assert.AreEqual(newMapping, null);

            var contentLink = _repository.GetContentLink("episerver");
            Assert.AreEqual(mappedIdentity.ContentLink, contentLink);
        }
        
        [TestMethod]
        public void Delete_By_ContentLink()
        {
            _repository.Delete(mappedIdentity.ContentLink);
            Assert.AreEqual(mappedIdentity, null);
        }
    }
}
