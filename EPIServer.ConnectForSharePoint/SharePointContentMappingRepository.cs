using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Helpers;
using System;

namespace EPiServer.ConnectForSharePoint
{
    /// <summary>
    /// Repository for manipulate with SharePointMediaMapping
    /// </summary>
    [ServiceConfiguration(typeof(ISharePointContentMappingRepository))]
    public class SharePointContentMappingRepository : ISharePointContentMappingRepository
    {
        private IdentityMappingService _identityMappingService;

        /// <summary>
        /// Initialize a instance of SharePointContentMappingRespository
        /// </summary>
        /// <param name="identityMappingService"></param>
        public SharePointContentMappingRepository()
            : this(ServiceLocator.Current.GetInstance<IdentityMappingService>())
        {
        }

        /// <summary>
        /// Initialize a instance of SharePointContentMappingRespository
        /// </summary>
        /// <param name="identityMappingService"></param>
        public SharePointContentMappingRepository(IdentityMappingService identityMappingService)
        {
            _identityMappingService = identityMappingService;
        }

        /// <summary>
        /// Delete mapped identity of content link
        /// </summary>
        /// <param name="contentLink">The content reference of sharepointcontent provider</param>
        public void Delete(ContentReference contentLink)
        {
            var mappedIdentity = _identityMappingService.Get(contentLink);
            if (mappedIdentity != null)
            {
                _identityMappingService.Delete(mappedIdentity.ExternalIdentifier);
            }
        }

        /// <summary>
        /// Return mapped identity of content
        /// </summary>
        /// <param name="contentGuid">The content guid of content</param>
        /// <returns>The mappped identity of content</returns>
        public MappedIdentity GetIdentity(Guid contentGuid)
        {
            return _identityMappingService.Get(contentGuid);
        }

        /// <summary>
        /// Return mapped identity of content
        /// </summary>
        /// <param name="contentLink">The content reference of content</contentLink>
        /// <returns>The mappped identity of content</returns>
        public MappedIdentity GetIdentity(ContentReference contentLink)
        {
            return _identityMappingService.Get(contentLink);
        }

        /// <summary>
        /// Return content link of list/folder/file
        /// </summary>
        /// <param name="mappingKey">The unique key of each item include provider</param>
        /// <returns>The content link</returns>
        public ContentReference GetContentLink(string mappingKey)
        {
            var mappedIdentity = _identityMappingService.Get(MappedIdentity.ConstructExternalIdentifier(ContentProviderHelper.ContentProviderKey, mappingKey), true);
            return mappedIdentity != null ? mappedIdentity.ContentLink : null;
        }
    }
}