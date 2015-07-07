using EPiServer.Core;
using EPiServer.DataAbstraction;
using System;
namespace EPiServer.ConnectForSharePoint
{
    public interface ISharePointContentMappingRepository
    {
        /// <summary>
        /// Delete mapped identity of content link
        /// </summary>
        /// <param name="contentLink">The content reference of sharepointcontent provider</param>
        void Delete(ContentReference contentLink);

        /// <summary>
        /// Return content link of list/folder/file
        /// </summary>
        /// <param name="mappingKey">The unique key of each item include provider</param>
        ContentReference GetContentLink(string mappingKey);

        /// <summary>
        /// Return mapped identity based on the content link.
        /// </summary>
        /// <param name="contentLink">The content link.</param>
        MappedIdentity GetIdentity(ContentReference contentLink);

        /// <summary>
        /// Return mapped identity based on the content link.
        /// </summary>
        /// <param name="contentGuid">The content guid.</param>
        /// <returns></returns>
        MappedIdentity GetIdentity(Guid contentGuid);
    }
}
