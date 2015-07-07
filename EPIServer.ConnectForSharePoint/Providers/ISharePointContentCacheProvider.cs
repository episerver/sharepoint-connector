using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Providers
{
    /// <summary>
    /// Interface defining behaviors of a SharePoint content caching component.
    /// </summary>
    public interface ISharePointContentCacheProvider
    {

        /// <summary>
        /// Builds the mapping of SharePoint files/folders and adds to cache.
        /// </summary>
        void BuildMapping(ContentReference parentLink);

        /// <summary>
        /// Rebuilds the mapping of SharePoint files/folders and adds to cache.
        /// </summary>
        void RebuildMapping();

        /// <summary>
        /// Return a refresh list of supported media list
        /// </summary>
        IEnumerable<SPList> LatestMediaLists { get;}

        /// <summary>
        /// Return a list of supported media list
        /// </summary>
        IEnumerable<SPList> CachedMediaLists { get;}

        /// <summary>
        /// Gets the files stored in the cache.
        /// </summary>
        /// <remarks>
        /// The SharePoint files are stored in a dictionary with the key is the content link of content corresponding to the parent folder,
        /// the value is a list of <see cref="ModelWrapper"/> mapped with all files which are the direct children of the parent folder
        /// </remarks>
        Dictionary<string, List<IModelWrapper>> Files { get;}

        /// <summary>
        /// Gets the folders stored in the cache.
        /// </summary>
        /// <remarks>
        /// The SharePoint folders are stored in a dictionary with the key is the content link of content corresponding to the parent folder,
        /// the value is a list of <see cref="ModelWrapper"/> mapped with all folder which are the direct children of the parent folder
        /// </remarks>
        Dictionary<string, List<IModelWrapper>> Folders { get;}


        /// <summary>
        /// Removes a file from cache.
        /// </summary>
        /// <param name="contentLink">The content link associate with the file.</param>
        /// <param name="serverRelativeUrl">The server relative URL of the file.</param>
        void RemoveFileFromCache(ContentReference contentLink, string serverRelativeUrl);
    }
}