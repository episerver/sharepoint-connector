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
        /// Return a refresh list of supported media list
        /// </summary>
        IEnumerable<SPList> LatestMediaLists { get;}

        /// <summary>
        /// Return a list of supported media list
        /// </summary>
        IEnumerable<SPList> CachedMediaLists { get;}
    }
}