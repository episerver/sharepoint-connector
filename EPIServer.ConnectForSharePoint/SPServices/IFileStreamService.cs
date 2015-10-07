using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using System;
using System.IO;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IFileStreamService
    {
        /// <summary>
        /// Get file stream from SharePoint by its GUID string
        /// </summary>
        /// <param name="blobId">The given blob identification</param>
        /// <returns>File stream</returns>
        Stream GetFileStream(Uri blobId);

        /// <summary>
        /// Get file stream of a SharePoint's ListItem
        /// </summary>
        /// <param name="listItem">The file item from share point site</param>
        /// <param name="contentLink">The content link has been mapped identity</param>
        /// <returns></returns>
        Stream GetFileStream(SPListItem listItem, ContentReference contentLink);

        Stream GetFileSteam(string relativeUrl);

        string GetFileRelativeUrl(SPListItem listItem);

        Stream GetListItemXmlData(string url);
        string GetListItemJsonData(SPListItem item);
    }
}