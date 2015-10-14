using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using System;
using System.IO;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IFileStreamService
    {
        /// <summary>
        /// Get file stream of a SharePoint's ListItem
        /// </summary>
        /// <param name="listItem">The file item from share point site</param>
        /// <param name="contentLink">The content link has been mapped identity</param>
        /// <returns></returns>
        Stream GetFileStream(SPListItem listItem, ContentReference contentLink);

        Stream GetFileSteam(string relativeUrl);

        string GetFileRelativeUrl(SPListItem listItem);

        string GetListItemJsonData(SPListItem item);
    }
}