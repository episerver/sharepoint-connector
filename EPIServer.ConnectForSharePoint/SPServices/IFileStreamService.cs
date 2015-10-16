using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using System;
using System.IO;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IFileStreamService
    {
        Stream GetFileSteam(string relativeUrl);

        string GetFileRelativeUrl(SPListItem listItem);

        string GetListItemJsonData(SPListItem item);
    }
}