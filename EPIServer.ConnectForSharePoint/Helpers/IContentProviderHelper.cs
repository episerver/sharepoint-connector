using EPiServer.Core;
using EPiServer.Security;
using System;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface IContentProviderHelper
    {
        ContentReference FindRecursiveFolders(string name, ContentReference startRef);
        ContentReference FindMediaInFolder(Guid SPItemGuid, ContentReference folderRef);
    }
}
