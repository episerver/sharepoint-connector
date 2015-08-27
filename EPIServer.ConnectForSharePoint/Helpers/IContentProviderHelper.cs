using EPiServer.Core;
using EPiServer.Security;
using System;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface IContentProviderHelper
    {
        ContentReference FindRecursiveFolders(string name, ContentReference startRef);
        void SetAccessRight(ContentReference contentReference, AccessLevel accessLevel);
        ContentReference FindMediaInFolder(string name, ContentReference folderRef);
        ContentReference FindMediaInFolder(Guid SPItemGuid, ContentReference folderRef);
    }
}
