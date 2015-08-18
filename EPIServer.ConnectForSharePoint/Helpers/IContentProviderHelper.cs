using EPiServer.Core;
using EPiServer.Security;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface IContentProviderHelper
    {
        ContentReference FindRecursiveFolders(string name, ContentReference startRef);
        void SetAccessRight(ContentReference contentReference, AccessLevel accessLevel);
        ContentReference FindMediaInFolder(string name, ContentReference folderRef);
    }
}
