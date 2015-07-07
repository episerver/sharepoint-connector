namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     FolderCollection resource
    ///         Represents a collection of Folder resources.
    /// Endpoint URI
    ///     http://<site url>/_api/web/folders (top-level folders at the root site)
    ///     http://<site url>/_api/web/getfolderbyserverrelativeurl('/<folder name>')/folders
    /// </summary>
    public class JsonFolderCollection
    {
        public D d { get; set; }

        public class D
        {
            public JsonFolder.D[] results { get; set; }
        }
    }
}