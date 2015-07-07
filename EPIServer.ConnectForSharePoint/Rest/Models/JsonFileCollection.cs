namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     FileCollection resource
    ///         Represents a collection of File resources.
    /// Endpoint URI
    ///     http://<site url>/_api/web/getfolderbyserverrelativeurl('/<folder name>')/files
    /// </summary>
    public class JsonFileCollection
    {
        public D d { get; set; }

        public class D
        {
            public JsonFile.D[] results { get; set; }
        }
    }
}