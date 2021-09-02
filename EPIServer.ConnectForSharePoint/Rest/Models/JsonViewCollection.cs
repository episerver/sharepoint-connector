namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     ViewCollection resource
    ///         Represents a collection of View resources.
    /// Endpoint URI
    ///     http://<site url>/_api/web/lists(guid'<list id>')/views
    /// </summary>
    public class JsonViewCollection
    {
        public D d { get; set; }

        public class D
        {
            public JsonView.D[] results { get; set; }
        }
    }
}