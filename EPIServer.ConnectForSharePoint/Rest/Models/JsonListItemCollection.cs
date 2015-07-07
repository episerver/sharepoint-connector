namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     ListItemCollection resource
    ///         Represents a collection of ListItem resources.
    /// Endpoint URI
    ///     http://<site url>/web/lists(guid'<list id>')/items
    ///     http://<site url>/web/lists/getbytitle('<list title>')/items
    /// </summary>
    public class JsonListItemCollection
    {
        public D d { get; set; }

        public class D
        {
            public JsonListItem.D[] results { get; set; }
        }
    }
}