using System;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     ListCollection resource
    ///         Represents a collection of lists.
    /// Endpoint URI
    ///     http://<site url>/web/lists
    /// </summary>
    public class JsonListCollection
    {
        public D d { get; set; }

        public class D
        {
            public JsonList.D[] results { get; set; }
        }
    }
}

