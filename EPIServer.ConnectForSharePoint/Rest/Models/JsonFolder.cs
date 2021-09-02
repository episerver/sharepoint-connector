using System;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     Folder resource
    ///         Represents a folder on a SharePoint Web site.
    /// Endpoint URI
    ///     http://<site url>/_api/web/getfolderbyserverrelativeurl('/<folder name>')
    /// </summary>
    public class JsonFolder
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public Files Files { get; set; }
            public Listitemallfields ListItemAllFields { get; set; }
            public Parentfolder ParentFolder { get; set; }
            public Properties Properties { get; set; }
            public Folders Folders { get; set; }
            public int ItemCount { get; set; }
            public string Name { get; set; }
            public string ServerRelativeUrl { get; set; }
            public DateTime TimeCreated { get; set; }
            public DateTime TimeLastModified { get; set; }
            public string WelcomePage { get; set; }
            public string UniqueId { get; set; }
        }
    }
}