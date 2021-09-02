using System;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     File resource
    ///         Represents a file in a SharePoint Web site that can be a Web Part Page, an item in a document library, or a file in a folder.
    /// Endpoint URI
    ///     http://<site url>/_api/web/getfilebyserverrelativeurl('/<folder name>/<file name>')
    /// </summary>
    public class JsonFile
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public JsonUser.D Author { get; set; }
            public Checkedoutbyuser CheckedOutByUser { get; set; }
            public Listitemallfields ListItemAllFields { get; set; }
            public Lockedbyuser LockedByUser { get; set; }
            public JsonUser.D ModifiedBy { get; set; }
            public Versions Versions { get; set; }
            public string CheckInComment { get; set; }
            public int CheckOutType { get; set; }
            public string ContentTag { get; set; }
            public int CustomizedPageStatus { get; set; }
            public string ETag { get; set; }
            public bool Exists { get; set; }
            public string Length { get; set; }
            public int Level { get; set; }
            public int MajorVersion { get; set; }
            public int MinorVersion { get; set; }
            public string Name { get; set; }
            public string ServerRelativeUrl { get; set; }
            public DateTime TimeCreated { get; set; }
            public DateTime TimeLastModified { get; set; }
            public string Title { get; set; }
            public int UIVersion { get; set; }
            public string UIVersionLabel { get; set; }
            public string UniqueId { get; set; }
        }
    }
}