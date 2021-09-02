namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    public class Groups
    {
        public __Deferred __deferred { get; set; }
    }

    public class Userid
    {
        public __Metadata __metadata { get; set; }
        public string NameId { get; set; }
        public string NameIdIssuer { get; set; }
    }

    public class __Deferred
    {
        public string uri { get; set; }
    }

    public class __Metadata
    {
        public string id { get; set; }
        public string uri { get; set; }
        public string etag { get; set; }
        public string type { get; set; }
    }

    public class Informationrightsmanagementsettings
    {
        public __Deferred __deferred { get; set; }
    }

    public class Firstuniqueancestorsecurableobject
    {
        public __Deferred __deferred { get; set; }
    }

    public class Roleassignments
    {
        public __Deferred __deferred { get; set; }
    }

    public class Attachmentfiles
    {
        public __Deferred __deferred { get; set; }
    }

    public class Contenttypes
    {
        public __Deferred __deferred { get; set; }
    }

    public class Defaultview
    {
        public __Deferred __deferred { get; set; }
    }

    public class Eventreceivers
    {
        public __Deferred __deferred { get; set; }
    }

    public class Fields
    {
        public __Deferred __deferred { get; set; }
    }

    public class Forms
    {
        public __Deferred __deferred { get; set; }
    }

    public class Viewfields
    {
        public __Deferred __deferred { get; set; }
    }

    public class Contenttype
    {
        public __Deferred __deferred { get; set; }
    }

    public class Contenttypeid
    {
        public __Metadata __metadata { get; set; }
        public string StringValue { get; set; }
    }

    public class Fieldvaluesashtml
    {
        public __Deferred __deferred { get; set; }
    }

    public class Fieldvaluesastext
    {
        public __Deferred __deferred { get; set; }
    }

    public class Fieldvaluesforedit
    {
        public __Deferred __deferred { get; set; }
    }

    public class File
    {
        public __Deferred __deferred { get; set; }
    }

    public class Files
    {
        public __Deferred __deferred { get; set; }
    }

    public class Folder
    {
        public __Deferred __deferred { get; set; }
    }

    public class Folders
    {
        public __Deferred __deferred { get; set; }
    }

    public class Parentfolder
    {
        public __Deferred __deferred { get; set; }
    }

    public class Parentlist
    {
        public __Deferred __deferred { get; set; }
    }

    public class Predecessorsid
    {
        public int?[] results { get; set; }
    }

    public class Assignedtoid
    {
        public int[] results { get; set; }
    }

    public class Properties
    {
        public __Deferred __deferred { get; set; }
    }

    /// <summary>
    /// SharePoint - REST API
    ///     Listitemallfields resource
    ///         Represents an item correspnding to a file or folder.
    /// Endpoint URI
    ///     http://<site url>/_api/web/getfilebyserverrelativeurl('serverrelativeurl')/listitemallfields
    ///     http://<site url>/_api/web/getfolderbyserverrelativeurl('serverrelativeurl')/listitemallfields
    /// </summary>
    public class Listitemallfields : JsonListItem.D
    {
        public __Deferred __deferred { get; set; }
    }

    public class Author
    {
        public __Deferred __deferred { get; set; }
    }

    public class Checkedoutbyuser
    {
        public __Deferred __deferred { get; set; }
    }

    public class Lockedbyuser
    {
        public __Deferred __deferred { get; set; }
    }

    public class Modifiedby
    {
        public __Deferred __deferred { get; set; }
    }

    public class Versions
    {
        public __Deferred __deferred { get; set; }
    }
}