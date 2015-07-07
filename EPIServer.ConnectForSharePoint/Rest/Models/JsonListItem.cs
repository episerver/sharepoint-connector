using System;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     ListItem resource
    ///         Represents an item or row in a list. A list item is a type of SP.SecurableObject.
    /// Endpoint URI
    ///     http://<site url>/web/lists(guid'<list id>')/items(<item id>)
    ///     http://<site url>/web/lists/getbytitle('<list title>')/items(<item id>)
    /// </summary>
    public class JsonListItem
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public Firstuniqueancestorsecurableobject FirstUniqueAncestorSecurableObject { get; set; }
            public Roleassignments RoleAssignments { get; set; }
            public Attachmentfiles AttachmentFiles { get; set; }
            public Contenttype ContentType { get; set; }
            public Fieldvaluesashtml FieldValuesAsHtml { get; set; }
            public Fieldvaluesastext FieldValuesAsText { get; set; }
            public Fieldvaluesforedit FieldValuesForEdit { get; set; }
            public JsonFile.D File { get; set; }
            public JsonFolder.D Folder { get; set; }
            public Parentlist ParentList { get; set; }
            public int FileSystemObjectType { get; set; }
            public int Id { get; set; }
            public string ContentTypeId { get; set; }
            public string Title { get; set; }
            public DateTime Modified { get; set; }
            public DateTime Created { get; set; }
            public int AuthorId { get; set; }
            public int EditorId { get; set; }
            public string OData__UIVersionString { get; set; }
            public bool Attachments { get; set; }
            public string GUID { get; set; }
            public Predecessorsid PredecessorsId { get; set; }
            public string Priority { get; set; }
            public string Status { get; set; }
            public int PercentComplete { get; set; }
            public Assignedtoid AssignedToId { get; set; }
            public string Body { get; set; }
            public DateTime StartDate { get; set; }
            public DateTime DueDate { get; set; }
            public string Checkmark { get; set; }
            public string RelatedItems { get; set; }
            public string EncodedAbsUrl { get; set; }

            public bool? VideoSetShowDownloadLink { get; set; }
            public bool? VideoSetShowEmbedLink { get; set; }
        }
    }
}