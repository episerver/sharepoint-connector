using System;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     List resource
    ///         Represents a SharePoint list. A list is a type of SP.SecurableObject.
    /// Endpoint URI
    ///     http://<site url>/web/lists(guid'<list id>')
    /// </summary>
    public class JsonList
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public Firstuniqueancestorsecurableobject FirstUniqueAncestorSecurableObject { get; set; }
            public Roleassignments RoleAssignments { get; set; }
            public Contenttypes ContentTypes { get; set; }
            public Defaultview DefaultView { get; set; }
            public Eventreceivers EventReceivers { get; set; }
            public Fields Fields { get; set; }
            public Forms Forms { get; set; }
            public Informationrightsmanagementsettings InformationRightsManagementSettings { get; set; }
            public JsonFolder.D RootFolder { get; set; }
            public JsonViewCollection.D Views { get; set; }
            public bool AllowContentTypes { get; set; }
            public int BaseTemplate { get; set; }
            public int BaseType { get; set; }
            public bool ContentTypesEnabled { get; set; }
            public DateTime Created { get; set; }
            public string DefaultContentApprovalWorkflowId { get; set; }
            public string Description { get; set; }
            public string Direction { get; set; }
            public string DocumentTemplateUrl { get; set; }
            public int DraftVersionVisibility { get; set; }
            public bool EnableAttachments { get; set; }
            public bool EnableFolderCreation { get; set; }
            public bool EnableMinorVersions { get; set; }
            public bool EnableModeration { get; set; }
            public bool EnableVersioning { get; set; }
            public string EntityTypeName { get; set; }
            public bool ForceCheckout { get; set; }
            public bool HasExternalDataSource { get; set; }
            public bool Hidden { get; set; }
            public string Id { get; set; }
            public string ImageUrl { get; set; }
            public bool IrmEnabled { get; set; }
            public bool IrmExpire { get; set; }
            public bool IrmReject { get; set; }
            public bool IsApplicationList { get; set; }
            public bool IsCatalog { get; set; }
            public bool IsPrivate { get; set; }
            public int ItemCount { get; set; }
            public DateTime LastItemDeletedDate { get; set; }
            public DateTime LastItemModifiedDate { get; set; }
            public string ListItemEntityTypeFullName { get; set; }
            public bool MultipleDataList { get; set; }
            public bool NoCrawl { get; set; }
            public string ParentWebUrl { get; set; }
            public bool ServerTemplateCanCreateFolders { get; set; }
            public string TemplateFeatureId { get; set; }
            public string Title { get; set; }
        }
    }
}