namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     View resource
    ///         Specifies a list view.
    /// Endpoint URI
    ///     http://<site url>/_api/web/lists(guid'<list id>')/views('<view id>')
    /// </summary>
    public class JsonView
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public Viewfields ViewFields { get; set; }
            public string Aggregations { get; set; }
            public string AggregationsStatus { get; set; }
            public string BaseViewId { get; set; }
            public Contenttypeid ContentTypeId { get; set; }
            public bool DefaultView { get; set; }
            public bool DefaultViewForContentType { get; set; }
            public bool EditorModified { get; set; }
            public string Formats { get; set; }
            public bool Hidden { get; set; }
            public string HtmlSchemaXml { get; set; }
            public string Id { get; set; }
            public string ImageUrl { get; set; }
            public bool IncludeRootFolder { get; set; }
            public string ViewJoins { get; set; }
            public string JSLink { get; set; }
            public string ListViewXml { get; set; }
            public string Method { get; set; }
            public bool MobileDefaultView { get; set; }
            public bool MobileView { get; set; }
            public string ModerationType { get; set; }
            public bool OrderedView { get; set; }
            public bool Paged { get; set; }
            public bool PersonalView { get; set; }
            public string ViewProjectedFields { get; set; }
            public string ViewQuery { get; set; }
            public bool ReadOnlyView { get; set; }
            public bool RequiresClientIntegration { get; set; }
            public int RowLimit { get; set; }
            public int Scope { get; set; }
            public string ServerRelativeUrl { get; set; }
            public string StyleId { get; set; }
            public bool Threaded { get; set; }
            public string Title { get; set; }
            public string Toolbar { get; set; }
            public string ToolbarTemplateName { get; set; }
            public string ViewType { get; set; }
            public string ViewData { get; set; }
        }
    }
}