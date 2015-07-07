namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     SP.User resource
    ///         Represents a user on a SharePoint Web site.
    /// Endpoint URI
    ///     http://<site url>/_api/web/siteusers(@v)?@v='<login name>'
    /// </summary>
    public class JsonUser
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public Groups Groups { get; set; }
            public int Id { get; set; }
            public bool IsHiddenInUI { get; set; }
            public string LoginName { get; set; }
            public string Title { get; set; }
            public int PrincipalType { get; set; }
            public string Email { get; set; }
            public bool IsShareByEmailGuestUser { get; set; }
            public bool IsSiteAdmin { get; set; }
            public Userid UserId { get; set; }
        }
    }
}