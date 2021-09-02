using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// The object for storing data from /_api/contextinfo endpoint
    /// </summary>
    public class GetContextWebInformation
    {
        public __Metadata __metadata { get; set; }
        public int FormDigestTimeoutSeconds { get; set; }
        public string FormDigestValue { get; set; }
        public string LibraryVersion { get; set; }
        public string SiteFullUrl { get; set; }
        public SupportedSchemaVersions SupportedSchemaVersions { get; set; }
        public string WebFullUrl { get; set; }
    }

    public class SupportedSchemaVersions
    {
        public __Metadata __metadata { get; set; }
        public List<string> results { get; set; }
    }

    public class ContextWebInformationRestResponse
    {
        public D d { get; set; }
    }

    public class D
    {
        public GetContextWebInformation GetContextWebInformation { get; set; }
    }
}