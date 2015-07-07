using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class ContextWebInformation
    {
        public int FormDigestTimeoutSeconds { get; set; }
        public string FormDigestValue { get; set; }
        public string LibraryVersion { get; set; }
        public string SiteFullUrl { get; set; }
        public string WebFullUrl { get; set; }
        public IEnumerable<string> SupportedSchemaVersions { get; set; }
    }
}