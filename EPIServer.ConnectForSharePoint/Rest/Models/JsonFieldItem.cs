using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    public class JsonFieldItem
    {
        public D d { get; set; }
        public class D
        {
            public __Metadata __metadata { get; set; }

            public string InternalName { get; set; }
            public string StaticName { get; set; }
            public bool Hidden { get; set; }
            public string Group { get; set; }
            public string Title { get; set; } // display name
            public string DefaultValue { get; set; }
            public bool Filterable { get; set; }
        }
    }
}