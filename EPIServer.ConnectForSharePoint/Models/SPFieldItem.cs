using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class SPFieldItem
    {
        public string InternalName { get; set; }
        public string StaticName { get; set; }
        public bool Hidden { get; set; }
        public string Group { get; set; }
        public string Title { get; set; } // display name
        public string DefaultValue { get; set; }
        public bool Filterable { get; set; }
    }
}