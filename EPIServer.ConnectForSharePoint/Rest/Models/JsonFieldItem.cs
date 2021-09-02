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

            public bool CanBeDeleted { get; set; }
            public string DefaultValue { get; set; }
            public string Description { get; set; }
            public string Direction { get; set; }
            public bool EnforceUniqueValues { get; set; }
            public string EntityPropertyName { get; set; }
            public bool Filterable { get; set; }
            public bool FromBaseType { get; set; }
            public string Group { get; set; }
            public bool Hidden { get; set; }
            public string Id { get; set; }
            public string Indexed { get; set; }
            public string InternalName { get; set; }
            public string JSLink { get; set; }
            public bool ReadOnlyField { get; set; }
            public bool Required { get; set; }
            public string SchemaXml { get; set; }
            public string Scope { get; set; }
            public bool Sealed { get; set; }
            public bool Sortable { get; set; }
            public string StaticName { get; set; }
            public string Title { get; set; } // display name
            public int FieldTypeKind { get; set; }
            public string TypeAsString { get; set; }
            public string TypeDisplayName { get; set; }
            public string TypeShortDescription { get; set; }
            public string ValidationFormula { get; set; }
            public string ValidationMessage { get; set; }
            public bool FillInChoice { get; set; }
            public string Mappings { get; set; }
            public Object Choices { get; set; }
            public string EditFormat { get; set; }
        }
    }
}

