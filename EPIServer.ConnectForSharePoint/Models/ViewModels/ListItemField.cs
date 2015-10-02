using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models.ViewModels
{
    /// <summary>
    /// Represents a single sharepoint listitem field.
    /// </summary>
    public class ListItemField
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Type { get; set; }
        public string Url { get; set; }
        public string Description { get; set; }

        public override string ToString()
        {
            return "[" + Name + "," + Value + "," + Type + "," + Description + "," + Url + "]";
        }
    }
}