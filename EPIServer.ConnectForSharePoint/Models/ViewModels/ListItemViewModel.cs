using EPiServer.ConnectForSharePoint.Messaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace EPiServer.ConnectForSharePoint.Models.ViewModels
{
    public class ListItemViewModel
    {
        public string Name { get; set; }
        public string XmlData { get; set; }

        public List<ListItemField> Fields
        {
            get
            {
                var pro = new SharePointProcessor();
                return pro.getFields(XmlData);
            }
        }
    }
}