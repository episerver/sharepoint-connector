using EPiServer.ConnectForSharePoint.Helpers;
using System;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Models.ViewModels
{
    public class ListItemViewModel
    {
        public string Name { get; set; }
        public string JsonData { get; set; }
        public Guid ListID { get; set; }

        public List<ListItemField> Fields
        {
            get
            {
                var helper = new SharePointProcessorHelper();
                SPList list = new SPList() { Id = ListID };
                return helper.ParseRawListItemData(helper.GetMetaData(list), JsonData);
            }
        }
    }
}