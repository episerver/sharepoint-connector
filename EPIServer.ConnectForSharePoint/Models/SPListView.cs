using System;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class SPListView
    {
        /// <summary>
        /// Gets or sets the view Id.
        /// </summary>
        /// <value>The view id.</value>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets the view's title.
        /// </summary>
        /// <value>The title.</value>
        public string Title { get; set; }

        public string BaseViewId { get; set; }

        public string ViewType { get; set; }

        public string ListViewXml { get; set; }

        public string ViewQuery { get; set; }

        public bool Hidden { get; set; }

        public bool IncludeRootFolder { get; set; }

        public int Scope { get; set; }
    }

    public class Views
    {
        public List<SPListView> results { get; set; }
    }
}