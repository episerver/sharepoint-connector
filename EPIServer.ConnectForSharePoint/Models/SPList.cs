using System;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class SPList
    {
        /// <summary>
        /// Gets or sets the id.
        /// </summary>
        /// <value>The id.</value>
        public Guid Id { get; set; }

        /// <summary>
        /// Gets or sets the title.
        /// </summary>
        /// <value>The title.</value>
        public string Title { get; set; }

        /// <summary>
        /// Gets the list's base template value.
        /// </summary>
        /// <value>The value of the list base template.</value>
        public int BaseTemplate { get; set; }

        public int BaseType { get; set; }

        /// <summary>
        /// Get the URL points to the list' views, e.g. /_api/Web/Lists(guid'8b75786d-bf8a-4b66-95a9-6a04f197879a')/Views
        /// </summary>
        public string ViewsUrl
        {
            get
            {
                return string.Format(Consts.REST_LIST_VIEWS_PATH_FORMAT, Id);
            }
        }

        public IEnumerable<SPListView> Views { get; set; }

        public DateTime Created { get; set; }

        public DateTime LastItemModifiedDate { get; set; }

        public DateTime LastItemDeletedDate { get; set; }

        public bool Hidden { get; set; }

        public int ItemCount { get; set; }

        public string Description { get; set; }

        public SPFolder RootFolder { get; set; }
    }
}