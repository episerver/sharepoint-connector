using System;
using EPiServer.Core;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.Models
{
    /// <summary>
    /// Wraps the SharePoint List, View, ListItem into a model
    /// </summary>
    public class ModelWrapper : IModelWrapper
    {
        private ContentReference _contentLink;
        public ContentReference ContentLink
        {
            get
            {
                if (ContentReference.IsNullOrEmpty(_contentLink))
                {
                    _contentLink = ServiceLocator.Current.GetInstance<ISharePointContentMappingRepository>().GetContentLink(this.UniqueKey());
                }

                return _contentLink;
            }
            set
            {
                _contentLink = value;
            }
        }

        public bool IsFile { get; private set; }

        /// <summary>
        /// SharePoint List object
        /// </summary>
        public SPList List { get; set; }

        /// <summary>
        /// SharePoint View object
        /// </summary>
        public SPListView View { get; set; }

        /// <summary>
        /// SharePoint ListItem object (file/folder)
        /// </summary>
        public SPListItem ListItem { get; set; }

        public Uri BlobUri { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ModelWrapper"/> class.
        /// </summary>
        /// <param name="rootId">The given SharePoint List identification</param>
        /// <param name="listItem">The given SharePoint ListItem object.</param>
        /// <param name="view">The given SharePoint View object.</param>
        public ModelWrapper(SPList list, SPListView view, SPListItem listItem, bool isFile = false)
        {
            List = list;
            View = view;
            ListItem = listItem;
            IsFile = isFile;
        }

        /// <summary>
        /// Return a string in the concatenation of the list, view identification and the listItem identification format
        /// </summary>
        /// <returns></returns>
        public string UniqueKey()
        {
            return (ListItem != null) ?
                string.Format("{0}~{1}~{2}", List.Id, View.Id, ListItem.Id) :
                string.Format("{0}~{1}", List.Id, View.Id);
        }

        /// <summary>
        /// Mapping on fly for list item
        /// </summary>
        /// <returns></returns>
        public string ToMappingUrl()
        {
            if (View != null && ListItem != null)
            {
                return string.Format("{0}~{1}", View.Id, IsFile ?
                    ListItem.File.ServerRelativeUrl : ListItem.Folder.ServerRelativeUrl);
            }

            if (View != null && List != null)
            {
                return string.Format("{0}~{1}", View.Id, List.RootFolder.ServerRelativeUrl);
            }

            return string.Empty;
        }

    }
}