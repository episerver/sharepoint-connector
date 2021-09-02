using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    /// <summary>
    /// Interface that defines the share point processor module.
    /// </summary>
    public interface ISharePointProcessor
    {
        /// <summary>
        /// Called by the framework when an document has been retrieved from sharepoint and needs to be proccessed.
        /// </summary>
        /// <param name="item"></param>
        void AddDocument(SPListItem item);
        /// <summary>
        /// Called by the framework when a folder object should be processed.
        /// </summary>
        /// <param name="item"></param>
        void AddFolder(SPListItem item);
        /// <summary>
        /// Called by the framework when an list item has been retrieved from sharepoint and needs to be proccessed.
        /// </summary>
        /// <param name="item"></param>
        void AddListItem(SPListItem item);
        /// <summary>
        /// Called by the framework when a document has been updated on the sharepoint side.
        /// </summary>
        /// <param name="item"></param>
        void UpdateDocument(SPListItem item);
        /// <summary>
        /// Called by the framework when a folder hs been updated on the sharepoint side.
        /// </summary>
        /// <param name="item"></param>
        void UpdateFolder(SPListItem item);
        /// <summary>
        /// Called by the framework when a list item has been updated on the sharepoint side.
        /// </summary>
        /// <param name="item"></param>
        void UpdateListItem(SPListItem item);
        /// <summary>
        /// Called by the framework when a document should be deleted.
        /// </summary>
        /// <param name="item"></param>
        void DeleteDocument(SPListItem item);
        /// <summary>
        /// Called by the framework when a folder should be deleted.
        /// </summary>
        /// <param name="item"></param>
        void DeleteFolder(SPListItem item);
        /// <summary>
        /// Called by the framework when a list item should be deleted.
        /// </summary>
        /// <param name="item"></param>
        void DeleteListItem(SPListItem item);
    }
}
