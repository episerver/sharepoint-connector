using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    interface ISharepointToCmsUtilities
    {
        /// <summary>
        /// Find all instances sharepoint items in the cms database.
        /// </summary>
        /// <returns></returns>
        IEnumerable<IContent> FindSharePointItems();

        /// <summary>
        /// Makes a call to sharepoint and gets all of the lists that are supported by the connector.
        /// </summary>
        /// <returns></returns>
        IEnumerable<SPList> FindSupportedLists();

        /// <summary>
        /// Removes all folders under the specified folder that are not tagged as Sharepoint folders.
        /// (Users are not allowed to add folders within the folder we sync content data too)
        /// </summary>
        /// <param name="folderContentReference"></param>
        void RemoveUserAddedFolders(IMessageEmitterFactory emmiter, ContentReference folderContentReference);

        /// <summary>
        /// If any subfolders of the specified folderRef exist, that are not in the sharepoint list
        /// we will fire the delete message.
        /// </summary>
        /// <param name="emmiter"></param>
        /// <param name="folderRef"></param>
        /// <param name="sharepointListItems"></param>
        void RemoveSharePointAddedFolders(IMessageEmitterFactory emmiter, ContentReference folderRef, IEnumerable<SPListItem> sharepointListItems);

        /// <summary>
        /// Adds or updates SPListItems that are not folders 
        /// </summary>
        /// <param name=""></param>
        void AddOrUpdateItem(IMessageEmitterFactory emmiter, SPListItem spItem, IEnumerable<IContent> cmsItems);

        /// <summary>
        /// Emmits the add or update message for the specified item
        /// </summary>
        /// <param name="emmiter"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        ContentReference AddOrUpdateFolder(IMessageEmitterFactory emmiter, SPListItem item);

        /// <summary>
        /// Emmits a delete message for every CMS Item not in the sharepointListItems 
        /// </summary>
        /// <param name="cmsListItems"></param>
        /// <param name="sharepointListItems"></param>
        void DeleteDocumentsAndItems(IMessageEmitterFactory emmiter, IEnumerable<IContent> cmsListItems, IEnumerable<SPListItem> sharepointListItems);

        /// <summary>
        /// Converts the specified list of SPLists into a collection of keyvalue pair objects where the 
        /// key is the list id and the view is the value
        /// </summary>
        /// <param name="lists"></param>
        /// <returns></returns>
        IEnumerable<KeyValuePair<string, string>> GetListViewPairs(IEnumerable<SPList> lists);

        /// <summary>
        /// Retrieves the keyvalue pair values from the configuration that contains the configured lists view combinations.
        /// The list passed in is typically the ALL lists.
        /// </summary>
        /// <param name="lists"></param>
        /// <returns></returns>
        IEnumerable<KeyValuePair<string, string>> GetSelectedListViewIds(IEnumerable<SPList> lists);

        /// <summary>
        /// Deletes all folders in the specified splists collection, UNLESS there list ids to exclude from deletion.
        /// </summary>
        /// <param name="splists"></param>
        /// <param name="listids"></param>
        void QueueRootFoldersForDelete(IMessageEmitterFactory emmiter, IEnumerable<SPList> splists, IEnumerable<string> listids = null);

        /// <summary>
        /// REturns a SPListItem that represents a parent folder.
        /// </summary>
        /// <param name="list"></param>
        /// <param name="view"></param>
        /// <returns></returns>
        SPListItem CreateParentSPListItem(SPList list, SPListView view);
    }
}
