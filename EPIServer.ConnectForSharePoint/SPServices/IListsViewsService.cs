using EPiServer.ConnectForSharePoint.Models;
using System;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IListsViewsService
    {
        /// <summary>
        /// Get all lists have match base type
        /// </summary>
        /// <param name="baseType">The base type value.</param>
        IEnumerable<SPList> GetLists(int baseType);

        /// <summary>
        /// Get list from an Id
        /// </summary>
        /// <param name="listId">The list Id.</param>
        SPList GetList(string listId);

        /// <summary>
        /// Get all items of a list
        /// </summary>
        /// <param name="list">The list to get items from.</param>
        IEnumerable<SPListItem> GetAllItems(SPList list);

        /// <summary>
        /// Get items of a list bases on view query
        /// </summary>
        /// <param name="list">The list to get items from.</param>
        /// <param name="viewQuery">The view query CAML including View tag.</param>
        IEnumerable<SPListItem> GetItems(SPList list, string viewQuery);

        /// <summary>
        /// Get items in a list's folder bases on a view
        /// </summary>
        /// <param name="list">The list to get items from.</param>
        /// <param name="view">The view for filtering items.</param>
        /// <param name="folder">The folder to get items from.</param>
        IEnumerable<SPListItem> GetItems(SPList list, SPListView view, SPFolder folder);

        /// <summary>
        /// Gets a list item from its parent list
        /// </summary>
        /// <param name="list">The list to get item from.</param>
        /// <param name="listItemId">The list item identifier.</param>
        SPListItem GetListItem(SPList list, string listItemId);

        /// <summary>
        /// Gets all folders for the specified list and view
        /// </summary>
        /// <param name="list"></param>
        /// <param name="view"></param>
        /// <returns></returns>
        IEnumerable<SPListItem> GetAllFolders(SPList list, SPListView view);

        /// <summary>
        /// Gets the field schema for a particular list.
        /// </summary>
        /// <param name="list"></param>
        /// <returns>Dictionary key is the internal name</returns>
        Dictionary<string, SPFieldItem> GetFieldItemDictionary(Guid ListID);
    }
}