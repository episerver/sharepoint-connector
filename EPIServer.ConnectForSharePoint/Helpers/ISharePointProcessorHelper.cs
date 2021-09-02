using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using System.Collections.Generic;
using System.IO;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Helper class for processing SPListItems and retrieving document data, list metdata etc.
    /// </summary>
    public interface ISharePointProcessorHelper
    {
        /// <summary>
        /// Gets the stream data assocatied with a specific SPListItem, 
        /// useful when adding or updating a document in cms
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        Stream GetDocumentStream(SPListItem item);

        /// <summary>
        /// Gets the Json encoded data associated with the specific list item,
        /// useful when adding or updating a list item in cms
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        string GetRawListItemData(SPListItem item);

        /// <summary>
        /// Returns a collection of  ListItemFields that represent the metadata of the list.
        /// Recommend to not set inValidateCache to true since the cache is automatically cleared
        /// out when the user clicks refresh, save, or the scheduled task runs.
        /// </summary>
        /// <param name="list">The list item to get the fields for</param>
        /// <param name="inValidateCache">Clear out the cached dictionary for this item.</param>
        /// <returns></returns>
        /// 
        IDictionary<string, SPFieldItem> GetMetaData(SPList list, bool inValidateCache = false);

        /// <summary>
        /// Parses the raw Json data for a list and returns a collection of list item fields that have the 
        /// internal name changed to display name
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="rawdata"></param>
        /// <returns></returns>
        List<ListItemField> ParseRawListItemData(IDictionary<string, SPFieldItem> dictionary, string rawdata);
    }
}
