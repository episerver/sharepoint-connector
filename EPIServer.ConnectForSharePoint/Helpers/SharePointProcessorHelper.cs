using System.Runtime.Caching;
using System.Collections.Generic;
using System.IO;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.SPServices;
using System;
using EPiServer.ConnectForSharePoint.Models.ViewModels;
using Newtonsoft.Json.Linq;
using System.Linq;
using System.Diagnostics.CodeAnalysis;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Helper class for processing SPListItems and retrieving document data, list metdata etc.
    /// </summary>
    public class SharePointProcessorHelper : ISharePointProcessorHelper
    {
        private IServiceLocator     _serviceLocator;
        private IFileStreamService  _fileStreamService;
        private IListsViewsService  _listViewsService;
        private ObjectCache _metaDataCache = MemoryCache.Default;

        [ExcludeFromCodeCoverage]
        public SharePointProcessorHelper()
        {
            _serviceLocator = ServiceLocator.Current;
            _fileStreamService = _serviceLocator.GetInstance<IFileStreamService>();
            _listViewsService = _serviceLocator.GetInstance<IListsViewsService>();
        }

        internal SharePointProcessorHelper(IServiceLocator serviceLocator)
        {
            _serviceLocator = serviceLocator;
            _fileStreamService = serviceLocator.GetInstance<IFileStreamService>();
            _listViewsService = serviceLocator.GetInstance<IListsViewsService>();
        }

        /// <summary>
        /// Gets the stream data assocatied with a specific SPListItem, 
        /// useful when adding or updating a document in cms
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public Stream GetDocumentStream(SPListItem item)
        {
            var relativeUrl = _fileStreamService.GetFileRelativeUrl(item);
            return _fileStreamService.GetFileSteam(relativeUrl);
        }

        /// <summary>
        /// Gets the Json encoded data associated with the specific list item,
        /// useful when adding or updating a list item in cms
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public string GetRawListItemData(SPListItem item)
        {
            string str = _fileStreamService.GetListItemJsonData(item);

            // Remove stuff that causes the parsing to fail.
            string lineSeparator = ((char)0x2028).ToString();
            string paragraphSeparator = ((char)0x2029).ToString();
            str = str.Replace("\\u200b", string.Empty);
            str = str.Replace(@"\r\n", string.Empty).Replace(@"\r", string.Empty).Replace(@"\n", string.Empty);

            return str;
        }

        /// <summary>
        /// Returns a collection of  ListItemFields that represent the metadata of the list.
        /// Recommend to not set inValidateCache to true since the cache is automatically cleared
        /// out when the user clicks refresh, save, or the scheduled task runs.
        /// </summary>
        /// <param name="list">The list item to get the fields for</param>
        /// <param name="inValidateCache">Clear out the cached dictionary for this item.</param>
        /// <returns></returns>
        public IDictionary<string, SPFieldItem> GetMetaData(SPList list, bool inValidateCache = false)
        {
            if (inValidateCache)
                _metaDataCache.Remove(list.Id.ToString());

            IDictionary<string, SPFieldItem> dictionary = _metaDataCache.Get(list.Id.ToString()) as IDictionary<string, SPFieldItem>;
            if(dictionary == null)
            {
                dictionary = _listViewsService.GetFieldItemDictionary(list.Id);
                CacheItemPolicy p = new CacheItemPolicy(); // infinite by default
                _metaDataCache.Add(list.Id.ToString(), dictionary, p);
            }

            return dictionary;
        }

        /// <summary>
        /// Parses the raw Json data for a list and returns a collection of list item fields that have the 
        /// internal name changed to display name
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="rawdata"></param>
        /// <returns></returns>
        public List<ListItemField> ParseRawListItemData(IDictionary<string, SPFieldItem> dictionary, string rawdata)
        {
            var list = new List<ListItemField>();
            if (string.IsNullOrEmpty(rawdata) || dictionary == null)
                return list;

            JObject obj = JObject.Parse(rawdata);
            if (obj["d"]["results"].HasValues == false)
                return list;

            var data = obj["d"]["results"][0].Value<JObject>();
            List<string> keys = data.Properties().Select(p => p.Name).ToList();
            foreach (var prop in data.Properties())
            {
                string property = prop.Name;
                string propvalue = prop.Value.ToString();
                string url = "";
                string description = "";
                bool addToList = true;

                // get the DisplayName aka title from the property dictionary for this list
                SPFieldItem field;
                if (dictionary.TryGetValue(property, out field))
                {
                    property = field.Title;
                }
                else
                {
                    // well this is wierd, sometimes the value we get from the sharepoint webservice
                    // has data in a different format like OData__"property" - hack
                    //
                    foreach (var entry in dictionary)
                    {
                        if (property.Contains(entry.Key))
                        {
                            property = entry.Value.Title;
                            break;
                        }
                    }
                }

                try
                {
                    if (obj["d"]["results"].HasValues == true)
                    {
                        var data2 = obj["d"]["results"][0][prop.Name].Value<JObject>();
                        if (data2 != null)
                        {
                            foreach (var prop2 in data2.Properties())
                            {
                                string property2 = prop2.Name;
                                string propvalue2 = prop2.Value.ToString();

                                if (string.Equals(property2, "Url", StringComparison.OrdinalIgnoreCase))
                                {
                                    url = propvalue2;
                                }
                                if (string.Equals(property2, "Description", StringComparison.OrdinalIgnoreCase))
                                {
                                    description = propvalue2;
                                    propvalue = description;
                                }
                                if (string.Equals(property2, "__deferred", StringComparison.OrdinalIgnoreCase))
                                { // Dont add these things they are urls to objects in sharepoint for the folder, file etc.
                                    addToList = false;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // Unknown why we have an empty catch here. 
                    // I.e. some document types do not return the expected data in the results field maybe? 
                }

                // More nodes we dont want to add, these actually come back by default with the view
                if (string.Equals(property, "FileSystemObjectType", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "__metadata", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "IDDD", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "ID", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "created", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(property, "modified", StringComparison.OrdinalIgnoreCase))
                {
                    addToList = false;
                }

                if (addToList)
                    list.Add(new ListItemField() { Name = property, Value = propvalue, Url = url, Description = description });
            }
            return list;
        }
    }
}