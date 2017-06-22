using EPiServer.ConnectForSharePoint.Mappers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EPiServer.ConnectForSharePoint.SPServices.REST
{
    [ServiceConfiguration(ServiceType = typeof(IListsViewsService))]
    public class ListsViewsService : RESTServiceBase, IListsViewsService
    {
        private IMapper<JsonListItem.D, SPListItem> _itemMapper;
        private IMapper<JsonList.D, SPList> _listMapper;
        private IMapper<JsonView.D, SPListView> _viewMapper;
        private IMapper<JsonFolder.D, SPFolder> _folderMapper;
        private IMapper<JsonFile.D, SPFile> _fileMapper;
        private IMapper<JsonFieldItem.D, SPFieldItem> _fieldMapper;

        public ListsViewsService() : this(new SPListMapper(), new SPViewMapper(), new SPFolderMapper(), 
            new SPFileMapper(), new SPListItemMapper(), new SPFieldItemMapper() ){ }

        public ListsViewsService(
            IMapper<JsonList.D, SPList> listMapper,
            IMapper<JsonView.D, SPListView> viewMapper,            
            IMapper<JsonFolder.D, SPFolder> folderMapper,
            IMapper<JsonFile.D, SPFile> fileMapper,
            IMapper<JsonListItem.D, SPListItem> itemMapper,
            IMapper<JsonFieldItem.D, SPFieldItem> fieldMapper )
        {
            _listMapper = listMapper;
            _viewMapper = viewMapper;
            _folderMapper = folderMapper;
            _fileMapper = fileMapper;
            _itemMapper = itemMapper;
            _fieldMapper = fieldMapper;
        }

        public IEnumerable<SPList> GetLists(int baseType)
        {
            var restRequest = new RestJsonRequest(Consts.REST_LISTS_PATH, "GET", GetListExpandQuery(), string.Format("filter=BaseType eq {0} and Hidden eq false", baseType));

            return BuildLists(SendRequestWithRetry(restRequest, null, null));
        }

        public SPList GetList(string listId)
        {
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_PATH_FORMAT, listId), "GET", GetListExpandQuery());
            var jsonResult = Deserialize<JsonList>(SendRequest(restRequest, null, null));

            return jsonResult != null ? BuildList(jsonResult.d) : null;
        }

        public IEnumerable<SPListItem> GetAllItems(SPList list)
        {
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_ITEMS_PATH_FORMAT, list.Id), "GET", GetItemExpandQuery(false));

            return BuildListItems(SendRequest(restRequest, null, null), list);
        }

        public IEnumerable<SPListItem> GetPagedItems(SPList list, SPListView view, SPFolder folder, out string nextUri)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(view.ListViewXml);
            var tags = xmlDoc.GetElementsByTagName("RowLimit");
            int rowcount = 256;
            if (tags != null && tags.Count > 0 )
            {
                Int32.TryParse(tags[0].InnerText, out rowcount); 
            }

            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_ITEMS_PATH_FORMAT, list.Id),
                    "GET", GetItemExpandQuery(false, true, rowcount));

            return BuildPagedListItems(SendRequestWithRetry(restRequest, null, null, true), list, out nextUri);
        }

        public IEnumerable<SPListItem> GetNextPagedItems(SPList list, string Uri, out string nextUri)
        {
            var restApiPath = Uri.Substring(Uri.IndexOf("/_api"));
            var restRequest = new RestJsonRequest(restApiPath);

            return BuildPagedListItems(SendRequestWithRetry(restRequest, null, null, true), list, out nextUri);
        }

        public IEnumerable<SPListItem> GetItems(SPList list, string viewQuery)
        {
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_GETITEMS_PATH_FORMAT, list.Id), "POST", GetItemExpandQuery(true));
            var bodyBuilder = new StringBuilder("{ 'query' : {'__metadata': { 'type': 'SP.CamlQuery' }, ");
            bodyBuilder.AppendFormat("'ViewXml': '{0}'", EncodeDataForCAMLQuery(viewQuery));
            bodyBuilder.Append(" } }");

            return BuildListItems(SendRequest(restRequest, bodyBuilder.ToString(), null, true), list);
        }

        public IEnumerable<SPListItem> GetAllFolders(SPList list, SPListView view)
        {
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_GETITEMS_PATH_FORMAT, list.Id), "POST", GetItemExpandQuery(true));
            var bodyBuilder = new StringBuilder("{ 'query' : {'__metadata': { 'type': 'SP.CamlQuery' }, ");

            bodyBuilder.AppendFormat("'ViewXml': '{0}'", EncodeDataForCAMLQuery(GetListViewXmlWithoutViewFields(view)));
            bodyBuilder.AppendFormat(",'FolderServerRelativeUrl': '{0}'", EncodeDataForCAMLQuery(list.RootFolder.ServerRelativeUrl));

            bodyBuilder.Append(",'DatesInUtc': 'true'");
            bodyBuilder.Append(" } }");

            string withQueryString = bodyBuilder.ToString().Replace("<Query>", "<Query><Where><BeginsWith>" +
                        "<FieldRef Name=\"ContentTypeId\"/><Value Type =\"ContentTypeId\">" + "0x0120" + "</Value>" +
                        "</BeginsWith></Where>");
            withQueryString = withQueryString.Replace("<View ", "<View Scope=\"RecursiveAll\" ");           // search sub folders too.

           return BuildListItems(SendRequest(restRequest, withQueryString, null, true), list);
         }

        public IEnumerable<SPListItem> GetItems(SPList list, SPListView view, SPFolder folder)
        {
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_GETITEMS_PATH_FORMAT, list.Id), "POST", GetItemExpandQuery(true));
            var bodyBuilder = new StringBuilder("{ 'query' : {'__metadata': { 'type': 'SP.CamlQuery' }, ");
            bodyBuilder.AppendFormat("'ViewXml': '{0}'", EncodeDataForCAMLQuery(GetListViewXmlWithoutViewFields(view)));
            bodyBuilder.AppendFormat(",'FolderServerRelativeUrl': '{0}'", EncodeDataForCAMLQuery(folder.ServerRelativeUrl));

            // NOTE: this is on our addon Team virtual machine, haven't known other premise SharePoint environments.
            // NOTE: only with POST has this way defining DatesInUtc in the request 
            // NOTE: without DatesInUtc==true the premise SharePoint (addon Team virtual machine) returns local DateTime string 
            // (a DateTime value string without 'Z' at end and applied Regional Settings of SiteCollection already)
            // because without Z the desirialize one more time coverts to server local time. Force DatesInUtc:
            bodyBuilder.Append(",'DatesInUtc': 'true'");
            bodyBuilder.Append(" } }");

            return BuildListItems(SendRequest(restRequest, bodyBuilder.ToString(), null, true), list);
        }

        public SPListItem GetListItem(SPList list, string listItemId)
        {
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LISTITEM_PATH_FORMAT, list.Id, listItemId), "GET", GetItemExpandQuery(false));
            var jsonResult = Deserialize<JsonListItem>(SendRequestWithRetry(restRequest, null, null));

            return jsonResult != null ? BuildListItem(jsonResult.d, list) : null;
        }

        /// <summary>
        /// Return query including File/Folder data for list items
        /// </summary>
        private static string[] GetItemExpandQuery(bool getVideoInfo, bool paged = false, int count = 256)
        {
            // NOTE: could not get EncodedAbsUrl without select clause, but including select clause 
            // we MUST mention explicitly all the fields we need
            // NOTE: don't include the File's Author, ModifiedBy e.g. File/ModifiedBy
            // because of duplicating user data returned - an user modifies multi files 
            var selectClause = new StringBuilder();
            // default fields returned from list/items endpoint
            selectClause.Append("select=Id,GUID,Created,Modified,FileSystemObjectType,Title,EncodedAbsUrl");
            
            // some folder generated automatically when upload a video file to SharePoint,
            // it will contains some properties about the video
            if(getVideoInfo)
            {
                selectClause.Append(",VideoSetShowDownloadLink,VideoSetShowEmbedLink");
            }
            
            // fields expand to File/Folder fields
            // NOTE: some fields e.g File/UniqueId, Folder/UniqueId are not available on default configuration of a premise SharePoint server
            selectClause.Append(",File/Name,File/Length,File/Level,File/ServerRelativeUrl,File/TimeCreated,File/TimeLastModified");
            selectClause.Append(",Folder/Name,Folder/ItemCount,Folder/ServerRelativeUrl,Folder/TimeCreated,Folder/TimeLastModified");

            // NOTE: our addon Team virtual machine returns no data for TimeCreated, TimeLastModified, haven't known other premise SharePoint environments.
            // NOTE: want using Folder/Properties to get more data from premise SharePoint environment (SPFolder TimeCreated, TimeLastModified),
            // but also have seen that "Folder/Properties" in expand clause ONLY works without select clause 
            // then does NOT include Folder/Properties here
            var expandClause = "expand=File/Name,Folder/Name";

            if (paged)
            {
                return new string[] { selectClause.ToString(), expandClause, "skiptoken=Paged=TRUE&$top=" + count };
            }
            else
            {
                return new string[] { selectClause.ToString(), expandClause };
            }
        }

        /// <summary>
        /// Return query including RootFodler/View data for lists
        /// </summary>
        private static string GetListExpandQuery()
        {
            // NOTE: the select/expand for including RootFolder and View data           
            return "expand=Views/Id,RootFolder/Title";
        }

        /// <summary>
        /// Return the ListViewXml without ist ViewFields so that the request's select clause has effect
        /// </summary>
        private string GetListViewXmlWithoutViewFields(SPListView view)
        {
            var viewXml = view.ListViewXml;
            var startIndex = viewXml.IndexOf("<ViewFields>", StringComparison.InvariantCultureIgnoreCase);
            var endIndex = viewXml.IndexOf("</ViewFields>", StringComparison.InvariantCultureIgnoreCase);
            var length = "</ViewFields>".Length;

            if (startIndex >= 0 && endIndex >= 0)
            {
                viewXml = viewXml.Remove(startIndex, endIndex - startIndex + length);
            }

            return viewXml;
        }

        #region buiding SPList, SPListView, SPListItem methods

        public Dictionary<string, SPFieldItem> GetFieldItemDictionary(Guid ListID)
        {
            Dictionary<string, SPFieldItem> dictionary = new Dictionary<string, SPFieldItem>();
            if (ListID.Equals(Guid.Empty))
                return dictionary;

            string requeststring = "/_api/web/lists(guid'{0}')/fields";
            var restRequest = new RestJsonRequest(string.Format(requeststring, ListID), "GET", null);
            var fields = BuildListFieldItems(SendRequestWithRetry(restRequest, null, null, true));

            dictionary = new Dictionary<string, SPFieldItem>();
            foreach (var field in fields)
            {
                dictionary.Add(field.InternalName, field);
            }
                
            return dictionary;
        }

        private IEnumerable<SPFieldItem> BuildListFieldItems(string responseString)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return Enumerable.Empty<SPFieldItem>();
            }

            var jsonResult = Deserialize<JsonFieldItemCollection>(responseString);

            return BuildListFieldItems(jsonResult);
        }

        private IEnumerable<SPFieldItem> BuildListFieldItems(JsonFieldItemCollection jsonResult)
        {
            if (jsonResult == null)
            {
                yield return default(SPFieldItem);
            }

            foreach (var jsonItem in jsonResult.d.results)
            {
                yield return BuildListFieldItem(jsonItem);
            }
        }

        private SPFieldItem BuildListFieldItem(JsonFieldItem.D jsonResult)
        {
            var fieldItem = _fieldMapper.Map(jsonResult);

            if (fieldItem != null)
            {
            }

            return fieldItem;
        }
        private IEnumerable<SPListItem> BuildListItems(string responseString, SPList parentList)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return Enumerable.Empty<SPListItem>();
            }

            var jsonResult = Deserialize<JsonListItemCollection>(responseString);

            return BuildListItems(jsonResult, parentList);
        }

        private IEnumerable<SPListItem> BuildPagedListItems(string responseString, SPList parentList, out string nextUri)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                nextUri = null;
                return Enumerable.Empty<SPListItem>();
            }

            var jsonResult = Deserialize<JsonListItemCollection>(responseString);
            nextUri = jsonResult.d.__next;

            return BuildListItems(jsonResult, parentList);
        }

        private IEnumerable<SPListItem> BuildListItems(JsonListItemCollection jsonResult, SPList parentList)
        {
            if (jsonResult == null)
            {
                yield return default(SPListItem);
            }

            foreach (var jsonItem in jsonResult.d.results)
            {
                yield return BuildListItem(jsonItem, parentList);
            }
        }

        private SPListItem BuildListItem(JsonListItem.D jsonResult, SPList parentList)
        {
            var listItem = _itemMapper.Map(jsonResult);

            if (listItem != null)
            {
                listItem.ParentList = parentList;

                var folder = _folderMapper.Map(jsonResult.Folder);
                // reference each other
                if (folder != null)
                {
                    folder.ListItem = listItem;
                    listItem.Folder = folder;
                }

                var file = _fileMapper.Map(jsonResult.File);
                // reference each other
                if (file != null)
                {
                    file.ListItem = listItem;
                    listItem.File = file;
                }
            }

            return listItem;
        }

        private IEnumerable<SPList> BuildLists(string responseString)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return Enumerable.Empty<SPList>();
            }

            var jsonResult = Deserialize<JsonListCollection>(responseString);

            return BuildLists(jsonResult);
        }

        private IEnumerable<SPList> BuildLists(JsonListCollection jsonResult)
        {
            if (jsonResult == null)
            {
                yield return default(SPList);
            }

            foreach (var jsonList in jsonResult.d.results)
            {
                yield return BuildList(jsonList);
            }
        }

        private SPList BuildList(JsonList.D jsonList)
        {
            var list = _listMapper.Map(jsonList);

            if (list != null)
            {
                list.Views = (jsonList.Views != null) ? _viewMapper.MapList(jsonList.Views.results) : null;
                list.RootFolder = (jsonList.RootFolder != null) ? _folderMapper.Map(jsonList.RootFolder) : null;
            }

            return list;
        }
        #endregion
    }
}