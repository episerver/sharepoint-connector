using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint
{
    public class Consts
    {
        /// <summary>
        /// The REST endpoint of SharePoint contextinfo "/_api/contextinfo"
        /// </summary>
        public const string REST_CONTEXTINFO_PATH = "/_api/contextinfo";

        /// <summary>
        /// The RequestDigest header for POST webrequest
        /// </summary>
        public const string REQUESTDIGEST_HEADER_KEY = "X-RequestDigest";

        /// <summary>
        /// The REST endpoint of lists "/_api/web/lists"
        /// </summary>
        public const string REST_LISTS_PATH = "/_api/web/lists";

        /// <summary>
        /// The REST endpoint format of list "/_api/web/lists(guid'[List Id]')"
        /// </summary>
        public const string REST_LIST_PATH_FORMAT = "/_api/web/lists(guid'{0}')";

        /// <summary>
        /// The REST endpoint format of list's views "/_api/web/lists(guid'[List Id]')/views"
        /// </summary>
        public const string REST_LIST_VIEWS_PATH_FORMAT = "/_api/web/lists(guid'{0}')/views";

        /// <summary>
        /// The REST endpoint format of list's items "/_api/web/lists(guid'[List Id]')/items";
        /// </summary>
        public const string REST_LIST_ITEMS_PATH_FORMAT = "/_api/web/lists(guid'{0}')/items";

        /// <summary>
        /// The REST endpoint format of list's getitems method "/_api/web/lists(guid'[List Id]')/getitems"
        /// </summary>
        public const string REST_LIST_GETITEMS_PATH_FORMAT = "/_api/web/lists(guid'{0}')/getitems";

        /// <summary>
        /// The REST endpoint format of list's rootfolder "/_api/web/lists(guid'[List Id]')/rootfolder"
        /// </summary>
        public const string REST_LIST_ROOTFOLDER_PATH_FORMAT = "/_api/web/lists(guid'{0}')/rootfolder";

        /// <summary>
        /// The REST endpoint format of listitem "/_api/web/lists(guid'[List Id]')/items([Item Id])";
        /// </summary>
        public const string REST_LISTITEM_PATH_FORMAT = "/_api/web/lists(guid'{0}')/items({1})";

        /// <summary>
        /// The REST endpoint to get file stream /_api/Web/GetFileByServerRelativeUrl('[relative path]')/openbinarystream
        /// </summary>
        public const string REST_OPENBINARYSTREAM_FORMAT = "/_api/Web/GetFileByServerRelativeUrl('{0}')/openbinarystream";

        /// <summary>
        /// The REST endpoint format of folder "/_api/web/GetFolderByServerRelativeUrl('[relative path]')"
        /// </summary>
        public const string REST_FOLDER_PATH_FORMAT = "/_api/web/GetFolderByServerRelativeUrl('{0}')";

        /// <summary>
        /// The REST endpoint format of folder's parent "/_api/web/GetFolderByServerRelativeUr('[relative path]')/parentfolder"
        /// </summary>
        public const string REST_FOLDER_PARENT_PATH_FORMAT = "/_api/web/GetFolderByServerRelativeUrl('{0}')/parentfolder";

        /// <summary>
        /// The REST endpoint format of folder' proerties "/_api/web/GetFolderByServerRelativeUrl('[relative path]')/properties"
        /// </summary>
        public const string REST_FOLDER_PROPERTIES_PATH_FORMAT = "/_api/web/GetFolderByServerRelativeUrl('{0}')/properties";

        /// <summary>
        /// The REST endpoint format of folder' files "/_api/web/GetFolderByServerRelativeUrl('[relative path]')/properties"
        /// </summary>
        public const string REST_FOLDER_FILES_PATH_FORMAT = "/_api/web/GetFolderByServerRelativeUrl('{0}')/files";
        
        /// <summary>
        /// The prefix for SharePoint cache key
        /// </summary>
        public const string SHAREPOINT_CACHE_KEY_FORMAT = "SharePointCache_{0}";

        /// <summary>
        /// Name to create MemoryCache instance
        /// </summary>
        public const string SHAREPOINT_CONTENT_CACHE_NAME = "SharePointContentCache";
    }
}