using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.Core;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;

namespace EPiServer.ConnectForSharePoint.SPServices.REST
{
    [ServiceConfiguration(ServiceType = typeof(IFileStreamService))]
    public class FileStreamService : RESTServiceBase, IFileStreamService
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        private IListsViewsService _spListsService;
        private IFilesFoldersService _spFilesFoldersService;
        private ISharePointContentMappingRepository _spContentMappingRepository;

        private static object _lock = new object();

        public FileStreamService()
            : this(
                SharePointHelper.SPServicesHub.SPListsViewsService,
                SharePointHelper.SPServicesHub.SPFilesFoldersService,
                ServiceLocator.Current.GetInstance<ISharePointContentMappingRepository>())
        {
        }

        public FileStreamService(
            IListsViewsService spListsService,
            IFilesFoldersService spFilesFoldersService,
            ISharePointContentMappingRepository spContentMappingRepository)
        {
            _spListsService = spListsService;
            _spFilesFoldersService = spFilesFoldersService;
            _spContentMappingRepository = spContentMappingRepository;
        }

        /// <summary>
        /// Get file stream from SharePoint by its GUID string
        /// </summary>
        /// <param name="blobId">The given blob identification</param>
        /// <returns>File stream</returns>
        public Stream GetFileStream(Uri blobId)
        {
            _logger.Information("GetFileStream blobId: {0} .", blobId);
            var memoryStream = new MemoryStream();

            try
            {
                // Get mapped identity from MappingService
                var mappedIdentity = _spContentMappingRepository.GetIdentity(Guid.Parse(blobId.Segments[1].Replace("/", string.Empty)));
                if (mappedIdentity == null || mappedIdentity.ExternalIdentifier == null)
                {
                    _logger.Information("Could not find MappedIdentity, return empty MemoryStream.");
                    return memoryStream;
                }

                // Segments[1] is providerUniquePath of IdentityMappingService,
                // the value made by ListID~ViewID~ListItemID(optional)
                var segments = mappedIdentity.ExternalIdentifier.Segments[1].ToString().Split('~');
                if (segments.Count() != 3)
                {
                    _logger.Information("mappedIdentity.ExternalIdentifier.Segments[1] value is wrong format ({0}), return empty MemoryStream.", mappedIdentity.ExternalIdentifier.Segments[1]);
                    return memoryStream;
                }

                // segments[0]: List identification
                // segments[2]: ListItem identification
                var parentList = _spListsService.GetList(segments[0]);
                var listItem = (parentList != null) ? _spListsService.GetListItem(parentList, segments[2]) : null;

                return GetFileStream(listItem, mappedIdentity.ContentLink);
            }
            catch (Exception ex)
            {
                _logger.Error("Error occurs when getting file stream from SharePoint", ex);
            }

            return memoryStream;
        }

        /// <summary>
        /// Get file stream of a SharePoint's ListItem
        /// </summary>
        /// <param name="listItem">The file item from share point site</param>
        /// <param name="contentLink">The content link has been mapped identity</param>
        /// <returns></returns>
        public Stream GetFileStream(SPListItem listItem, ContentReference contentLink)
        {
            if (listItem == null || listItem.File == null)
            {
                _logger.Information("GetFileStream: listItem is null, return null.");
                return null;
            }

            var relativeUrl = GetFileRelativeUrl(listItem);
            // Try get data memory stream from SharePoint site first
            var memoryStream = GetFileSteam(relativeUrl);

            // The file was renamed or even the item was deleted then try to reload the item (along with the file)
            if (memoryStream == null)
            {
                _logger.Information("Reload data from SharePoint to check whether the file has been renamed or not.");
                // Note: the same file is "duplicatedly" to different SPListItem object and cached in different parent folders,
                // then this code happens for each contentLink. Must recall that SharePoint views are for groupping list items by certain conditions
                // and avoid having the same list items in different views in mind
                var parent = listItem.ParentList;
                listItem = _spListsService.GetListItem(parent, listItem.Id.ToString());

                if (listItem == null || listItem.File == null)
                {
                    _logger.Information("The file has been deleted, remove it from cache.");
                    // code removes the item before returning
                    RemoveFile(relativeUrl, contentLink);
                    return memoryStream;
                }

                // get refreshed file URL then stream
                relativeUrl = GetFileRelativeUrl(listItem);
                _logger.Information("The file has been renamed, try to read it based on the new Url.");
                memoryStream = GetFileSteam(relativeUrl);
            }

            return memoryStream;
        }

        public Stream GetFileSteam(string relativeUrl)
        {
            // get stream from SharePoint site
            return OpenDirectFile(relativeUrl);
        }

        private static string[] GetItemExpandQuery(bool getVideoInfo)
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
            if (getVideoInfo)
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
            return new string[2] { selectClause.ToString(), expandClause };
        }

        public string GetListItemJsonData(SPListItem item)
        {
            SPList list = item.ParentList;
            SPListView view = null;
            if (item.selectedView == null)
                view = item.ParentList.Views.First();
            else
                view = item.selectedView;

            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_GETITEMS_PATH_FORMAT, list.Id), "POST", null);
            var bodyBuilder = new StringBuilder("{ 'query' : {'__metadata': { 'type': 'SP.CamlQuery' }, ");

            var fixedString = view.ListViewXml.Replace("<ViewFields>", "<ViewFields Properties=\"true\">");
            fixedString = fixedString.Replace("<View ", "<View Scope=\"RecursiveAll\"" );           // search sub folders too.

            // We should probably be using xmldocument in this method....
            string withQueryString;
            if (fixedString.Contains("<Query>"))
                withQueryString = fixedString.Replace("<Query>", "<Query><Where><Eq>" +
                        "<FieldRef Name=\"ID\"/><Value Type =\"Lookup\">" + item.Id + "</Value>" +
                        "</Eq></Where>");
            else // custom views with no ordering have <Query/>
            {
                withQueryString = fixedString.Replace("<Query />", "<Query><Where><Eq>" +
                        "<FieldRef Name=\"ID\"/><Value Type =\"Lookup\">" + item.Id + "</Value>" +
                        "</Eq></Where></Query>");
            }

            bodyBuilder.AppendFormat("'ViewXml': '{0}'", EncodeDataForCAMLQuery(withQueryString));

            // NOTE: this is on our addon Team virtual machine, haven't known other premise SharePoint environments.
            // NOTE: only with POST has this way defining DatesInUtc in the request 
            // NOTE: without DatesInUtc==true the premise SharePoint (addon Team virtual machine) returns local DateTime string 
            // (a DateTime value string without 'Z' at end and applied Regional Settings of SiteCollection already)
            // because without Z the desirialize one more time coverts to server local time. Force DatesInUtc:
            bodyBuilder.Append(",'DatesInUtc': 'true'");
            bodyBuilder.Append(" } }");

            String str = SendRequest(restRequest, bodyBuilder.ToString(), null, true);
            // remove this Id because it causes Json parsing issues... not sure why sharepoint is returning two id values
            var str2 = str.Replace("\"FileSystemObjectType\":0,\"Id\"", "\"FileSystemObjectType\":0,\"IDDD\"");
            return str2;
        }

        public Stream GetListItemXmlData(string url)
        {
            // get stream from SharePoint site
            return OpenDirectListItemXmlData(url);
        }

        /// <summary>
        /// Open and read binary data directly from SharePoint site
        /// </summary>
        /// <param name="url">The complete rest url to fetch an item from a custom list</param>
        /// <returns></returns>
        private MemoryStream OpenDirectListItemXmlData(string url)
        {
            try
            {
                var reqInfo = new RestJsonRequest(url, "GET", null);
                var headers = new NameValueCollection();
                ////headers.Add("binaryStringResponseBody", "true");

                var memoryStream = new MemoryStream();

                // Get binary data from SharePoint
                using (var stream = GetResponseStream(reqInfo, null, headers))
                {
                    stream.CopyTo(memoryStream); // we do this to support seeking
                    stream.Flush();
                }

                memoryStream.Seek(0, SeekOrigin.Begin);

                return memoryStream;
            }
            catch (Exception ex)
            {
                _logger.Error("Error occurs when reading binary data directly from SharePoint site", ex);
                return null;
            }
        }

        /// <summary>
        /// Open and read binary data directly from SharePoint site
        /// </summary>
        /// <param name="filePath">The image path, thumbnail or orginal relative url</param>
        /// <returns></returns>
        private MemoryStream OpenDirectFile(string fileRelativeUrl)
        {
            try
            {
                var reqInfo = new RestJsonRequest(string.Format(Consts.REST_OPENBINARYSTREAM_FORMAT, fileRelativeUrl));
                var headers = new NameValueCollection();
                headers.Add("binaryStringResponseBody", "true");

                var memoryStream = new MemoryStream();
                // Get binary data from SharePoint
                using (var stream = GetResponseStream(reqInfo, null, headers))
                {
                    stream.CopyTo(memoryStream); // we do this to support seeking
                    stream.Flush();
                }

                memoryStream.Seek(0, SeekOrigin.Begin);

                return memoryStream;
            }
            catch (Exception ex)
            {
                _logger.Error("Error occurs when reading binary data directly from SharePoint site", ex);
                return null;
            }
        }

        /// <summary>
        /// Delete mapping to a sharepoint file and flush caches related to that file.
        /// </summary>
        /// <param name="serverRelativeUrl">The server relative URL.</param>
        /// <param name="contentLink">The content link of file</param>
        private void RemoveFile(string serverRelativeUrl, ContentReference contentLink)
        {
            _spContentMappingRepository.Delete(contentLink);
        }

        /// <summary>
        /// Gets thumbnail url from the the given SharePoint's image url
        /// </summary>
        /// <param name="imageUrl">The given SharePoint's image url</param>
        /// <returns></returns>
        private string GetSharePointThumbnailUrl(string imageUrl)
        {
            if (string.IsNullOrWhiteSpace(imageUrl))
            {
                return string.Empty;
            }

            var lastSlashPos = imageUrl.LastIndexOf('/');
            if (lastSlashPos < 0)
            {
                return imageUrl;
            }

            var nameNoExtension = Path.GetFileNameWithoutExtension(imageUrl);
            var extension = Path.GetExtension(imageUrl);
            // file without extension
            if (string.IsNullOrWhiteSpace(nameNoExtension) || string.IsNullOrWhiteSpace(extension))
            {
                return imageUrl;
            }

            var thumbnailUrl = new StringBuilder(imageUrl.Substring(0, lastSlashPos + 1));
            // Note: SharePoint thumbnail is generated in .jpg extension
            thumbnailUrl.AppendFormat("_t/{0}_{1}.jpg", nameNoExtension, extension.Remove(0, 1));

            return thumbnailUrl.ToString();
        }

        public string GetFileRelativeUrl(SPListItem listItem)
        {
            if (listItem.File == null)
            {
                return string.Empty;
            }

            return EncodeDataForRESTRequest(listItem.File.ServerRelativeUrl);
        }
    }
}