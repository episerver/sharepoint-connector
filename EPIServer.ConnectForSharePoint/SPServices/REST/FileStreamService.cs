using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.Core;
using EPiServer.ImageLibrary;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using EPiServer.Web;
using EPiServer.Web.Hosting;
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
        private ISharePointContentCacheProvider _spContentCacheProvider;

        private static object _lock = new object();
        private static byte[] _defaultThumbnailData;
        /// <summary>
        /// Gets default thumbnail for any SharePoint's image
        /// </summary>
        private static byte[] DefaultThumbnailData
        {
            get
            {
                if (_defaultThumbnailData == null)
                {
                    lock (_lock)
                    {
                        if (_defaultThumbnailData == null)
                        {
                            var defaultThumbnailFilePath = ModuleHelper.ToResource("ClientResources/Images/DefaultThumbnail.png");

                            // reading resources in a zip file
                            if (VirtualPathUnifiedProvider.HasAccessToVirtualPath(defaultThumbnailFilePath))
                            {
                                using (var fileStream = VirtualPathUnifiedProvider.OpenFile(defaultThumbnailFilePath))
                                {
                                    _defaultThumbnailData = new byte[fileStream.Length];

                                    fileStream.Read(_defaultThumbnailData, 0, (int)fileStream.Length);
                                }
                            }
                        }
                    }
                }

                return _defaultThumbnailData;
            }
        }

        public FileStreamService()
            : this(
                SharePointHelper.SPServicesHub.SPListsViewsService,
                SharePointHelper.SPServicesHub.SPFilesFoldersService,
                ServiceLocator.Current.GetInstance<ISharePointContentMappingRepository>(),
                ServiceLocator.Current.GetInstance<ISharePointContentCacheProvider>(),
                null)
        {
        }

        public FileStreamService(
            IListsViewsService spListsService,
            IFilesFoldersService spFilesFoldersService,
            ISharePointContentMappingRepository spContentMappingRepository,
            ISharePointContentCacheProvider spContentCacheProvider,
            byte[] defaultThumbnailData)
        {
            _spListsService = spListsService;
            _spFilesFoldersService = spFilesFoldersService;
            _spContentMappingRepository = spContentMappingRepository;
            _spContentCacheProvider = spContentCacheProvider;
            _defaultThumbnailData = defaultThumbnailData;
        }

        /// <summary>
        /// Get file stream from SharePoint by its GUID string
        /// </summary>
        /// <param name="blobId">The given blob identification</param>
        /// <returns>File stream</returns>
        public Stream GetFileStream(Uri blobId, bool isThumbnail)
        {
            _logger.Information("GetFileStream blobId: {0} - isThumbnail: {1}.", blobId, isThumbnail);
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

                return GetFileStream(listItem, mappedIdentity.ContentLink, isThumbnail);
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
        public Stream GetFileStream(SPListItem listItem, ContentReference contentLink, bool isThumbnail)
        {
            if (listItem == null || listItem.File == null)
            {
                _logger.Information("GetFileStream: listItem is null, return null.");
                return null;
            }

            // if require thumbnail and always using generic default icon(image) or 
            // the item does not support thumbnail implicitly
            if (isThumbnail && (SharePointSettings.Current.ThumbnailOption == ThumbnailOptions.GenericThumbnailForAll || (
                SharePointSettings.Current.ThumbnailOption == ThumbnailOptions.OnlySharePointThumbnail && !HasThumbnailImplicitly(listItem))))
            {
                return OpenGenericDefaultImage();
            }

            var relativeUrl = GetFileRelativeUrl(listItem, isThumbnail);
            // Try get data memory stream from SharePoint site first
            var memoryStream = GetFileSteam(relativeUrl, isThumbnail);

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
                    // return a replaced stream if thumbnail required
                    return isThumbnail ? OpenGenericDefaultImage() : memoryStream;
                }

                // get refreshed file URL then stream
                relativeUrl = GetFileRelativeUrl(listItem, isThumbnail);
                _logger.Information("The file has been renamed, try to read it based on the new Url.");
                memoryStream = GetFileSteam(relativeUrl, isThumbnail);
            }

            return memoryStream;
        }

        public Stream GetFileSteam(string relativeUrl, bool isThumbnail)
        {
            // get stream from SharePoint site
            var stream = OpenDirectFile(relativeUrl);
            if (stream != null)
            {
                _logger.Information("Return binary data for file, relativeUrl: {0} - isThumbnail: {1}.", relativeUrl, isThumbnail);
                return (!isThumbnail) ? stream :
                    MakeThumbnail(stream, relativeUrl);
            }
            return null;
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
        /// Generate default icon (image) 
        /// </summary>
        /// <returns></returns>
        private MemoryStream OpenGenericDefaultImage()
        {
            return new MemoryStream(DefaultThumbnailData);
        }

        /// <summary>
        /// Gets thumbnail data by ImageServiceClient
        /// </summary>
        /// <param name="reader">The given Stream reader</param>
        /// <param name="itemFileName">The given file name</param>
        /// <returns></returns>
        private MemoryStream MakeThumbnail(Stream reader, string itemFileName)
        {
            if (reader == null)
            {
                return null;
            }

            var originalImageBuffer = new byte[checked((long)reader.Length)];
            reader.Read(originalImageBuffer, 0, (int)reader.Length);

            var imageOperations = new List<ImageOperation>() {
                new ImageOperation(ImageEditorCommand.ResizeKeepScale, 48, 48)
            };

            // Render new image
            var byteArray = new ImageServiceClient("ImageServiceClientEndPoint").RenderImage(originalImageBuffer, imageOperations, MimeMapping.GetMimeMapping(itemFileName), 1f, 50);

            return new MemoryStream(byteArray);
        }

        /// <summary>
        /// Delete mapping to a sharepoint file and flush caches related to that file.
        /// </summary>
        /// <param name="serverRelativeUrl">The server relative URL.</param>
        /// <param name="contentLink">The content link of file</param>
        private void RemoveFile(string serverRelativeUrl, ContentReference contentLink)
        {
            // Note: a file (associate with a list item) can be in different folders (of different views of a list)
            // then need to remove the file from all folders contain it, the file relative URL is unique
            _spContentCacheProvider.RemoveFileFromCache(contentLink, serverRelativeUrl);

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

        public string GetFileRelativeUrl(SPListItem listItem, bool isThumbnail)
        {
            if (listItem.File == null)
            {
                return string.Empty;
            }

            var serverRelativeUrl = EncodeDataForRESTRequest(listItem.File.ServerRelativeUrl);

            return (isThumbnail && HasThumbnailImplicitly(listItem)) ? GetSharePointThumbnailUrl(serverRelativeUrl) : serverRelativeUrl;
        }

        private bool HasThumbnailImplicitly(SPListItem listItem)
        {
            if (listItem.ParentList == null)
            {
                return false;
            }

            return (listItem.ParentList.BaseTemplate == (int)SupportedBaseTemplates.AssetLibary ||
               listItem.ParentList.BaseTemplate == (int)SupportedBaseTemplates.PictureLibrary);
        }
    }
}