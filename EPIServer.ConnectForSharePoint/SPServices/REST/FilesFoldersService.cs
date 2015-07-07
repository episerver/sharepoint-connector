using EPiServer.ConnectForSharePoint.Mappers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System.Collections.Generic;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.SPServices.REST
{
    [ServiceConfiguration(ServiceType = typeof(IFilesFoldersService))]
    public class FilesFoldersService : RESTServiceBase, IFilesFoldersService
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        private readonly IMapper<JsonFile.D, SPFile> _fileMapper;
        private readonly IMapper<JsonFolder.D, SPFolder> _folderMapper;
        private readonly IMapper<JsonUser.D, SPUser> _userMapper;
        private readonly IMapper<JsonListItem.D, SPListItem> _itemMapper;

        public FilesFoldersService()
            : this(new SPFileMapper(), new SPFolderMapper(), new SPUserMapper(), new SPListItemMapper())
        {
        }

        public FilesFoldersService(
            IMapper<JsonFile.D, SPFile> fileMapper,
            IMapper<JsonFolder.D, SPFolder> folderMapper,
            IMapper<JsonUser.D, SPUser> userMapper,
            IMapper<JsonListItem.D, SPListItem> itemMapper)
        {
            _fileMapper = fileMapper;
            _folderMapper = folderMapper;
            _userMapper = userMapper;
            _itemMapper = itemMapper;
        }

        public SPFolder LoadFolder(SPListItem listItem)
        {
            _logger.Information("LoadFolder from SharePoint");
            if (listItem.FileSystemObjectType != 1)
            {
                _logger.Information("The FileSystemObjectType value ({0}) is not expected type, return null", listItem.FileSystemObjectType);
                return null;
            }

            var restRequest = new RestJsonRequest(listItem.FolderUrl);
            var folder = BuildFolder(SendRequest(restRequest, null, null));

            // reference eachother
            if (folder != null)
            {
                folder.ListItem = listItem;
            }
            listItem.Folder = folder;

            _logger.Information("Compete LoadFolder, return folder: {0}", folder != null ? folder.Name : "null");
            return folder;
        }

        public SPFile LoadFile(SPListItem listItem)
        {
            _logger.Information("LoadFile from SharePoint");
            if (listItem.FileSystemObjectType != 0)
            {
                _logger.Information("The FileSystemObjectType value ({0}) is not expected type, return null", listItem.FileSystemObjectType);
                return null;
            }

            var restRequest = new RestJsonRequest(listItem.FileUrl, "GET", "expand=Author/Title,ModifiedBy/Title");
            var file = BuildFile(SendRequest(restRequest, null, null));

            // reference eachother
            if (file != null)
            {
                file.ListItem = listItem;
            }
            listItem.File = file;

            _logger.Information("Compete LoadFile, return file: {0}", file != null ? file.Name : "null");
            return file;
        }

        public SPFolder LoadParentFolder(SPFolder currentFolder)
        {
            _logger.Information("LoadParentFolder from SharePoint");
            if (currentFolder.IsRootFolder)
            {
                currentFolder.ParentFolder = currentFolder;
            }

            var restRequest = new RestJsonRequest(string.Format(Consts.REST_FOLDER_PARENT_PATH_FORMAT, EncodeDataForRESTRequest(currentFolder.ServerRelativeUrl)));
            var folder = BuildFolder(SendRequest(restRequest, null, null));

            // reference eachother
            if (folder != null)
            {
                currentFolder.ParentFolder = folder;
            }

            _logger.Information("Compete LoadParentFolder, return folder: {0}", folder != null ? folder.Name : "null");
            return folder;
        }

        public SPFolder LoadRootFolder(SPList currentList)
        {
            _logger.Information("LoadRootFolder from SharePoint");

            var restRequest = new RestJsonRequest(string.Format(Consts.REST_LIST_ROOTFOLDER_PATH_FORMAT, currentList.Id));
            var folder = BuildFolder(SendRequest(restRequest, null, null));

            // reference eachother
            if (folder != null)
            {
                currentList.RootFolder = folder;
            }

            _logger.Information("Compete LoadRootFolder, return folder: {0}", folder != null ? folder.Name : "null");
            return folder;
        }

        public SPFolder LoadFolderProperties(SPFolder currentFolder)
        {
            _logger.Information("Load folder {0} properties from SharePoint.", currentFolder.Name);

            var restRequest = new RestJsonRequest(string.Format(Consts.REST_FOLDER_PROPERTIES_PATH_FORMAT, EncodeDataForRESTRequest(currentFolder.ServerRelativeUrl)));

            UpdateFolderProperties(currentFolder, SendRequest(restRequest, null, null));

            _logger.Information("Compete LoadFoderProperties, return folder: {0}.", currentFolder.Name);
            return currentFolder;
        }

        public IEnumerable<SPFile> GetFiles(SPFolder currentFolder)
        {
            // explicitly get ListItem's EncodedAbsUrl hidden field
            var restRequest = new RestJsonRequest(string.Format(Consts.REST_FOLDER_FILES_PATH_FORMAT, EncodeDataForRESTRequest(currentFolder.ServerRelativeUrl)),
                "GET", "expand=ListItemAllFields/EncodedAbsUrl,Author,ModifiedBy");

            return BuildFiles(SendRequest(restRequest, null, null));
        }

        #region Builds SPFolder, SPFile methods

        private void UpdateFolderProperties(SPFolder currentFolder, string responseString)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return;
            }

            // NOTE: this is on our addon Team virtual machine, haven't known other premise SharePoint environments.
            // NOTE: this is weird that the string of vti_x005f_timecreated, vti_x005f_timelastmodified
            // are UTC value but have no "Z" at end, test with the premise SharePoint (addon Team virtual machine)
            var jsonResult = Deserialize<JsonFolderProperties>(responseString);

            if (jsonResult != null && jsonResult.d != null)
            {
                currentFolder.TimeCreated = jsonResult.d.vti_x005f_timecreated;
                currentFolder.TimeLastModified = jsonResult.d.vti_x005f_timelastmodified;
            }
        }

        private SPFolder BuildFolder(string responseString)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return null;
            }

            var jsonResult = Deserialize<JsonFolder>(responseString);

            return BuildFolder(jsonResult);
        }

        private SPFolder BuildFolder(JsonFolder jsonFolder)
        {
            if (jsonFolder == null)
            {
                return null;
            }

            return _folderMapper.Map(jsonFolder.d);
        }

        private SPFile BuildFile(string responseString)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return null;
            }

            var jsonResult = Deserialize<JsonFile>(responseString);

            return BuildFile(jsonResult);
        }

        private SPFile BuildFile(JsonFile jsonFile)
        {
            if (jsonFile == null)
            {
                return null;
            }

            return BuildFile(jsonFile.d);
        }

        private SPFile BuildFile(JsonFile.D jsonFile)
        {
            var file = _fileMapper.Map(jsonFile);

            if (file == null)
            {
                return file;
            }

            file.Author = _userMapper.Map(jsonFile.Author);
            file.ModifiedBy = _userMapper.Map(jsonFile.ModifiedBy);

            return file;
        }

        private IEnumerable<SPFile> BuildFiles(string responseString)
        {
            if (string.IsNullOrEmpty(responseString))
            {
                return Enumerable.Empty<SPFile>();
            }

            var jsonResult = Deserialize<JsonFileCollection>(responseString);

            return BuildFiles(jsonResult);
        }

        private IEnumerable<SPFile> BuildFiles(JsonFileCollection jsonResult)
        {
            if (jsonResult == null)
            {
                yield return default(SPFile);
            }

            foreach (var jsonItem in jsonResult.d.results)
            {
                var file = BuildFile(jsonItem);
                var listItem = _itemMapper.Map(jsonItem.ListItemAllFields);

                // reference eachother
                if (listItem != null && listItem.FileSystemObjectType == 0)
                {
                    file.ListItem = listItem;
                    listItem.File = file;
                }

                yield return file;
            }
        }
        #endregion
    }
}