using EPiServer.ConnectForSharePoint.Mappers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;

namespace EPiServer.ConnectForSharePoint.Test
{
    [TestClass]
    public class FilesFoldersServiceTest
    {
        Mock<IMapper<JsonFile.D, SPFile>> _fileMapper;
        Mock<IMapper<JsonFolder.D, SPFolder>> _folderMapper;
        Mock<IMapper<JsonUser.D, SPUser>> _userMapper;
        Mock<IMapper<JsonListItem.D, SPListItem>> _itemMapper;

        Mock<FilesFoldersService> _filesFoldersService;

        SPList _list;
        SPListItem _fileItem;
        SPListItem _folderItem;

        [TestInitialize()]
        public void Initialize()
        {
            // Init List
            _list = new SPList
            {
                Id = Guid.NewGuid()
            };

            // Init File
            _fileItem = new SPListItem
            {
                FileUrl = "abc.gif"
            };

            // Init Folder
            _folderItem = new SPListItem
            {
                FolderUrl = "SP"
            };

            // Init dependency and mock functions
            _fileMapper = new Mock<IMapper<JsonFile.D, SPFile>>();
            _folderMapper = new Mock<IMapper<JsonFolder.D, SPFolder>>();
            _userMapper = new Mock<IMapper<JsonUser.D, SPUser>>();
            _itemMapper = new Mock<IMapper<JsonListItem.D, SPListItem>>();

            _folderMapper.Setup(s => s.Map(It.IsAny<JsonFolder.D>())).Returns(new SPFolder());
            _fileMapper.Setup(s => s.Map(It.IsAny<JsonFile.D>())).Returns(new SPFile());
            _userMapper.Setup(s => s.Map(It.IsAny<JsonUser.D>())).Returns(new SPUser());
            _itemMapper.Setup(s => s.Map(It.IsAny<JsonListItem.D>())).Returns(new SPListItem());

            // Mock for SendRequest method of base class            
            var jsonString = "Mock";
            _filesFoldersService = new Mock<FilesFoldersService>(_fileMapper.Object, _folderMapper.Object, _userMapper.Object, _itemMapper.Object) { CallBase = true };
            _filesFoldersService.Setup(x => x.SendRequest(It.IsAny<IRestRequest>(), null, null)).Returns(jsonString);
            _filesFoldersService.Setup(x => x.SendRequest(It.IsAny<IRestRequest>(), null, null, false)).Returns(jsonString);
            _filesFoldersService.Setup(x => x.Deserialize<JsonFolder>(It.IsAny<string>())).Returns(new JsonFolder());
            _filesFoldersService.Setup(x => x.Deserialize<JsonFile>(It.IsAny<string>())).Returns(MockJsonFile);
        }

        [TestMethod]
        public void LoadFolder_is_not_folder_type_then_expected_is_null()
        {
            Assert.IsNull(_filesFoldersService.Object.LoadFolder(_folderItem));
        }

        [TestMethod]
        public void LoadFolder_is_folder_type_then_expected_not_null()
        {
            _folderItem.FileSystemObjectType = 1;

            Assert.IsNotNull(_filesFoldersService.Object.LoadFolder(_folderItem));
        }

        [TestMethod]
        public void LoadFile_is_not_file_type_then_expected_is_null()
        {
            _fileItem.FileSystemObjectType = 1;
            Assert.IsNull(_filesFoldersService.Object.LoadFile(_fileItem));
        }

        [TestMethod]
        public void LoadFile_is_file_type_then_expected_not_null()
        {
            _fileItem.FileSystemObjectType = 0;

            Assert.IsNotNull(_filesFoldersService.Object.LoadFile(_fileItem));
        }
        
        [TestMethod]
        public void LoadRootFolder_of_list_expected_is_null()
        {
            _filesFoldersService.Setup(x => x.SendRequest(It.IsAny<IRestRequest>(), null, null)).Returns(string.Empty);
            Assert.IsNull(_filesFoldersService.Object.LoadRootFolder(_list));
        }

        [TestMethod]
        public void LoadRootFolder_of_list_then_expected_not_null()
        {
            Assert.IsNotNull(_filesFoldersService.Object.LoadRootFolder(_list));
        }

        [TestMethod]
        public void LoadParentFolder_expected_is_null()
        {
            var spFolder = new SPFolder
            {
                ParentFolder = null,
                ServerRelativeUrl = "Test"
            };

            _filesFoldersService.Setup(x => x.SendRequest(It.IsAny<IRestRequest>(), null, null)).Returns(string.Empty);
            Assert.IsNull(_filesFoldersService.Object.LoadParentFolder(spFolder));
        }

        [TestMethod]
        public void LoadParentFolder_of_folder_then_expected_not_null()
        {
            var spFolder = new SPFolder
            {
                ParentFolder = null,
                ServerRelativeUrl = "Test"
            };

            Assert.IsNotNull(_filesFoldersService.Object.LoadParentFolder(spFolder));
        }
        
        private JsonFile MockJsonFile()
        {
            var f = new JsonFile();
            f.d = new JsonFile.D();
            f.d.Author = new JsonUser().d;
            f.d.ModifiedBy = new JsonUser().d;
            return f;
        }
    }
}
