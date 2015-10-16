using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ServiceLocation;
using Moq;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.Models;
using System.Runtime.Caching;
using System.Collections.Generic;
using System;
using EPiServer.ConnectForSharePoint.Models.ViewModels;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class SharePointProcessorHelperTests
    {
        Mock<IServiceLocator>       _locator;
        Mock<IFileStreamService>    _fileStreamService;
        Mock<IListsViewsService>    _listViewsService;

        private SharePointProcessorHelper GetUnitUnderTest()
        {
            _locator = new Mock<IServiceLocator>();
            _fileStreamService = new Mock<IFileStreamService>();
            _listViewsService = new Mock<IListsViewsService>();
            _locator.Setup(sl => sl.GetInstance<IFileStreamService>()).Returns(_fileStreamService.Object);
            _locator.Setup(sl => sl.GetInstance<IListsViewsService>()).Returns(_listViewsService.Object);

            return new SharePointProcessorHelper(_locator.Object);
        }

        [TestMethod]
        public void SharePointProcessorHelper_GetDocumentStream()
        {
            var helper = GetUnitUnderTest();
            string str = @"ExpectedString";

            // verify that the filestream methods are called and that the relative url is passed into getfilestream
            _fileStreamService.Setup(fs => fs.GetFileRelativeUrl(It.IsAny<SPListItem>())).Verifiable();
            _fileStreamService.Setup(fs => fs.GetFileRelativeUrl(It.IsAny<SPListItem>())).Returns(str);
            _fileStreamService.Setup(fs => fs.GetFileSteam(It.Is<string>(st => st.Equals(str)))).Verifiable();

            helper.GetDocumentStream(new SPListItem());

            _fileStreamService.Verify();

        }

        [TestMethod]
        public void SharePointProcessorHelper_GetRawListItemData()
        {
            var helper = GetUnitUnderTest();
            string str = @"\r\n CR and LF test \r fasdf \n asdf";
            string lineSeparator = ((char)0x2028).ToString();
            string paragraphSeparator = ((char)0x2029).ToString();
            str = str + lineSeparator + " asdf " + paragraphSeparator + "\\u200b" + "asdf";

            // verify that the method gets called and the string returned has the special chars removed.
            _fileStreamService.Setup(fs => fs.GetListItemJsonData(It.IsAny<SPListItem>())).Verifiable();
            _fileStreamService.Setup(fs => fs.GetListItemJsonData(It.IsAny<SPListItem>())).Returns(str);

            var rawdata = helper.GetRawListItemData(new SPListItem());

            _fileStreamService.Verify();

            // Validate that the method removed the chars that cause a failure in parsing.
            Assert.AreEqual(-1, rawdata.IndexOf("\r\n"), "CRLF still exist in string");
            Assert.AreEqual(-1, rawdata.IndexOf("\r"), "CR still exist in string");
            Assert.AreEqual(-1, rawdata.IndexOf("\n"), "LF still exist in string");
            Assert.AreEqual(-1, rawdata.IndexOf("lineSeparator"), "lineSeparator still exist in string");
            Assert.AreEqual(-1, rawdata.IndexOf("paragraphSeparator"), "paragraphSeparator still exist in string");
            Assert.AreEqual(-1, rawdata.IndexOf("\\u200b"), "\\u200b still exist in string");
        }

        [TestMethod]
        public void GetMetaData_InvalidateCache()
        {
            var helper = GetUnitUnderTest();
            SPList list = new SPList() { Id = Guid.Empty };
            Dictionary<string, SPFieldItem> emptyDictionary = new Dictionary<string, SPFieldItem>();
            Dictionary<string, SPFieldItem> oneItemDictionary = new Dictionary<string, SPFieldItem>();
            oneItemDictionary.Add(@"NameOfFirstIem", new SPFieldItem() );

            MemoryCache.Default.Add(list.Id.ToString(), emptyDictionary, new CacheItemPolicy());

            _listViewsService.Setup(fs => fs.GetFieldItemDictionary(It.IsAny<Guid>())).Verifiable();
            _listViewsService.Setup(fs => fs.GetFieldItemDictionary(It.IsAny<Guid>())).Returns(oneItemDictionary);

            var dictionary = helper.GetMetaData(new SPList() { Id = Guid.Empty}, true);

            _listViewsService.Verify();
            Assert.AreEqual(1, dictionary.Count, "Count of items in dictionary is not the same.");
        }

        [TestMethod]
        public void ParseRawListItemData()
        {
            string rawdata = "{\"d\":{\"results\":[{\"__metadata\":{\"id\":\"Web / Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6') / Items(17)\",\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)\",\"etag\":\"\\\"8\\\"\",\"type\":\"SP.Data.Herbs_x0020_Custom_x0020_ListsListItem\"},\"FirstUniqueAncestorSecurableObject\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/FirstUniqueAncestorSecurableObject\"}},\"RoleAssignments\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/RoleAssignments\"}},\"AttachmentFiles\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/AttachmentFiles\"}},\"ContentType\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/ContentType\"}},\"FieldValuesAsHtml\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/FieldValuesAsHtml\"}},\"FieldValuesAsText\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/FieldValuesAsText\"}},\"FieldValuesForEdit\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/FieldValuesForEdit\"}},\"File\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/File\"}},\"Folder\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/Folder\"}},\"ParentList\":{\"__deferred\":{\"uri\":\"http://nicenewvm/sites/SP2013QA/_api/Web/Lists(guid'cf57b2f2-3c57-42c4-825c-b65e54b3d6e6')/Items(17)/ParentList\"}},\"FileSystemObjectType\":0,\"IDDD\":17,\"ID\":17,\"Title\":\"Amendola\",\"richtextcolumn\":null,\"Created\":\"2015-09-29T17:28:16Z\",\"Modified\":\"2015-10-01T15:45:07Z\"}]}}";
            Dictionary<string, SPFieldItem> dictionary = new Dictionary<string, SPFieldItem>();

            // Just a quick test to make sure that the parsing of an item doesnt crash the parser.
            var helper = GetUnitUnderTest();
            List<ListItemField> items = helper.ParseRawListItemData(dictionary, rawdata);
        }
    }
}
