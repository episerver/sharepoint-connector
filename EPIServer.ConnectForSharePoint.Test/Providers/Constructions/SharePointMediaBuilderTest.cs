using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers.Constructions;
using EPiServer.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Test.Providers.Constructions
{
    [TestClass]
    public class SharePointMediaBuilderTest
    {
        [TestInitialize()]
        public void Initialize()
        {

        }

        [TestMethod]
        public void SharePointMediaBuilder_ShouldHandle_MediaContent_Type()
        {
            var sharePointMediaBuilder = new SharePointMediaBuilder();
            Assert.IsTrue(sharePointMediaBuilder.HandleType(typeof(MediaData)));
        }

        [TestMethod]
        public void SharePointMediaBuilder_ShouldHandle_ImageContent_Type()
        {
            // because the ImageData type inherited from MediaData,
            // so that the builder also need to handle this type
            var sharePointMediaBuilder = new SharePointMediaBuilder();
            Assert.IsTrue(sharePointMediaBuilder.HandleType(typeof(ImageData)));
        }
        
        [TestMethod]
        public void SharePointMediaBuilder_ShouldNotHandle_OtherTypes_But_MediaContent_Type()
        {
            var sharePointMediaBuilder = new SharePointMediaBuilder();
            Assert.IsFalse(sharePointMediaBuilder.HandleType(typeof(ContentData)));
            Assert.IsFalse(sharePointMediaBuilder.HandleType(typeof(BlockData)));
            Assert.IsFalse(sharePointMediaBuilder.HandleType(typeof(ContentFolder)));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SharePointMediaBuilder_Build_ShouldThrowException_IfContentIsNull()
        {
            var sharePointMediaBuilder = new SharePointMediaBuilder();
            sharePointMediaBuilder.Build(null, new Mock<ContentReference>().Object, new Mock<ContentReference>().Object, new Mock<ModelWrapper>(null, null, null, null).Object);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SharePointMediaBuilder_Build_ShouldThrowException_IfModelWrapperIsNull()
        {
            var sharePointMediaBuilder = new SharePointMediaBuilder();
            sharePointMediaBuilder.Build(new PageData(), new Mock<ContentReference>().Object, new Mock<ContentReference>().Object, null);
        }
    }
}
