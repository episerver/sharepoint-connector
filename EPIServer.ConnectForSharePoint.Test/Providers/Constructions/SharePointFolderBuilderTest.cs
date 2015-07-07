using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Providers.Constructions;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Models;
using Moq;

namespace EPiServer.ConnectForSharePoint.Test
{
    [TestClass]
    public class SharePointFolderBuilderTest
    {
        [TestInitialize()]
        public void Initialize()
        {
            
        }

        [TestMethod]
        public void SharePointFolderBuilder_ShouldHandle_ContentFolder_Type()
        {
            var sharePointFolderBuilder = new SharePointFolderBuilder();
            Assert.IsTrue(sharePointFolderBuilder.HandleType(typeof(ContentFolder)));
        }


        [TestMethod]
        public void SharePointFolderBuilder_ShouldNotHandle_OtherTypes_But_ContentFolder_Type()
        {
            var sharePointFolderBuilder = new SharePointFolderBuilder();
            Assert.IsFalse(sharePointFolderBuilder.HandleType(typeof(ContentData)));
            Assert.IsFalse(sharePointFolderBuilder.HandleType(typeof(BlockData)));
            Assert.IsFalse(sharePointFolderBuilder.HandleType(typeof(MediaData)));
            Assert.IsFalse(sharePointFolderBuilder.HandleType(typeof(ImageData)));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SharePointFolderBuilder_Build_ShouldThrowException_IfContentIsNull()
        {
            var sharePointFolderBuilder = new SharePointFolderBuilder();
            sharePointFolderBuilder.Build(null, new Mock<ContentReference>().Object, new Mock<ContentReference>().Object, new Mock<ModelWrapper>(null, null, null, null).Object);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SharePointFolderBuilder_Build_ShouldThrowException_IfModelWrapperIsNull()
        {
            var sharePointFolderBuilder = new SharePointFolderBuilder();
            sharePointFolderBuilder.Build(new PageData(), new Mock<ContentReference>().Object, new Mock<ContentReference>().Object, null);
        }
    }
}
