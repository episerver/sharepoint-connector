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
    public class SharePointImageBuilderTest
    {
        [TestMethod]
        public void SharePointImageBuilder_ShouldHandle_ContentImage_Type()
        {
            var sharePointImageBuilder = new SharePointImageBuilder();
            Assert.IsTrue(sharePointImageBuilder.HandleType(typeof(ImageData)));
        }

        [TestMethod]
        public void SharePointImageBuilder_ShouldNotHandle_OtherTypes_But_ContentImage_Type()
        {
            var sharePointImageBuilder = new SharePointImageBuilder();
            Assert.IsFalse(sharePointImageBuilder.HandleType(typeof(ContentData)));
            Assert.IsFalse(sharePointImageBuilder.HandleType(typeof(BlockData)));
            Assert.IsFalse(sharePointImageBuilder.HandleType(typeof(MediaData)));
            Assert.IsFalse(sharePointImageBuilder.HandleType(typeof(ContentFolder)));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SharePointImageBuilder_Build_ShouldThrowException_IfContentIsNull()
        {
            var sharePointImageBuilder = new SharePointImageBuilder();
            sharePointImageBuilder.Build(null, new Mock<ContentReference>().Object, new Mock<ContentReference>().Object, new Mock<ModelWrapper>(null, null, null, null).Object);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void SharePointImageBuilder_Build_ShouldThrowException_IfModelWrapperIsNull()
        {
            var sharePointImageBuilder = new SharePointImageBuilder();
            sharePointImageBuilder.Build(new PageData(), new Mock<ContentReference>().Object, new Mock<ContentReference>().Object, null);
        }

    }
}
