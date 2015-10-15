using Moq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ServiceLocation;
using EPiServer.Core;
using EPiServer.Security;
using EPiServer.DataAbstraction;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class ContentProviderHelperTests
    {
        private Mock<IServiceLocator> serviceLocator;
        private Mock<IContentSecurityRepository> contentSecurityRepo;
        private Mock<IContentRepository> contentRepo;

        private ContentProviderHelper GetUnitUnderTest()
        {
            serviceLocator = new Mock<IServiceLocator>();
            contentSecurityRepo = new Mock<IContentSecurityRepository>();
            contentRepo = new Mock<IContentRepository>();
            serviceLocator.Setup(sl => sl.GetInstance<IContentSecurityRepository>()).Returns(contentSecurityRepo.Object);
            serviceLocator.Setup(sl => sl.GetInstance<IContentRepository>()).Returns(contentRepo.Object);

            return new ContentProviderHelper(serviceLocator.Object);
        }
    }
}
