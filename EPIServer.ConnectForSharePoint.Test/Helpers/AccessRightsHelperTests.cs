using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPiServer.ConnectForSharePoint.Helpers;
using Moq;
using EPiServer.ServiceLocation;
using EPiServer.DataAbstraction;
using EPiServer.Core;
using EPiServer.Security;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class AccessRightsHelperTests
    {
        Mock<IServiceLocator> serviceLocator;
        Mock<IContentSecurityRepository> contentSecurityRepo;

        private AccessRightsHelper GetUnitUnderTest()
        {
            serviceLocator = new Mock<IServiceLocator>();
            contentSecurityRepo = new Mock<IContentSecurityRepository>();
            serviceLocator.Setup(sl => sl.GetInstance<IContentSecurityRepository>()).Returns(contentSecurityRepo.Object);

            return new AccessRightsHelper(serviceLocator.Object);
        }

        [TestMethod]
        public void SetAccessRight_Calls_Save_With_Proper_Arguments()
        {
            var arHelper = GetUnitUnderTest();
            var contentRef = new ContentReference(123);
            var actualEntry = new AccessControlEntry("blah", AccessLevel.Read); //Make sure the values are different than the ones passed into the SetAccessRight method
            var secDescriptor = new Mock<IContentSecurityDescriptor>();
            secDescriptor.Setup(sd => sd.CreateWritableClone()).Returns(secDescriptor.Object);
            secDescriptor.Setup(sd => sd.AddEntry(It.IsAny<AccessControlEntry>()))
                .Callback((AccessControlEntry entry) =>
                { actualEntry = entry; }
                );

            contentSecurityRepo.Setup(csr => csr.Save(It.IsAny<ContentReference>(), It.IsAny<IContentSecurityDescriptor>(), It.IsAny<SecuritySaveType>())).Verifiable();
            contentSecurityRepo.Setup(csr => csr.Get(It.IsAny<ContentReference>())).Returns(secDescriptor.Object);

            arHelper.SetAccessRight(contentRef, AccessLevel.Edit);

            contentSecurityRepo.Verify(csr => csr.Save(It.Is<ContentReference>(cr => cr.Equals(contentRef)),
                It.IsAny<IContentSecurityDescriptor>(),
                It.Is<SecuritySaveType>(sa => sa == SecuritySaveType.Replace)),
                Times.Once,
                "Save was not called with the expected values");

            Assert.AreEqual(AccessLevel.Edit, actualEntry.Access, "Access Level different than the expected value");
            Assert.AreEqual("Everyone", actualEntry.Name, "Access entry name is not the expected value");
        }

        [TestMethod]
        public void SetAccessRight_Exits_When_The_SecurityDescriptor_Is_Null()
        {
            var arHelper = GetUnitUnderTest();
            var contentRef = new ContentReference(123);
            var actualEntry = new AccessControlEntry("blah", AccessLevel.Read); //Make sure the values are different than the ones passed into the SetAccessRight method
            var secDescriptor = new Mock<IContentSecurityDescriptor>();
            secDescriptor.Setup(sd => sd.CreateWritableClone()).Returns(null);

            contentSecurityRepo.Setup(csr => csr.Save(It.IsAny<ContentReference>(), It.IsAny<IContentSecurityDescriptor>(), It.IsAny<SecuritySaveType>())).Verifiable();
            contentSecurityRepo.Setup(csr => csr.Get(It.IsAny<ContentReference>())).Returns(secDescriptor.Object);

            arHelper.SetAccessRight(contentRef, AccessLevel.Edit);

            contentSecurityRepo.Verify(csr => csr.Save(It.IsAny<ContentReference>(), It.IsAny<IContentSecurityDescriptor>(), It.IsAny<SecuritySaveType>()),
                Times.Never,
                "Save was called unexpectedly");

        }

        [TestMethod]
        public void SetAccessRight_Exits_When_The_ContentReference_Is_Null()
        {
            var arHelper = GetUnitUnderTest();
            var actualEntry = new AccessControlEntry("blah", AccessLevel.Read); //Make sure the values are different than the ones passed into the SetAccessRight method
            var secDescriptor = new Mock<IContentSecurityDescriptor>();
            secDescriptor.Setup(sd => sd.CreateWritableClone()).Returns(secDescriptor.Object);
            secDescriptor.Setup(sd => sd.AddEntry(It.IsAny<AccessControlEntry>()))
                .Callback((AccessControlEntry entry) =>
                { actualEntry = entry; }
                );

            contentSecurityRepo.Setup(csr => csr.Save(It.IsAny<ContentReference>(), It.IsAny<IContentSecurityDescriptor>(), It.IsAny<SecuritySaveType>())).Verifiable();
            contentSecurityRepo.Setup(csr => csr.Get(It.IsAny<ContentReference>())).Returns(secDescriptor.Object);

            arHelper.SetAccessRight(null, AccessLevel.Edit);

            contentSecurityRepo.Verify(csr => csr.Save(It.IsAny<ContentReference>(), It.IsAny<IContentSecurityDescriptor>(), It.IsAny<SecuritySaveType>()),
                Times.Never,
                "Save was called unexpectedly");
        }

        // ToLocal is in a static class so it cannot be mocked without some expensive solutions
        // Leaving this here in case we want to come back to it
        //[TestMethod]
        //public void SetAccessRight_Converts_Security_To_Local_When_Security_Is_Inherited()
        //{
        //    var processor = GetUnitUnderTest();
        //    var contentRef = new ContentReference(123);
        //    var actualEntry = new AccessControlEntry("blah", AccessLevel.Read); //Make sure the values are different than the ones passed into the SetAccessRight method

        //    var secDescriptor = new Mock<IEditableSecurityDescriptor>();
        //    secDescriptor.Setup(sd => sd.CreateWritableClone()).Returns(secDescriptor.Object);
        //    secDescriptor.SetupGet<bool>(secD => secD.IsInherited).Returns(true);
        //    secDescriptor.Setup(sd => sd.ToLocal()).Verifiable();

        //    contentSecurityRepo.Setup(csr => csr.Get(It.IsAny<ContentReference>())).Returns(secDescriptor.Object as IContentSecurityDescriptor);

        //    processor.SetAccessRight(contentRef, AccessLevel.Edit);
        //    //calls ToLocal when the security descriptor says the object is Inheriting Access rights
        //    secDescriptor.Verify();
        //}

    }
}
