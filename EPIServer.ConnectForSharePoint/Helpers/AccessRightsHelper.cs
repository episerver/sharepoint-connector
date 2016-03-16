using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Security;
using EPiServer.ServiceLocation;
using System;
using System.Diagnostics.CodeAnalysis;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Helper class for checking access right
    /// </summary>
    [Obsolete("Obsolete : The methods in this class are very specific and assign specified access level to the Everyone group.")]
    public class AccessRightsHelper : IAccessRightsHelper
    {
        private IServiceLocator _serviceLocator;
        private static IContentSecurityRepository _contentSecurityRepository;

        [ExcludeFromCodeCoverage]
        public AccessRightsHelper()
        {
            _serviceLocator = ServiceLocator.Current;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
        }

        internal AccessRightsHelper(IServiceLocator serviceLocator)
        {
            _serviceLocator = serviceLocator;
            _contentSecurityRepository = _serviceLocator.GetInstance<IContentSecurityRepository>();
        }


        /// <summary>
        /// Set the given access rights for the given content reference
        /// </summary>
        /// <param name="contentReference">The given content reference</param>
        /// <param name="accessLevel">The given access rights</param>
        public void SetAccessRight(ContentReference contentReference, AccessLevel accessLevel)
        {
            if (contentReference == null)
            {
                return;
            }

            var contentSecurity = _contentSecurityRepository.Get(contentReference).CreateWritableClone() as IContentSecurityDescriptor;
            if (contentSecurity == null)
            {
                return;
            }

            // The Access Control List does not support adding or modifying entries when it is set to inherited
            // So that, convert it to local
            if (contentSecurity.IsInherited)
            {
                contentSecurity.ToLocal();
            }

            contentSecurity.Clear();
            // Add access right for the default groups
            contentSecurity.AddEntry(new AccessControlEntry("Everyone", accessLevel));

            _contentSecurityRepository.Save(contentReference, contentSecurity, SecuritySaveType.Replace);
        }
    }
}