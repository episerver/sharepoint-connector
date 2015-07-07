using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Logging;
using EPiServer.Security;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Helper class for checking access right
    /// </summary>
    public class AccessRightsHelper
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        private static readonly IContentSecurityRepository _contentSecurityRepository = ServiceLocator.Current.GetInstance<IContentSecurityRepository>();

        /// <summary>
        /// Set the given access rights for the given content reference
        /// </summary>
        /// <param name="contentReference">The given content reference</param>
        /// <param name="accessLevel">The given access rights</param>
        public static void SetAccessRight(ContentReference contentReference, AccessLevel accessLevel)
        {
            if (contentReference == null)
            {
                _logger.Debug("SetAccessRight: The content link is empty");
                return;
            }

            var contentSecurity = _contentSecurityRepository.Get(contentReference).CreateWritableClone() as IContentSecurityDescriptor;
            if (contentSecurity == null)
            {
                _logger.Debug("SetAccessRight: The content security descriptor is null");
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