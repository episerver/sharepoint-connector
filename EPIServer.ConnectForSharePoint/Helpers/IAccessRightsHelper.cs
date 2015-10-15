using EPiServer.Core;
using EPiServer.Security;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface IAccessRightsHelper
    {
        void SetAccessRight(ContentReference contentReference, AccessLevel accessLevel);
    }
}
