using System;
using System.Net;

namespace EPiServer.ConnectForSharePoint.Authentication
{
    /// <summary>
    /// Web client for Claims
    /// </summary>
    public class ClaimsWebClient : WebClient
    {
        private readonly MsOnlineClaimsHelper _claimsHelper;

        public ClaimsWebClient(Uri host, string username, string password)
        {
            _claimsHelper = new MsOnlineClaimsHelper(host, username, password);
        }

        /// <summary>
        /// Gets web request from the given address and then append cookie container
        /// </summary>
        /// <param name="address">The requested address</param>
        /// <returns>The web request that appended cookie container data</returns>
        protected override WebRequest GetWebRequest(Uri address)
        {
            var request = base.GetWebRequest(address);
            if (request is HttpWebRequest)
            {
                (request as HttpWebRequest).CookieContainer = _claimsHelper.CookieContainer;
            }

            return request;
        }
    }
}