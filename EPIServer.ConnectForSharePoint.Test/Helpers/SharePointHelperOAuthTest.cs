using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Specialized;

namespace EPiServer.ConnectForSharePoint.Test.Helpers
{
    [TestClass]
    public class SharePointHelperOAuthTest
    {
        /// <summary>
        /// A fake <see cref="IOAuthTokenClient"/> that captures the request URL and parameters
        /// instead of making a real HTTP call, then returns a fixed JSON token response.
        /// </summary>
        private class FakeTokenClient : IOAuthTokenClient
        {
            public string CapturedUrl { get; private set; }
            public NameValueCollection CapturedParams { get; private set; }

            public string RequestToken(string requestUrl, NameValueCollection parameters)
            {
                CapturedUrl = requestUrl;
                CapturedParams = parameters;

                return "{\"token_type\":\"Bearer\",\"expires_in\":\"3600\",\"access_token\":\"fake-token\"}";
            }
        }

        private SharePointSettings CreateModernAuthSettings(
            string siteUrl = "https://contoso.sharepoint.com/sites/media",
            string clientId = "11111111-2222-3333-4444-555555555555",
            string clientSecret = "test-secret",
            string realm = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee")
        {
            return new SharePointSettings
            {
                IsModernAuth = true,
                IsSharePointOnline = true,
                SiteUrl = siteUrl,
                ClientId = clientId,
                ClientSecret = clientSecret,
                Realm = realm
            };
        }

        [TestMethod]
        public void GetOAuthResponse_ReturnsNull_WhenModernAuthDisabled()
        {
            var settings = new SharePointSettings { IsModernAuth = false };
            // GetOAuthResponse checks IsModernAuth before touching the cache,
            // so this path is safe to call without EPiServer runtime.
            var result = SharePointHelper.GetOAuthResponse(settings);
            Assert.IsNull(result);
        }

        [TestMethod]
        public void RequestOAuthToken_UsesEntraIdTokenEndpoint()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings(realm: "my-tenant-id");

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual("https://login.microsoftonline.com/my-tenant-id/oauth2/v2.0/token", fake.CapturedUrl);
        }

        [TestMethod]
        public void RequestOAuthToken_SendsCorrectScope()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings(siteUrl: "https://contoso.sharepoint.com/sites/media");

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual("https://contoso.sharepoint.com/.default", fake.CapturedParams["scope"]);
        }

        [TestMethod]
        public void RequestOAuthToken_SendsPlainClientId()
        {
            var fake = new FakeTokenClient();
            var clientId = "11111111-2222-3333-4444-555555555555";
            var settings = CreateModernAuthSettings(clientId: clientId);

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual(clientId, fake.CapturedParams["client_id"]);
        }

        [TestMethod]
        public void RequestOAuthToken_SendsClientCredentialsGrantType()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings();

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual("client_credentials", fake.CapturedParams["grant_type"]);
        }

        [TestMethod]
        public void RequestOAuthToken_SendsClientSecret()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings(clientSecret: "my-secret-value");

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual("my-secret-value", fake.CapturedParams["client_secret"]);
        }

        [TestMethod]
        public void RequestOAuthToken_ReturnsDeserializedToken()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings();

            var result = SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.IsNotNull(result);
            Assert.AreEqual("Bearer", result.TokenType);
            Assert.AreEqual("fake-token", result.AccessToken);
            Assert.AreEqual("3600", result.ExpiresIn);
        }

        [TestMethod]
        public void RequestOAuthToken_ScopeStripsPathFromSiteUrl()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings(siteUrl: "https://contoso.sharepoint.com/sites/deep/path");

            SharePointHelper.RequestOAuthToken(settings, fake);

            // Scope should be host-only, not include the path
            Assert.AreEqual("https://contoso.sharepoint.com/.default", fake.CapturedParams["scope"]);
        }

        [TestMethod]
        public void RequestOAuthToken_ScopeUsesHttpsEvenWhenSiteUrlIsHttp()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings(siteUrl: "http://contoso.sharepoint.com/sites/media");

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual("https://contoso.sharepoint.com/.default", fake.CapturedParams["scope"]);
        }

        [TestMethod]
        public void RequestOAuthToken_DoesNotIncludeResourceParam()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings();

            SharePointHelper.RequestOAuthToken(settings, fake);

            // Entra ID v2.0 uses "scope", not "resource"
            Assert.IsNull(fake.CapturedParams["resource"]);
        }

        [TestMethod]
        public void RequestOAuthToken_WhenNotCertificateAuth_SendsClientSecret()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings(clientSecret: "my-secret");
            settings.UseCertificateAuth = false;

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.AreEqual("my-secret", fake.CapturedParams["client_secret"]);
            Assert.IsNull(fake.CapturedParams["client_assertion"]);
            Assert.IsNull(fake.CapturedParams["client_assertion_type"]);
        }

        [TestMethod]
        public void RequestOAuthToken_WhenNotCertificateAuth_DoesNotSendAssertion()
        {
            var fake = new FakeTokenClient();
            var settings = CreateModernAuthSettings();
            settings.UseCertificateAuth = false;

            SharePointHelper.RequestOAuthToken(settings, fake);

            Assert.IsNull(fake.CapturedParams["client_assertion"]);
            Assert.IsNull(fake.CapturedParams["client_assertion_type"]);
        }
    }
}
