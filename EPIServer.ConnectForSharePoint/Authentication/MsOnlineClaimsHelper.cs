using EPiServer.Logging;
using System;
using System.IdentityModel.Protocols.WSTrust;
using System.IdentityModel.Tokens;
using System.Linq;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Channels;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace EPiServer.ConnectForSharePoint.Authentication
{
    public class MsOnlineClaimsHelper
    {
        private static readonly ILogger _logger = LogManager.GetLogger();

        #region Properties

        readonly string _username;
        readonly string _password;
        readonly bool _useRtfa;
        readonly Uri _host;

        CookieContainer _cachedCookieContainer = null;
        DateTime _expires = DateTime.MinValue;

        #endregion

        #region Constructors

        public MsOnlineClaimsHelper(string host, string username, string password)
            : this(new Uri(host), username, password) { }

        public MsOnlineClaimsHelper(Uri host, string username, string password)
        {
            _host = host;
            _username = username;
            _password = password;
            _useRtfa = true;
        }

        public MsOnlineClaimsHelper(Uri host, string username, string password, bool useRtfa)
        {
            _host = host;
            _username = username;
            _password = password;
            _useRtfa = useRtfa;
        }

        #endregion

        #region Constants

        public const string office365STS = "https://login.microsoftonline.com/extSTS.srf";
        public const string office365Login = "https://login.microsoftonline.com/login.srf";
        public const string office365Metadata = "https://nexus.microsoftonline-p.com/federationmetadata/2007-06/federationmetadata.xml";
        public const string wsse = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd";
        public const string wsu = "http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd";
        
        private const string userAgent = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)";

        #endregion

        class MsoCookies
        {
            public string FedAuth { get; set; }
            public string rtFa { get; set; }
            public DateTime Expires { get; set; }
            public Uri Host { get; set; }
        }

        // Creates or loads cached cookie container
        CookieContainer GetCookieContainer()
        {
            if (_cachedCookieContainer == null || DateTime.Now > _expires)
            {
                // Get the SAML tokens from SPO STS (via MSO STS) using fed auth passive approach
                var cookies = GetSamlToken();

                if (cookies != null && !string.IsNullOrEmpty(cookies.FedAuth))
                {
                    // Create cookie collection with the SAML token
                    _expires = cookies.Expires;
                    var cc = new CookieContainer();

                    // Set the FedAuth cookie
                    var samlAuth = new Cookie("FedAuth", cookies.FedAuth)
                    {
                        Expires = cookies.Expires,
                        Path = "/",
                        Secure = cookies.Host.Scheme == "https",
                        HttpOnly = true,
                        Domain = cookies.Host.Host
                    };
                    cc.Add(samlAuth);

                    if (_useRtfa)
                    {
                        // Set the rtFA (sign-out) cookie, added march 2011
                        var rtFa = new Cookie("rtFA", cookies.rtFa)
                        {
                            Expires = cookies.Expires,
                            Path = "/",
                            Secure = cookies.Host.Scheme == "https",
                            HttpOnly = true,
                            Domain = cookies.Host.Host
                        };
                        cc.Add(rtFa);
                    }
                    _cachedCookieContainer = cc;

                    return cc;
                }

                return null;
            }

            return _cachedCookieContainer;
        }

        public CookieContainer CookieContainer
        {
            get
            {
                if (_cachedCookieContainer == null || DateTime.Now > _expires)
                {
                    return GetCookieContainer();
                }

                return _cachedCookieContainer;
            }
        }

        private MsoCookies GetSamlToken()
        {
            var ret = new MsoCookies();

            try
            {
                var sharepointSite = new
                {
                    Wctx = office365Login,
                    Wreply = _host.GetLeftPart(UriPartial.Authority) + "/_forms/default.aspx?wa=wsignin1.0"
                };

                //get token from STS
                var stsResponse = GetResponse(office365STS, sharepointSite.Wreply);

                // parse the token response
                XDocument doc = XDocument.Parse(stsResponse);

                // get the security token
                var crypt = from result in doc.Descendants()
                            where result.Name == XName.Get("BinarySecurityToken", wsse)
                            select result;

                // get the token expiration
                var expires = from result in doc.Descendants()
                              where result.Name == XName.Get("Expires", wsu)
                              select result;
                ret.Expires = Convert.ToDateTime(expires.First().Value);

                var request = CreateRequest(sharepointSite.Wreply);
                var data = Encoding.UTF8.GetBytes(crypt.FirstOrDefault().Value);

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                    stream.Close();

                    using (var webResponse = request.GetResponse() as HttpWebResponse)
                    {
                        // Handle redirect, added may 2011 for P-subscriptions
                        if (webResponse.StatusCode == HttpStatusCode.MovedPermanently)
                        {
                            var request2 = CreateRequest(webResponse.Headers["Location"]);

                            using (var stream2 = request2.GetRequestStream())
                            {
                                stream2.Write(data, 0, data.Length);
                                stream2.Close();

                                using (var webResponse2 = request2.GetResponse() as HttpWebResponse)
                                {
                                    ret.FedAuth = webResponse2.Cookies["FedAuth"].Value;
                                    ret.rtFa = webResponse2.Cookies["rtFa"].Value;
                                    ret.Host = request2.RequestUri;
                                }
                            }
                        }
                        else
                        {
                            ret.FedAuth = webResponse.Cookies["FedAuth"].Value;
                            ret.rtFa = webResponse.Cookies["rtFa"].Value;
                            ret.Host = request.RequestUri;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error("Error occurs when getting Saml token", ex);
                return null;
            }

            return ret;
        }

        static HttpWebRequest CreateRequest(string url)
        {
            var request = HttpWebRequest.Create(url) as HttpWebRequest;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.CookieContainer = new CookieContainer();
            request.AllowAutoRedirect = false; // Do NOT automatically redirect
            request.UserAgent = userAgent;

            return request;
        }

        private string GetResponse(string stsUrl, string realm)
        {
            var rst = new RequestSecurityToken
            {
                RequestType = RequestTypes.Issue,   // changed from WSTrustFeb2005Constants.RequestTypes.Issue,
                AppliesTo = new EndpointReference(realm), //changed from new EndpointAddress(realm),
                KeyType = KeyTypes.Bearer,  // changed from WSTrustFeb2005Constants.KeyTypes.Bearer,
                TokenType = SecurityTokenTypes.Saml //changed from SecurityTokenTypes.Saml11TokenProfile11
            };

            WSTrustFeb2005RequestSerializer trustSerializer = new WSTrustFeb2005RequestSerializer();

            var binding = new WSHttpBinding();
            binding.Security.Mode = SecurityMode.TransportWithMessageCredential;
            binding.Security.Message.ClientCredentialType = MessageCredentialType.UserName;
            binding.Security.Message.EstablishSecurityContext = false;
            binding.Security.Message.NegotiateServiceCredential = false;
            binding.Security.Transport.ClientCredentialType = HttpClientCredentialType.None;

            var address = new EndpointAddress(stsUrl);

            using (WSTrustFeb2005ContractClient trustClient = new WSTrustFeb2005ContractClient(binding, address))
            {
                trustClient.ClientCredentials.UserName.UserName = _username;
                trustClient.ClientCredentials.UserName.Password = _password;

                Message response = trustClient.EndIssue(
                    trustClient.BeginIssue(
                        Message.CreateMessage(
                            MessageVersion.Default,
                            // Could not find value equals to WSTrustFeb2005Constants.Actions.Issue, workaround using this hardcodely, this from error 
                            // {"Outgoing request message for operation 'Issue' specified Action='http://schemas.microsoft.com/idfx/requesttype/issue', 
                            // but contract for that operation specifies Action='http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue'.  
                            // The Action specified in the Message must match the Action in the contract, or the operation contract must specify Action='*'."}
                            "http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue",
                            new RequestBodyWriter(trustSerializer, rst)
                        ),
                        null,
                        null));

                trustClient.Close();

                using (XmlDictionaryReader reader = response.GetReaderAtBodyContents())
                {
                    return reader.ReadOuterXml();
                }
            }
        }
    }
}