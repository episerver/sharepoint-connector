using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.Framework.Serialization;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Text;

namespace EPiServer.ConnectForSharePoint.Rest
{
    public class SPRestClient
    {
        private static readonly ILogger _logger = LogManager.GetLogger();

        private IObjectSerializer _serializer = ServiceLocator.Current.GetInstance<IObjectSerializerFactory>().GetSerializer(KnownContentTypes.Json);
        private SharePointSettings _spSettings;
        private Func<SharePointSettings, CookieContainer> _getCookieContainer;
        private SharePointOAuthTokenResult _oAuthTokenResult;

        public SPRestClient()
        {
            _spSettings = SharePointSettings.Current;

            if (_spSettings.IsModernAuth)
            {
                _oAuthTokenResult = SharePointHelper.GetOAuthResponse(_spSettings);
            }
            else
            {
                _getCookieContainer = (Settings) =>
                {
                    return SharePointHelper.GetCookies(Settings);
                };
            }
        }

        /// <summary>
        /// For testing from the UI.
        /// </summary>
        /// <param name="settings"></param>
        public SPRestClient(SharePointSettings settings)
        {
            _spSettings = settings;

            if (_spSettings.IsModernAuth)
            {
                _oAuthTokenResult = SharePointHelper.GetOAuthResponse(_spSettings);
            }
            else
            {
                _getCookieContainer = (Settings) =>
                {
                    return SharePointHelper.GetCookies(Settings);
                };
            }
        }

        /// <summary>
        /// Create and then send a <see cref="WebRequest"/> to specified REST endpoint.
        /// </summary>
        /// <typeparam name="requestInfo">Contain the request endpoint and method.</typeparam>
        /// <typeparam name="requestBody">Contain the request body.</typeparam>
        /// <typeparam name="headers">Contain the addition header values.</typeparam>
        public string SendRequest(IRestRequest requestInfo, string requestBody, NameValueCollection headers)
        {
            var responseString = string.Empty;
            var webRequest = CreateRequest(requestInfo, requestBody, headers);

            try
            {
                return WebHelper.GetResponseString(webRequest);
            }
            catch (Exception ex)
            {
                _logger.Debug("Sending '{0}' request, exception: {1}", webRequest.RequestUri.AbsoluteUri, ex);
                throw;
            }
        }

        public Stream GetResponseStream(IRestRequest requestInfo, string requestBody, NameValueCollection headers)
        {
            var webRequest = CreateRequest(requestInfo, requestBody, headers);

            try
            {
                return WebHelper.GetResponseStream(webRequest);
            }
            catch (Exception ex)
            {
                _logger.Debug("Sending '{0}' request, exception: {1}", webRequest.RequestUri.AbsoluteUri, ex);
                throw;
            }
        }

        public TResult Deserialize<TResult>(string responseString)
        {
            if (!string.IsNullOrEmpty(responseString))
            {
                return _serializer.Deserialize<TResult>(new StringReader(responseString));
            }

            return default(TResult);
        }

        private WebRequest CreateRequest(IRestRequest requestInfo, string requestBody, NameValueCollection headers)
        {
            var requestUrl = WebHelper.CombineUrl(_spSettings.SiteUrl, requestInfo.EncodedUrl);
            string acceptType = "application/json;odata=verbose";
            if (requestInfo.EncodedUrl.StartsWith(_spSettings.SiteUrl, StringComparison.OrdinalIgnoreCase))
            {
                acceptType = "application/atom+xml";
            }

            var webRequest = (HttpWebRequest)WebHelper.CreateRequest(requestUrl, requestInfo.RequestMethod, headers, acceptType);

            if (_spSettings.IsModernAuth)
            {
                webRequest.Headers.Add("Authorization", "Bearer " + _oAuthTokenResult.AccessToken);
            }
            else
            {
                // NOTE: the CookieContainer MUST be set before GetRequestStream() call
                // so that any POST request can work successfully
                if (_spSettings.IsSharePointOnline || SharePointHelper.IsConnectToSharePointOnline(_spSettings.Username))
                {
                    webRequest.CookieContainer = _getCookieContainer(_spSettings);
                }
                else
                {
                    webRequest.Credentials = SharePointHelper.GetCredentialCache(_spSettings);
                }
            }

            AddRequestBody(webRequest, requestBody);

            return webRequest;
        }

        /// <summary>
        /// Add body to the request. This must be call after assigning CookieContainer and Credentials
        /// </summary>
        /// <param name="request"></param>
        /// <param name="requestBody"></param>
        private void AddRequestBody(WebRequest request, string requestBody)
        {
            if (!string.IsNullOrWhiteSpace(requestBody))
            {
                // Note: support unicode chars in the requestBody
                byte[] requestBytes = Encoding.UTF8.GetBytes(requestBody);
                request.ContentLength = requestBytes.Length;

                using (var requestStream = request.GetRequestStream())
                {
                    requestStream.Write(requestBytes, 0, requestBytes.Length);
                    requestStream.Flush();
                }
            }
            else if (request.Method.EndsWith("POST", System.StringComparison.OrdinalIgnoreCase))
            {
                request.ContentLength = 0;
            }
        }
    }
}