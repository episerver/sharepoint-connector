using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest;
using EPiServer.ConnectForSharePoint.Rest.Models;
using EPiServer.ConnectForSharePoint.Rest.Requests;
using EPiServer.Framework.Cache;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;

namespace EPiServer.ConnectForSharePoint.SPServices.REST
{
    /// <summary>
    /// Base class for all SharePoint's REST services
    /// </summary>
    public abstract class RESTServiceBase
    {
        private static readonly ILogger _logger = LogManager.GetLogger();

        private IEndpointContextService<SPRestClient> _endpointContextService;
        private const int FormDigestTimeoutSecondsOffset = 5 * 60;

        protected SharePointSettings SPSettings
        {
            get
            {
                return SharePointSettings.Current;
            }
        }

        /// <summary>
        /// SharePoint's endpoint context service
        /// </summary>
        public IEndpointContextService<SPRestClient> EndpointContextService
        {
            get
            {
                if (_endpointContextService == null)
                {
                    try
                    {
                        ServiceLocator.Current.TryGetExistingInstance<IEndpointContextService<SPRestClient>>(out _endpointContextService);
                    }
                    catch { }

                    if (_endpointContextService == null)
                    {
                        _endpointContextService = new EndpointContextService();

                        ServiceLocator.Current.Buildup(_endpointContextService);
                    }
                }

                return _endpointContextService;
            }
            set
            {
                _endpointContextService = value;
            }
        }

        public virtual string SendRequest(IRestRequest requestInfo, string requestBody, NameValueCollection headers)
        {
            return SendRequest(requestInfo, requestBody, headers, false);
        }

        public virtual string SendRequestWithRetry(IRestRequest requestInfo, string requestBody, NameValueCollection headers, bool addDigestIfPOST = false)
        {
            int MAX_RETRYS = 3;
            int retry = 0;
            do
            {
                try
                {
                    return SendRequest(requestInfo, requestBody, headers, addDigestIfPOST);
                }
                catch( Exception e)
                {
                    retry++;
                    if (retry > MAX_RETRYS)
                    {
                        _logger.Error("SendRequestWithRetry failed after " + MAX_RETRYS + " attempts. Path=" + requestInfo.RestApiPath, e);
                        throw;
                    }
                    _logger.Warning("Will retry last request... Retry count = " + retry);
                }
            } while (true);
        }

        public virtual string SendRequest(IRestRequest requestInfo, string requestBody, NameValueCollection headers, bool addDigestIfPOST = false)
        {
            _logger.Debug("Send request, method: {0} - addDigetsIfPOST: {1}", requestInfo.RequestMethod, addDigestIfPOST);

            if ((requestInfo.RequestMethod == "POST" || requestInfo.RequestMethod == "GET") && addDigestIfPOST)
            {
                headers = headers ?? new NameValueCollection();
                headers.Add(Consts.REQUESTDIGEST_HEADER_KEY, GetDigestValue(false));
            }

            var response = string.Empty;
            try
            {
                response = EndpointContextService.EndpointContext.SendRequest(requestInfo, requestBody, headers);
            }
            catch (WebException webException)
            {
                if (webException.Status == WebExceptionStatus.ProtocolError)
                {
                    var httpRes = webException.Response as HttpWebResponse;
                    // the X-RequestDigest expired then trying a new one
                    if (httpRes != null && httpRes.StatusCode == HttpStatusCode.Unauthorized && requestInfo.RequestMethod == "POST" && addDigestIfPOST)
                    {
                        _logger.Information("Server denied the request, try to get new Digest value and send again.");
                        headers[Consts.REQUESTDIGEST_HEADER_KEY] = GetDigestValue(true);
                        response = EndpointContextService.EndpointContext.SendRequest(requestInfo, requestBody, headers);
                    }
                    // the request to maybe deleted resource, SharePoint will return BadRequest or NotFound,
                    // return empty response so that the service call will know the resource has been deleted/changed
                    else if (httpRes != null && (httpRes.StatusCode == HttpStatusCode.BadRequest || httpRes.StatusCode == HttpStatusCode.NotFound))
                    {
                        _logger.Debug("Server returns bad request status code, return empty response.", webException);
                        throw;
                    }
                    else
                    {
                        _logger.Debug(string.Format("Cannot send {0} request to SharePoint.", requestInfo.RequestMethod), webException);
                        throw;
                    }
                }
                else
                {
                    _logger.Debug(string.Format("Cannot send {0} request to SharePoint.", requestInfo.RequestMethod), webException);
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.Debug(string.Format("Cannot send {0} request to SharePoint.", requestInfo.RequestMethod), ex);
                throw;
            }

            return response;
        }

        public virtual Stream GetResponseStream(IRestRequest requestInfo, string requestBody, NameValueCollection headers)
        {
            return EndpointContextService.EndpointContext.GetResponseStream(requestInfo, requestBody, headers);
        }

        public virtual TResult Deserialize<TResult>(string responseString)
        {
            return EndpointContextService.EndpointContext.Deserialize<TResult>(responseString);
        }

        protected string GetDigestValue(bool ignoreCache)
        {
            _logger.Debug("GetDigestValue, ignoreCache={0}", ignoreCache);

            var digestValue = (string)CacheManager.Get(SharePointHelper.FormDigestValueCacheKey);

            if (ignoreCache || string.IsNullOrEmpty(digestValue))
            {
                var contextInfo = GetContextWebInformation();
                if (contextInfo != null)
                {
                    _logger.Debug("Update cache for FormDigestValue, value: {0}", contextInfo.FormDigestValue);

                    CacheManager.Insert(SharePointHelper.FormDigestValueCacheKey, contextInfo.FormDigestValue,
                        new CacheEvictionPolicy(TimeSpan.FromSeconds(contextInfo.FormDigestTimeoutSeconds - FormDigestTimeoutSecondsOffset), CacheTimeoutType.Absolute));
                    digestValue = (string)CacheManager.Get(SharePointHelper.FormDigestValueCacheKey);
                }
            }

            _logger.Debug("Return DigestValue: {0}", digestValue);
            return digestValue;
        }

        /// <summary>
        /// Replaces special characters in a REST request that used in SharePoint's REST API
        /// </summary>
        /// <remarks>
        /// For REST API:
        ///     ' (single quote) should be encode to '' (double single quote)
        /// </remarks>
        /// <param name="inputData">The given REST request string that wanted to encode</param>
        /// <returns>The encoded string</returns>
        protected string EncodeDataForRESTRequest(string inputData)
        {
            // EXAMPLE
            // =======
            // Description:
            //      DisplayName="View'sName" will encode to DisplayName="View%27sName"
            //  BEFORE:
            //      <View Name="{E98AA09C-4AFB-44E1-814A-69B4D012384F}" MobileView="TRUE" Type="HTML" DisplayName="View'sName" Url="/Asset0001/Forms/ViewsName.aspx" Level="1" BaseViewID="40" ContentTypeID="0x" ImageUrl="/_layouts/15/images/dlicon.png?rev=23" ><Query /><RowLimit Paged="TRUE">30</RowLimit><Aggregations Value="Off" /><JSLink>clienttemplates.js|SP.UI.TileView.js|SP.UI.AssetLibrary.js|callout.js</JSLink><XslLink>AssetPicker.xsl</XslLink><Toolbar Type="Standard"/></View>
            //  AFTER:
            //      <View Name="{E98AA09C-4AFB-44E1-814A-69B4D012384F}" MobileView="TRUE" Type="HTML" DisplayName="View%27sName" Url="/Asset0001/Forms/ViewsName.aspx" Level="1" BaseViewID="40" ContentTypeID="0x" ImageUrl="/_layouts/15/images/dlicon.png?rev=23" ><Query /><RowLimit Paged="TRUE">30</RowLimit><Aggregations Value="Off" /><JSLink>clienttemplates.js|SP.UI.TileView.js|SP.UI.AssetLibrary.js|callout.js</JSLink><XslLink>AssetPicker.xsl</XslLink><Toolbar Type="Standard"/></View>

            return inputData.Replace("'", "''");
        }

        /// <summary>
        /// Replaces special characters in a CAML query that used in a REST request
        /// </summary>
        /// For CAML query used in REST API:
        ///     ' (single quote) should be encode to %27
        /// <param name="inputData">The given CAML query string that wanted to encode</param>
        /// <returns>The encoded string</returns>
        protected string EncodeDataForCAMLQuery(string inputData)
        {
            // EXAMPLE
            // =======
            // Description:
            //      "file's name.jpg" will encode to "file''s name.jpg"
            //  BEFORE:
            //      /Asset0001/file's name.jpg
            //  AFTER:
            //      /Asset0001/file''s name.jpg

            return inputData.Replace("'", "%27");
        }

        private ContextWebInformation GetContextWebInformation()
        {
            var request = new RestJsonRequest(Consts.REST_CONTEXTINFO_PATH, "POST", null);
            var responseString = SendRequest(request, null, null);
            if (string.IsNullOrEmpty(responseString))
            {
                return null;
            }

            var response = Deserialize<ContextWebInformationRestResponse>(responseString);
            if ((response != null && response.d.GetContextWebInformation != null))
            {
                var contextWebInformation = response.d.GetContextWebInformation;

                return new ContextWebInformation()
                {
                    FormDigestTimeoutSeconds = contextWebInformation.FormDigestTimeoutSeconds,
                    FormDigestValue = contextWebInformation.FormDigestValue,
                    LibraryVersion = contextWebInformation.LibraryVersion,
                    SiteFullUrl = contextWebInformation.SiteFullUrl,
                    WebFullUrl = contextWebInformation.WebFullUrl,
                    SupportedSchemaVersions = contextWebInformation.SupportedSchemaVersions.results
                };
            }

            return null;
        }
    }
}