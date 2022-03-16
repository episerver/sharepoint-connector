using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.Framework.Cache;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// The helper class for SharePoint
    /// </summary>
    public class SharePointHelper : ISharePointHelper
    {
        IServiceLocator _serviceLocator;

        public SharePointHelper()
        {
            _serviceLocator = ServiceLocator.Current;
        }

        internal SharePointHelper(IServiceLocator locator)
        {
            _serviceLocator = locator;
        }
        private static readonly ILogger _logger = LogManager.GetLogger();

        private const int DefaultCookieContainerCacheSeconds = 1 * 60 * 60; // 1 hour
        private const int CookieContainerCacheOffsetSeconds = 5 * 60; // 5 minutes

        /// <summary>
        /// The cache key for form digest value
        /// </summary>
        public static string FormDigestValueCacheKey
        {
            get
            {
                return string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, "FormDigestValue");
            }
        }

        /// <summary>
        /// The cache key for Cookie container
        /// </summary>
        public static string CookieContainerCacheKey
        {
            get
            {
                return string.Format(Consts.SHAREPOINT_CACHE_KEY_FORMAT, "CookieContainer");
            }
        }

        /// <summary>
        /// Gets SharePoint's list and view Id from the given mapping identification
        /// </summary>
        /// <param name="MappingKey">The given mapping identification</param>
        public string[] GetListAndViewIds(string mappingId)
        {
            if (string.IsNullOrEmpty(mappingId) || (!string.IsNullOrEmpty(mappingId) & !mappingId.Contains('~')))
            {
                _logger.Error("Invalid mapping Id value - mappingId: {0}", string.IsNullOrEmpty(mappingId) ? string.Empty : mappingId);
                return null;
            }

            return mappingId.Split('~');
        }

        /// <summary>
        /// Get list of id pair of list and view from given list and the Settings's Id list
        /// </summary>
        /// <param name="lists">List of <see cref="SPList"/> for filtering.</param>
        public IEnumerable<KeyValuePair<string, string>> GetSelectedListViewIds(IEnumerable<SPList> lists)
        {
            var listIds = SharePointSettings.Current.ListIds.Split(new string[] { ";" }, StringSplitOptions.None);
            var idListViewPairs = GetListViewPairs(lists);
            return idListViewPairs.Where(li => listIds.Contains(li.Key));
        }

        /// <summary>
        /// Return a dictionary of ListId~ViewId and ListTitle - ViewTitle for unhidden View
        /// </summary>
        /// <param name="lists">List of SPList</param>
        public static IEnumerable<KeyValuePair<string, string>> GetListViewPairs(IEnumerable<SPList> lists)
        {
            var idPairs = lists.Where(list => Enum.IsDefined(typeof(SupportedBaseTemplates), list.BaseTemplate)).
                SelectMany<SPList, KeyValuePair<string, string>>(spl =>
            {
                return spl.Views.Where(view => !view.Hidden).Select<SPListView, KeyValuePair<string, string>>(spv =>
                {
                    return new KeyValuePair<string, string>(String.Format("{0}~{1}", spl.Id, spv.Id),
                        String.Format("{0} - {1}", spl.Title, spv.Title));
                });
            });

            return idPairs;
        }

        /// <summary>
        /// Return a List of ListId~ListTitle and a list of corresponding ViewId~ViewTitle for unhidden Views
        /// </summary>
        /// <param name="lists">List of SPList</param>
        public static List<KeyValuePair<string, List<KeyValuePair<string, string>>>> GetListViewPairsData(IEnumerable<SPList> lists)
        {
            List<KeyValuePair<string, List<KeyValuePair<string, string>>>> reqLists = new List<KeyValuePair<string, List<KeyValuePair<string, string>>>>();
            var supportedLists = lists.Where(list => Enum.IsDefined(typeof(SupportedBaseTemplates), list.BaseTemplate));
            foreach (SPList spl in supportedLists)
            {
                var reqViews = new List<KeyValuePair<string, string>>();
                var unhiddenViews = spl.Views.Where(views => !views.Hidden);
                foreach (SPListView v in unhiddenViews)
                {
                    reqViews.Add(new KeyValuePair<string, string>(string.Format("{0}~{1}", spl.Id, v.Id), v.Title));
                }

                reqLists.Add(new KeyValuePair<string, List<KeyValuePair<string, string>>>(String.Format("{0}~{1}", spl.Id, spl.Title), reqViews));
            }

            return reqLists;
        }

        /// <summary>
        /// Gets the CookieContainer.
        /// </summary>
        /// <param name="spSettings">The SharePoint settings.</param>
        public static CookieContainer GetCookies(SharePointSettings spSettings)
        {
            var cookies = CacheManager.Get(CookieContainerCacheKey) as CookieContainer;
            if (cookies == null)
            {
                var cookieContainer = GetLatestCookies(spSettings);
                if (cookieContainer == null)
                {
                    var exception = new Exception("Cannot get the cookies, please check authenticate information.");
                    _logger.Error("Cannot get the cookies, please check authenticate information.", exception);
                    throw exception;
                }

                // Store the cookies in cache
                CacheManager.Insert(CookieContainerCacheKey, cookieContainer, new CacheEvictionPolicy(TimeSpan.FromSeconds(DefaultCookieContainerCacheSeconds), CacheTimeoutType.Absolute));
                cookies = CacheManager.Get(CookieContainerCacheKey) as CookieContainer;
            }
            return cookies;
        }

        /// <summary>
        /// Get a fresh CookieContainer from the server
        /// </summary>
        /// <param name="spSettings">The SharePoint settings.</param>
        public static CookieContainer GetLatestCookies(SharePointSettings spSettings)
        {
            var securePassword = new SecureString();
            var siteUri = new Uri(spSettings.SiteUrl);
            foreach (var c in spSettings.Password) { securePassword.AppendChar(c); }
            var credentials = new SharePointOnlineCredentials(spSettings.Username, securePassword);
            var authCookie = credentials.GetAuthenticationCookie(siteUri);
            var cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(siteUri, authCookie);
            return cookieContainer;
        }

        /// <summary>
        /// Get OAuth response from Sharepoint.
        /// </summary>
        /// <param name="spSettings">The SharePoint settings.</param>
        /// <returns></returns>
        public static SharePointOAuthTokenResult GetOAuthResponse(SharePointSettings spSettings)
        {
            using (WebClient client = new WebClient())
            {
                var domain = new Uri(spSettings.TargetHost).Host;
                var resource = string.Format("00000003-0000-0ff1-ce00-000000000000/{0}@{1}", domain, spSettings.Realm);
                var reqParams = new System.Collections.Specialized.NameValueCollection();
                reqParams.Add("grant_type", "client_credentials");
                reqParams.Add("client_id", string.Format("{0}@{1}", spSettings.ClientId, spSettings.Realm));
                reqParams.Add("client_secret", spSettings.ClientSecret);
                reqParams.Add("resource", resource);

                var requestUrl = string.Format("https://accounts.accesscontrol.windows.net/{0}/tokens/OAuth/2", spSettings.Realm);
                client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                byte[] responsebytes = client.UploadValues(requestUrl, "POST", reqParams);
                string responsebody = Encoding.UTF8.GetString(responsebytes);
                return JsonConvert.DeserializeObject<SharePointOAuthTokenResult>(responsebody);
            }
        }

        /// <summary>
        /// Create NTLM CredentialCache from Settings
        /// </summary>
        /// <param name="spSettings">The settings includes site and authentication info.</param>
        public static CredentialCache GetCredentialCache(SharePointSettings spSettings)
        {
            var credentialCache = new CredentialCache();
            credentialCache.Add(new Uri(spSettings.SiteUrl), "NTLM", new NetworkCredential(spSettings.Username, spSettings.Password));

            return credentialCache;
        }

        /// <summary>
        /// Determines whether the current connected SharePoint site is the online site or not based on the username.
        /// </summary>
        /// <param name="username">The username.</param>
        /// <returns></returns>
        public static bool IsConnectToSharePointOnline(string username)
        {
            // SharePoint Online's username always end with @xxx.onmicrosoft.com.
            var regex = new Regex(@"@[^\s\.]+\.onmicrosoft\.com$", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

            return regex.IsMatch(username);
        }

        public IEnumerable<SPListItem> GetPagedItems(SPList list, SPListView view, SPFolder folder, out string nextUri)
        {
            var service = _serviceLocator.GetInstance<IListsViewsService>();
            return service.GetPagedItems(list, view, folder, out nextUri);
        }
        public IEnumerable<SPListItem> GetNextPagedItems(SPList list, string Uri, out string nextUri)
        {
            var service = _serviceLocator.GetInstance<IListsViewsService>();
            return service.GetNextPagedItems(list, Uri, out nextUri);
        }
    }
}