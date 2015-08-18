using EPiServer.ConnectForSharePoint.Authentication;
using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using EPiServer.Framework.Cache;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// The helper class for SharePoint
    /// </summary>
    public class SharePointHelper : ISharePointHelper
    {
        public SharePointHelper()
        {

        }

        internal SharePointHelper(IServicesHub serviceHub)
        {
            _spServicesHub = serviceHub;
        }

        private static readonly ILogger _logger = LogManager.GetLogger();

        private const int DefaultCookieContainerCacheSeconds = 1 * 60 * 60; // 1 hour
        private const int CookieContainerCacheOffsetSeconds = 5 * 60; // 5 minutes

        private static IServicesHub _spServicesHub;
        public static IServicesHub SPServicesHub
        {
            get
            {
                try
                {
                    ServiceLocator.Current.TryGetExistingInstance<IServicesHub>(out _spServicesHub);
                }
                catch { }

                if (_spServicesHub == null)
                {
                    _spServicesHub = new ServicesHub();

                    ServiceLocator.Current.Buildup(_spServicesHub);
                }
                return _spServicesHub;
            }
        }

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
                var firstExpireCookie = cookieContainer.GetCookies(new Uri(spSettings.SiteUrl))
                                        .Cast<Cookie>()
                                        .OrderByDescending(c => c.Expires)
                                        .FirstOrDefault();
                var cacheTimeoutSeconds = firstExpireCookie != null ?
                    (firstExpireCookie.Expires - DateTime.Now).Subtract(TimeSpan.FromSeconds(CookieContainerCacheOffsetSeconds)).TotalSeconds
                    : DefaultCookieContainerCacheSeconds;
                CacheManager.Insert(CookieContainerCacheKey, cookieContainer, new CacheEvictionPolicy(TimeSpan.FromSeconds(cacheTimeoutSeconds), CacheTimeoutType.Absolute));
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
            var claimHelper = new MsOnlineClaimsHelper(spSettings.SiteUrl, spSettings.Username, spSettings.Password);
            return claimHelper.CookieContainer;
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

        public IEnumerable<SPListItem> GetItems(SPList list, SPListView view, SPFolder folder)
        {
            return _spServicesHub.SPListsViewsService.GetItems(list, view, folder);
        }
    }
}