using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Text;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Helper class for web request, response
    /// </summary>
    public static class WebHelper
    {
        private static bool[] _validCharactersLookup;

        /// <summary>
        /// Create a WebRequest has "application/json;odata=verbose" ContentType from a SP REST endpoint
        /// </summary>
        /// <param name="requestUrl">The SP REST endpoint.</param>
        /// <param name="requestType">Request method type.</param>
        /// <param name="headerNameValues">Addition headers for the request.</param>
        /// <returns>WebRequest object</returns>
        public static WebRequest CreateRequest(string requestUrl, string requestType, NameValueCollection headerNameValues)
        {
            var request = (HttpWebRequest)HttpWebRequest.Create(requestUrl);
            request.Method = string.IsNullOrEmpty(requestType) ? "GET" : requestType;
            request.ContentType = "application/json;odata=verbose";
            request.Accept = "application/json;odata=verbose";
            request.UserAgent = "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)";

            if (headerNameValues != null && headerNameValues.Count > 0)
            {
                request.Headers.Add(headerNameValues);
            }

            return request;
        }

        /// <summary>
        /// Return response string from WebRequest
        /// </summary>
        /// <param name="webRequest">The request to get response.</param>
        public static string GetResponseString(WebRequest webRequest)
        {
            return GetResponseString(webRequest.GetResponse());
        }

        /// <summary>
        /// Return stream from WebRequest
        /// </summary>
        /// <param name="webRequest">The request to get response.</param>
        public static Stream GetResponseStream(WebRequest webRequest)
        {
            var webResponse = webRequest.GetResponse();
            return webResponse.GetResponseStream();
        }

        /// <summary>
        /// Strips special characters in the given url string
        /// </summary>
        /// <param name="url">The given url string</param>
        /// <returns>The stripped special characters url string</returns>
        public static string StripUrlSpecialCharacters(string url)
        {
            if (string.IsNullOrEmpty(url))
            {
                return string.Empty;
            }

            var lookup = GetValidCharacters();
            char[] buffer = new char[url.Length];
            int index = 0;

            foreach (char c in url)
            {
                if (lookup[c])
                {
                    buffer[index] = c;
                    index++;
                }
            }

            return new string(buffer, 0, index);
        }

        /// <summary>
        /// Combines the URL paths to make sure we do not have duplicated slashes or base Url
        /// </summary>
        /// <param name="baseUrl">The base URL path.</param>
        /// <param name="requestPath">The request URL path.</param>
        /// <returns></returns>
        public static string CombineUrl(string baseUrl, string requestPath)
        {
            return requestPath.StartsWith(baseUrl, StringComparison.OrdinalIgnoreCase) ?
                requestPath : UriSupport.Combine(baseUrl, requestPath);
        }

        /// <summary>
        /// Gets web response string
        /// </summary>
        /// <param name="webResponse">The web response object</param>
        /// <returns>Web response string</returns>
        private static string GetResponseString(WebResponse webResponse)
        {
            using (var responseStream = webResponse.GetResponseStream())
            {
                using (var streamReader = new StreamReader(responseStream, Encoding.UTF8))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Gets valid character array
        /// </summary>
        /// <returns></returns>
        private static bool[] GetValidCharacters()
        {
            if (_validCharactersLookup != null)
            {
                return _validCharactersLookup;
            }

            var validCharactersLookup = new bool[65535];

            for (char c = '0'; c <= '9'; c++)
            {
                validCharactersLookup[c] = true;
            }

            for (char c = 'A'; c <= 'Z'; c++)
            {
                validCharactersLookup[c] = true;
            }

            for (char c = 'a'; c <= 'z'; c++)
            {
                validCharactersLookup[c] = true;
            }

            validCharactersLookup['.'] = true;
            validCharactersLookup['_'] = true;

            return _validCharactersLookup = validCharactersLookup;
        }
    }
}