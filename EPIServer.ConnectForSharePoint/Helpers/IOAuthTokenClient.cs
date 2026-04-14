using System.Collections.Specialized;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Abstraction for HTTP token requests, allowing the OAuth request construction
    /// logic in <see cref="SharePointHelper"/> to be unit-tested without real network calls.
    /// </summary>
    public interface IOAuthTokenClient
    {
        /// <summary>
        /// Posts the given parameters to the token endpoint and returns the raw JSON response body.
        /// </summary>
        string RequestToken(string requestUrl, NameValueCollection parameters);
    }
}
