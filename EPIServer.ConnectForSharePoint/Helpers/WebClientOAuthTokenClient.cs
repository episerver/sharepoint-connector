using System.Collections.Specialized;
using System.Net;
using System.Text;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Default <see cref="IOAuthTokenClient"/> implementation that uses <see cref="WebClient"/>.
    /// </summary>
    public class WebClientOAuthTokenClient : IOAuthTokenClient
    {
        public string RequestToken(string requestUrl, NameValueCollection parameters)
        {
            using (var client = new WebClient())
            {
                client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                byte[] responseBytes = client.UploadValues(requestUrl, "POST", parameters);
                return Encoding.UTF8.GetString(responseBytes);
            }
        }
    }
}
