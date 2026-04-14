using System;
using System.Collections.Specialized;
using System.IO;
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
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
                    byte[] responseBytes = client.UploadValues(requestUrl, "POST", parameters);
                    return Encoding.UTF8.GetString(responseBytes);
                }
            }
            catch (WebException ex)
            {
                if (ex.Response == null)
                {
                    throw;
                }

                string responseBody;
                using (var reader = new StreamReader(ex.Response.GetResponseStream()))
                {
                    responseBody = reader.ReadToEnd();
                }

                throw new WebException(
                    string.Format("Token request failed: {0}", responseBody),
                    ex,
                    ex.Status,
                    ex.Response);
            }
        }
    }
}
