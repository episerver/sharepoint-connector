using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    /// <summary>
    /// Builds a signed JWT client assertion for certificate-based OAuth 2.0 client credentials flow
    /// with Microsoft Entra ID.
    /// </summary>
    internal static class CertificateAssertionHelper
    {
        /// <summary>
        /// Loads an X.509 certificate from the Windows Certificate Store by thumbprint.
        /// Searches LocalMachine\My first, then CurrentUser\My.
        /// </summary>
        internal static X509Certificate2 LoadCertificate(string thumbprint)
        {
            if (string.IsNullOrEmpty(thumbprint))
            {
                throw new ArgumentException("Certificate thumbprint must not be empty.");
            }

            // Normalize thumbprint: remove spaces, dashes, and convert to uppercase
            thumbprint = thumbprint.Replace(" ", "").Replace("-", "").ToUpperInvariant();

            var cert = FindCertificate(thumbprint, StoreLocation.LocalMachine);
            if (cert == null)
            {
                cert = FindCertificate(thumbprint, StoreLocation.CurrentUser);
            }

            if (cert == null)
            {
                throw new InvalidOperationException(
                    string.Format("Certificate with thumbprint '{0}' was not found in the LocalMachine or CurrentUser certificate stores.", thumbprint));
            }

            if (!cert.HasPrivateKey)
            {
                throw new InvalidOperationException(
                    string.Format("Certificate with thumbprint '{0}' does not have a private key. The certificate must be installed with its private key.", thumbprint));
            }

            return cert;
        }

        /// <summary>
        /// Creates a signed JWT client assertion for the Entra ID token endpoint.
        /// </summary>
        internal static string CreateClientAssertion(string clientId, string tenantId, X509Certificate2 certificate)
        {
            var now = DateTime.UtcNow;
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            long nbf = (long)(now - epoch).TotalSeconds;
            long exp = nbf + 600; // 10 minutes

            // Header
            var header = new Dictionary<string, string>
            {
                { "alg", "RS256" },
                { "typ", "JWT" },
                { "x5t", Base64UrlEncode(certificate.GetCertHash()) }
            };

            // Payload
            var payload = new Dictionary<string, object>
            {
                { "aud", string.Format("https://login.microsoftonline.com/{0}/oauth2/v2.0/token", tenantId) },
                { "iss", clientId },
                { "sub", clientId },
                { "jti", Guid.NewGuid().ToString() },
                { "nbf", nbf },
                { "exp", exp }
            };

            string headerJson = JsonConvert.SerializeObject(header);
            string payloadJson = JsonConvert.SerializeObject(payload);

            string headerB64 = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
            string payloadB64 = Base64UrlEncode(Encoding.UTF8.GetBytes(payloadJson));

            string dataToSign = string.Format("{0}.{1}", headerB64, payloadB64);

            byte[] signature;
            using (var rsa = (RSACryptoServiceProvider)certificate.PrivateKey)
            {
                signature = rsa.SignData(Encoding.UTF8.GetBytes(dataToSign), CryptoConfig.MapNameToOID("SHA256"));
            }

            return string.Format("{0}.{1}.{2}", headerB64, payloadB64, Base64UrlEncode(signature));
        }

        private static X509Certificate2 FindCertificate(string thumbprint, StoreLocation location)
        {
            var store = new X509Store(StoreName.My, location);
            try
            {
                store.Open(OpenFlags.ReadOnly);
                var certs = store.Certificates.Find(X509FindType.FindByThumbprint, thumbprint, false);
                if (certs.Count > 0)
                {
                    return certs[0];
                }
                return null;
            }
            finally
            {
                store.Close();
            }
        }

        private static string Base64UrlEncode(byte[] data)
        {
            return Convert.ToBase64String(data)
                .Replace('+', '-')
                .Replace('/', '_')
                .TrimEnd('=');
        }
    }
}
