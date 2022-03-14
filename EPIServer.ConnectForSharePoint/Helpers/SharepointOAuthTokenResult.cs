using Newtonsoft.Json;
using System;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public class SharepointOAuthTokenResult
    {
        [JsonProperty("token_type")]
        public String TokenType { get; set; }
        
        [JsonProperty("expires_in")]
        public String ExpiresIn { get; set; }
        
        [JsonProperty("not_before")]
        public String NotBefore { get; set; }
        
        [JsonProperty("expires_on")]
        public String ExpiresOn { get; set; }
        
        [JsonProperty("resource")]
        public String Resource { get; set; }
        
        [JsonProperty("access_token")]
        public String AccessToken { get; set; }

    }
}