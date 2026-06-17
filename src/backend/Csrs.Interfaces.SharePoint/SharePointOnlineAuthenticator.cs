using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace Csrs.Interfaces
{
    internal class SharePointOnlineAuthenticator : ISharePointOnlineAuthenticator
    {
        private const string CacheKey = "SharePointOnlineAccessToken";
        // Subtract 60 seconds from the token lifetime as a buffer to avoid using an expiring token
        private const int ExpiryBufferSeconds = 60;

        private readonly SharePointOnlineConfiguration _configuration;
        private readonly IMemoryCache _cache;
        private readonly ILogger<SharePointOnlineAuthenticator> _logger;

        public SharePointOnlineAuthenticator(
            SharePointOnlineConfiguration configuration,
            IMemoryCache cache,
            ILogger<SharePointOnlineAuthenticator> logger)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _cache = cache ?? throw new ArgumentNullException(nameof(cache));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public async Task<string> GetAccessTokenAsync()
        {
            if (_cache.TryGetValue(CacheKey, out string cachedToken))
            {
                _logger.LogTrace("Returning cached SharePoint Online access token");
                return cachedToken;
            }

            _logger.LogDebug("Requesting new SharePoint Online access token for tenant {TenantId}", _configuration.TenantId);

            string tokenUrl = $"https://login.microsoftonline.com/{_configuration.TenantId}/oauth2/v2.0/token";
            string scope = $"https://{_configuration.Resource.Host}/.default";

            using var client = new HttpClient();

            var requestBody = new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = _configuration.ClientId,
                ["client_secret"] = _configuration.ClientSecret,
                ["scope"] = scope,
            };

            using var content = new FormUrlEncodedContent(requestBody);
            var response = await client.PostAsync(tokenUrl, content);

            string responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogError("Failed to acquire SharePoint Online access token. StatusCode={StatusCode}", response.StatusCode);
                throw new SamlAuthenticationException($"SharePoint Online token request failed with status {response.StatusCode}");
            }

            var tokenResponse = JsonConvert.DeserializeObject<OAuthTokenResponse>(responseJson);
            if (tokenResponse?.AccessToken == null)
            {
                _logger.LogError("SharePoint Online token response did not contain an access token");
                throw new SamlAuthenticationException("SharePoint Online token response did not contain an access token");
            }

            int expiresIn = tokenResponse.ExpiresIn > ExpiryBufferSeconds
                ? tokenResponse.ExpiresIn - ExpiryBufferSeconds
                : tokenResponse.ExpiresIn;

            var expiry = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
            _cache.Set(CacheKey, tokenResponse.AccessToken, expiry);

            _logger.LogInformation("SharePoint Online access token acquired and cached until {Expiry}", expiry);

            return tokenResponse.AccessToken;
        }
    }

    internal class OAuthTokenResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }

        [JsonProperty("expires_in")]
        public int ExpiresIn { get; set; }

        [JsonProperty("token_type")]
        public string TokenType { get; set; }
    }
}
