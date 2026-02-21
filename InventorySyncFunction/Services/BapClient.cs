using System;
using System.Net.Http;
using System.Threading.Tasks;
using Azure.Identity;
using Azure.Core;
using Newtonsoft.Json.Linq;
using InventorySyncFunction.Interfaces;
using Microsoft.Extensions.Logging;

namespace InventorySyncFunction.Services
{
    public class BapClient : IBapClient
    {
        private readonly ClientSecretCredential _cred;
        private readonly ILogger _logger;
        private readonly Microsoft.PowerPlatform.Management.ServiceClient _mgmt;

        public BapClient(string tenantId, string clientId, string clientSecret, ILogger logger)
        {
            _cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
            _logger = logger;

            // Initialize Management SDK (Kiota)
            Microsoft.Kiota.Abstractions.Serialization.ParseNodeFactoryRegistry.DefaultInstance.ContentTypeAssociatedFactories.TryAdd("application/json", new Microsoft.Kiota.Serialization.Json.JsonParseNodeFactory());
            Microsoft.Kiota.Abstractions.Serialization.SerializationWriterFactoryRegistry.DefaultInstance.ContentTypeAssociatedFactories.TryAdd("application/json", new Microsoft.Kiota.Serialization.Json.JsonSerializationWriterFactory());

            var authProvider = new UnifiedAuthProvider(_cred);
            var adapter = new Microsoft.Kiota.Http.HttpClientLibrary.HttpClientRequestAdapter(authProvider) { BaseUrl = "https://api.powerplatform.com" };
            _mgmt = new Microsoft.PowerPlatform.Management.ServiceClient(adapter);
        }

        public Microsoft.PowerPlatform.Management.ServiceClient Management => _mgmt;

        public async Task<JArray?> FetchList(string scope, string url)
        {
            try
            {
                var token = await _cred.GetTokenAsync(new TokenRequestContext(new[] { $"https://{scope}/.default" }));
                using var http = new HttpClient();
                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
                var resp = await http.GetAsync(url);
                if (!resp.IsSuccessStatusCode) return null;
                var json = await resp.Content.ReadAsStringAsync();
                return (JArray?)JObject.Parse(json)["value"];
            }
            catch (Exception ex)
            {
                _logger.LogError($"BapClient Error ({url}): {ex.Message} | Inner: {ex.InnerException?.Message}");
                return null;
            }
        }

        public async Task<JArray?> FetchDataverse(string instanceUrl, string query)
        {
            try
            {
                var host = new Uri(instanceUrl).Host;
                var token = await _cred.GetTokenAsync(new TokenRequestContext(new[] { $"https://{host}/.default" }));
                using var http = new HttpClient();
                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
                var resp = await http.GetAsync($"{instanceUrl.TrimEnd('/')}/api/data/v9.2/{query}");
                if (!resp.IsSuccessStatusCode) return null;
                var json = await resp.Content.ReadAsStringAsync();
                return (JArray?)JObject.Parse(json)["value"];
            }
            catch (Exception ex)
            {
                _logger.LogError($"Dataverse Sync Error ({instanceUrl}): {ex.Message}");
                return null;
            }
        }

        public async Task<bool> CheckAccess(string instanceUrl)
        {
            try
            {
                var host = new Uri(instanceUrl).Host;
                var token = await _cred.GetTokenAsync(new TokenRequestContext(new[] { $"https://{host}/.default" }));
                using var http = new HttpClient();
                http.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token.Token);
                var resp = await http.GetAsync($"{instanceUrl}/api/data/v9.2/WhoAmI");
                return resp.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }
    }

    public class UnifiedAuthProvider : Microsoft.Kiota.Abstractions.Authentication.IAuthenticationProvider
    {
        private readonly ClientSecretCredential _cred;
        public UnifiedAuthProvider(ClientSecretCredential cred) => _cred = cred;
        public async Task AuthenticateRequestAsync(Microsoft.Kiota.Abstractions.RequestInformation request, Dictionary<string, object>? context = null, System.Threading.CancellationToken token = default)
        {
            var result = await _cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.powerplatform.com/.default" }), token);
            request.Headers.Add("Authorization", $"Bearer {result.Token}");
        }
    }
}
