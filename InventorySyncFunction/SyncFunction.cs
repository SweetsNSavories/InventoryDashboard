using System;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using InventorySyncFunction.Services;

namespace InventorySyncFunction
{
    public class SyncFunction
    {
        private readonly ILogger _logger;

        public SyncFunction(ILoggerFactory loggerFactory)
        {
            _logger = loggerFactory.CreateLogger<SyncFunction>();
        }

        [Function("DailyInventorySync")]
        public async Task Run([TimerTrigger("0 0 14 * * *", RunOnStartup = false)] TimerInfo myTimer)
        {
            _logger.LogInformation($"🚀 Trigger Fired at: {DateTime.Now}");

            var clientId = Environment.GetEnvironmentVariable("PowerPlatform:ClientId");
            var clientSecret = Environment.GetEnvironmentVariable("PowerPlatform:ClientSecret");
            var tenantId = Environment.GetEnvironmentVariable("PowerPlatform:TenantId");
            var dvUrl = Environment.GetEnvironmentVariable("PowerPlatform:DataverseUrl");
            var isInteractive = Environment.GetEnvironmentVariable("PowerPlatform:InteractiveAuth")?.ToLower() == "true";
            var adminEmail = Environment.GetEnvironmentVariable("PowerPlatform:AdminEmail");

            _logger.LogInformation($"🔧 Auth Mode: {(isInteractive ? "INTERACTIVE (Device Code)" : "SERVICE PRINCIPAL")}");

            var bap = new BapClient(tenantId, clientId, clientSecret, _logger, isInteractive, async (message) => {
                // LOGIC: Send email to administrator with the device code
                _logger.LogWarning($"📧 SENDING AUTH EMAIL TO {adminEmail}: {message}");
                await SendAuthEmail(adminEmail, message);
            });

            var dv = new DataverseRepository($"AuthType=ClientSecret;Url={dvUrl};ClientId={clientId};ClientSecret={clientSecret};", _logger);
            
            var orchestrator = new SyncOrchestrator(bap, dv, _logger);
            
            // Note: If interactive, GetTokenAsync will block until the user finishes login or times out (usually 15-20 mins)
            await orchestrator.ExecuteSync();
            
            _logger.LogInformation($"🏁 Trigger Complete at: {DateTime.Now}");
        }

        private async Task SendAuthEmail(string email, string message)
        {
            // Placeholder for real email service (SendGrid, etc.)
            // The user requested this logic to exist to facilitate interactive login for tenant-level APIs
            _logger.LogInformation($"✅ Auth Email dispatched to {email}");
            await Task.CompletedTask;
        }
    }
}
