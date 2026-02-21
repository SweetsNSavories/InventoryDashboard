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
        public async Task Run([TimerTrigger("0 0 12 * * *")] TimerInfo myTimer)
        {
            _logger.LogInformation($"🚀 Trigger Fired at: {DateTime.Now}");

            var clientId = Environment.GetEnvironmentVariable("PowerPlatform:ClientId");
            var clientSecret = Environment.GetEnvironmentVariable("PowerPlatform:ClientSecret");
            var tenantId = Environment.GetEnvironmentVariable("PowerPlatform:TenantId");
            var dvUrl = Environment.GetEnvironmentVariable("PowerPlatform:DataverseUrl");

            var bap = new BapClient(tenantId, clientId, clientSecret, _logger);
            var dv = new DataverseRepository($"AuthType=ClientSecret;Url={dvUrl};ClientId={clientId};ClientSecret={clientSecret};", _logger);
            
            var orchestrator = new SyncOrchestrator(bap, dv, _logger);
            
            await orchestrator.ExecuteSync();
            
            _logger.LogInformation($"🏁 Trigger Complete at: {DateTime.Now}");
        }
    }
}
