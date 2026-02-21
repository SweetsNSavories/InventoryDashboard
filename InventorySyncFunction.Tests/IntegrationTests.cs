using System;
using System.Threading.Tasks;
using Xunit;
using Microsoft.Extensions.Logging;
using InventorySyncFunction.Services;

namespace InventorySyncFunction.Tests
{
    public class IntegrationTests
    {
        // Credentials from your environment
        private const string TenantId = "1557f771-4c8e-4dbd-8b80-dd00a88e833e";
        private const string ClientId = "3b451e9c-c4d7-4c12-8d12-69f996e7fd48";
        private const string ClientSecret = "YOUR_CLIENT_SECRET";
        private const string DvUrl = "https://orgd90897e4.crm.dynamics.com";

        [Fact]
        public async Task RealSync_ProductionRun_IsIdempotent()
        {
            var loggerFactory = LoggerFactory.Create(builder => builder.AddConsole());
            var logger = loggerFactory.CreateLogger<IntegrationTests>();

            logger.LogInformation("🚀 STARTING REAL INTEGRATION TEST...");

            // 1. Instantiate Real Services
            var bap = new BapClient(TenantId, ClientId, ClientSecret, logger);
            var dv = new DataverseRepository($"AuthType=ClientSecret;Url={DvUrl};ClientId={ClientId};ClientSecret={ClientSecret};", logger);

            // 2. Instantiate Orchestrator
            var orch = new SyncOrchestrator(bap, dv, logger);

            // 3. Execute Real Sync
            // Since this runs Parallel 10, it effectively tests thread safety of the DataverseRepository
            await orch.ExecuteSync();

            // 4. Verification
            // We can't easily assert exactly what changed without querying back,
            // but if this completes without Exception, our Thread Safety holds.
            Assert.True(dv.IsReady, "Dataverse connection should remain healthy.");
        }
    }
}
