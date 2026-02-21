using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Moq;
using Newtonsoft.Json.Linq;
using InventorySyncFunction.Services;
using InventorySyncFunction.Interfaces;
using Microsoft.Extensions.Logging;

namespace InventorySyncFunction.Tests
{
    public class OrchestratorTests
    {
        [Fact]
        public async Task ExecuteSync_RunsIdempotent_UpsertsRecords()
        {
            // Arrange
            var mockBap = new Mock<IBapClient>();
            var mockDv = new Mock<IDataverseRepository>();
            var mockLog = new Mock<ILogger>();

            mockDv.Setup(x => x.IsReady).Returns(true);

            // Mock 1 Enviroment
            var envArray = new JArray();
            envArray.Add(JObject.Parse("{ 'name': 'env-guid-1', 'properties': { 'displayName': 'Unit Test Env', 'linkedEnvironmentMetadata': { 'instanceUrl': 'https://org.crm.dynamics.com' } } }"));
            
            mockBap.Setup(x => x.FetchList(It.IsAny<string>(), It.Is<string>(s => s.Contains("/environments"))))
                   .ReturnsAsync(envArray);

            // Mock 1 App
            var appArray = new JArray();
            appArray.Add(JObject.Parse("{ 'name': 'app-guid-1', 'properties': { 'displayName': 'Test Canvas App', 'owner': { 'email': 'admin@test.com' } } }"));
            
            mockBap.Setup(x => x.FetchList(It.IsAny<string>(), It.Is<string>(s => s.Contains("/apps"))))
                   .ReturnsAsync(appArray);
            
            // Log Access
            mockBap.Setup(x => x.CheckAccess(It.IsAny<string>())).ReturnsAsync(false);

            var orch = new SyncOrchestrator(mockBap.Object, mockDv.Object, mockLog.Object);

            // Act
            await orch.ExecuteSync();

            // Assert
            // 1. Verify Environment Upserted
            mockDv.Verify(x => x.UpsertRecord(
                "gov_environment", 
                "env-guid-1", 
                It.Is<Dictionary<string, object>>(d => d["gov_name"].ToString() == "Unit Test Env")
            ), Times.Once);

            // 2. Verify App Upserted
            mockDv.Verify(x => x.UpsertRecord(
                "gov_asset", 
                It.IsAny<string>(), 
                It.Is<Dictionary<string, object>>(d => d["gov_name"].ToString() == "Test Canvas App" && d["gov_type"].ToString() == "Canvas App")
            ), Times.Once);
        }
    }
}
