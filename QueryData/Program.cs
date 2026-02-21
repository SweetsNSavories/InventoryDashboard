using System;
using System.Threading.Tasks;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System.Linq;

class Program
{
    private const string ClientId = "3b451e9c-c4d7-4c12-8d12-69f996e7fd48";
    private const string ClientSecret = "YOUR_CLIENT_SECRET";
    private const string Url = "https://orgd90897e4.crm.dynamics.com";

    static async Task Main(string[] args)
    {
        var service = new ServiceClient(new Uri(Url), ClientId, ClientSecret, true);
        if (!service.IsReady) return;

        Console.WriteLine("🔎 FINAL VERIFICATION: SeaCass Priority Flows");
        Console.WriteLine("==============================================");

        var query = new QueryExpression("gov_asset") {
            ColumnSet = new ColumnSet("gov_displayname", "gov_type", "gov_envid", "gov_metadata")
        };
        query.Criteria.AddCondition("gov_type", ConditionOperator.Equal, "Power Page");

        var result = await service.RetrieveMultipleAsync(query);

        Console.WriteLine($"🔍 Total Power Pages Discovered: {result.Entities.Count}");
        
        foreach(var e in result.Entities) {
            string name = e.GetAttributeValue<string>("gov_displayname") ?? e.GetAttributeValue<string>("gov_name");
            string envId = e.GetAttributeValue<string>("gov_envid");
            string meta = e.GetAttributeValue<string>("gov_metadata");
            
            Console.WriteLine($"\n--- [{name}] ---");
            Console.WriteLine($"Env: {envId}");
            if (!string.IsNullOrEmpty(meta)) {
                Console.WriteLine(meta);
            }
        }
    }
}
