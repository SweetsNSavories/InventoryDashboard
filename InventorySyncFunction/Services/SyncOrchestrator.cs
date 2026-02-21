using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using InventorySyncFunction.Interfaces;

namespace InventorySyncFunction.Services
{
    public class SyncOrchestrator
    {
        private readonly IBapClient _bap;
        private readonly IDataverseRepository _dv;
        private readonly ILogger _logger;

        public SyncOrchestrator(IBapClient bap, IDataverseRepository dv, ILogger logger)
        {
            _bap = bap;
            _dv = dv;
            _logger = logger;
        }

        public async Task ExecuteGlobalSync()
        {
            if (!_dv.IsReady)
            {
                _logger.LogError($"❌ Dataverse Unreachable: {_dv.LastError}");
                return;
            }

            string tenantEnvId = "00000000-0000-0000-0000-000000000000";
            _logger.LogInformation("🌐 Starting Global Tenant Sync...");

            try {
                _dv.UpsertRecord("gov_environment", tenantEnvId, new Dictionary<string, object> { { "gov_name", "Global Tenant (System)" }, { "gov_envid", tenantEnvId } });

                // Capacity
                var caps = await _bap.FetchList("api.powerplatform.com", "https://api.powerplatform.com/licensing/tenantCapacity?api-version=2022-03-01-preview");
                if(caps != null) SyncAssets(caps, "Capacity", tenantEnvId);

                // Licensing
                var licenses = await _bap.FetchList("api.powerplatform.com", "https://api.powerplatform.com/licensing/tenantLicenses?api-version=2022-03-01-preview");
                 if(licenses != null) SyncAssets(licenses, "License", tenantEnvId);

                // DLP
                var policies = await _bap.FetchList("api.powerplatform.com", "https://api.powerplatform.com/governance/ruleBasedPolicies?api-version=2022-03-01-preview");
                if (policies != null) SyncAssets(policies, "DLP Policy", tenantEnvId);

            } catch (Exception ex) { _logger.LogError($"Global Sync Error: {ex.Message}"); }
        }

        public async Task ExecuteSync()
        {
            await ExecuteGlobalSync();
            
            var envs = await _bap.FetchList("service.powerapps.com", "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2021-04-01");
            if (envs == null) return;

            // Simple loop to avoid massive parallel overhead in Function
            foreach (var env in envs)
            {
                await ProcessEnvironment((JObject)env);
            }
        }

        public async Task ProcessEnvironment(JObject env)
        {
            try
            {
                string envId = env["name"]?.ToString() ?? "";
                var props = env["properties"];
                string name = props?["displayName"]?.ToString() ?? "Unnamed";
                string type = props?["type"]?.ToString() ?? "Unknown";
                string region = props?["location"]?.ToString() ?? props?["azureRegion"]?.ToString() ?? "";
                string sku = props?["environmentSku"]?.ToString() ?? "";
                string pState = props?["provisioningState"]?.ToString() ?? "";
                bool isDefault = props?["isDefault"]?.ToObject<bool>() ?? false;
                string createdOnRaw = props?["createdTime"]?.ToString() ?? "";
                string modifiedOnRaw = props?["lastModifiedTime"]?.ToString() ?? "";

                _logger.LogInformation($"🌍 SDK Syncing {name} ({type})...");

                // A. Upsert Env
                var fields = new Dictionary<string, object> { 
                    { "gov_name", name }, 
                    { "gov_envid", envId },
                    { "gov_type", type },
                    { "gov_region", region },
                    { "gov_sku", sku },
                    { "gov_provisioningstate", pState },
                    { "gov_isdefault", isDefault },
                    { "gov_metadata", env.ToString() }
                };
                if (DateTime.TryParse(createdOnRaw, out DateTime createdOn)) fields["gov_createdon"] = createdOn;
                if (DateTime.TryParse(modifiedOnRaw, out DateTime modifiedOn)) fields["gov_modifiedon"] = modifiedOn;

                _dv.UpsertRecord("gov_environment", envId, fields);

                // B. Capacity Hydration (from Puller)
                if (props?["capacity"] == null) {
                    _logger.LogInformation($"      ⚡ Capacity missing. Attempting direct hydration...");
                    var hydrated = await _bap.FetchList("api.bap.microsoft.com", $"https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{envId}?api-version=2020-10-01&$expand=properties/capacity");
                    if (hydrated != null) {
                       // Note: BapClient.FetchList returns JArray, but here we expect a single object properties update
                    }
                }

                // C. Apps
                try {
                    var apps = await _bap.FetchList("service.powerapps.com", $"https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/{envId}/apps?api-version=2021-02-01");
                    if (apps != null) SyncAssets(apps, "Canvas App", envId);
                } catch (Exception ex) { _logger.LogWarning($"    ! Apps failed for {envId}: {ex.Message}"); }

                // D. Flows
                try {
                    var flows = await _bap.FetchList("service.flow.microsoft.com", $"https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/{envId}/v2/flows?api-version=2016-11-01");
                    if (flows != null) SyncAssets(flows, "Cloud Flow", envId);
                } catch (Exception ex) { _logger.LogWarning($"    ! Flows failed for {envId}: {ex.Message}"); }
                
                // E. Power Pages (Standard Management)
                try {
                     var sites = await _bap.FetchList("api.powerplatform.com", $"https://api.powerplatform.com/powerpages/environments/{envId}/websites?api-version=2022-03-01-preview");
                     if (sites != null) SyncAssets(sites, "Power Page", envId);
                } catch (Exception ex) { _logger.LogWarning($"    ! Pages failed for {envId}: {ex.Message}"); }

                // F. Solutions
                try {
                    var solns = await _bap.FetchList("api.powerplatform.com", $"https://api.powerplatform.com/appmanagement/environments/{envId}/applicationPackages?api-version=2022-03-01-preview");
                    if (solns != null) SyncAssets(solns, "Solution", envId);
                } catch (Exception ex) { _logger.LogWarning($"    ! Solutions failed for {envId}: {ex.Message}"); }

                // G. Deep Dataverse Sync (Real Names & Solutions)
                string dvUrl = props?["linkedEnvironmentMetadata"]?["instanceUrl"]?.ToString();
                if (!string.IsNullOrEmpty(dvUrl))
                {
                    _logger.LogInformation($"    🔍 Probing Dataverse: {dvUrl}");
                    
                    // 1. Live Solutions
                    var solns = await _bap.FetchDataverse(dvUrl, "solutions?$select=uniquename,friendlyname,ismanaged,version,solutionid&$filter=isvisible eq true");
                    if (solns != null) {
                        foreach (var s in solns) {
                            var sObj = (JObject)s;
                            SyncAssetItem(new JObject {
                                ["name"] = sObj["solutionid"],
                                ["properties"] = new JObject {
                                    ["displayName"] = sObj["friendlyname"],
                                    ["name"] = sObj["uniquename"],
                                    ["isManaged"] = sObj["ismanaged"],
                                    ["version"] = sObj["version"]
                                }
                            }, "Solution", envId, "");
                        }
                    }

                    // 2. Real Flow Names (category 5 = modern flow)
                    var dwf = await _bap.FetchDataverse(dvUrl, "workflows?$select=name,workflowid,solutionid&$filter=category eq 5");
                    if (dwf != null) {
                        foreach (var wf in dwf) {
                            SyncAssetItem(new JObject {
                                ["name"] = wf["workflowid"],
                                ["properties"] = new JObject {
                                    ["displayName"] = wf["name"],
                                    ["solutionId"] = wf["solutionid"]
                                }
                            }, "Cloud Flow", envId, "");
                        }
                    }

                    // 3. Real App Names
                    var dapps = await _bap.FetchDataverse(dvUrl, "canvasapps?$select=displayname,canvasappid,solutionid");
                    if (dapps != null) {
                        foreach (var ap in dapps) {
                            SyncAssetItem(new JObject {
                                ["name"] = ap["canvasappid"],
                                ["properties"] = new JObject {
                                    ["displayName"] = ap["displayname"],
                                    ["solutionId"] = ap["solutionid"]
                                }
                            }, "Canvas App", envId, "");
                        }
                    }
                }

                // F. Users
                try {
                    var users = await _bap.FetchList("api.powerplatform.com", $"https://api.powerplatform.com/usermanagement/environments/{envId}/users?api-version=2022-03-01-preview");
                    SyncAssets(users, "User", envId);
                } catch (Exception ex) { _logger.LogWarning($"    ! Users failed for {envId}: {ex.Message}"); }
            }
            catch (Exception ex) { _logger.LogError($"Error processing env {env["name"]}: {ex.Message}"); }
        }

        private void SyncAssets(JArray? items, string type, string envId)
        {
            if (items == null) return;
            foreach (var item in items) SyncAssetItem(item, type, envId, "");
        }

        private void SyncAssetItem(JToken item, string type, string envId, string prefix)
        {
            var props = item["properties"] ?? item; // Fallback to root if properties missing
            
            // 1. ROBUST NAME RESOLUTION
            // Prioritize displayName > name > friendlyname (Dataverse) > GUID
            string rawName = props?["displayName"]?.ToString()
                           ?? props?["displayname"]?.ToString()
                           ?? item?["displayName"]?.ToString()
                           ?? item?["displayname"]?.ToString()
                           ?? props?["name"]?.ToString() 
                           ?? item?["name"]?.ToString()
                           ?? props?["friendlyname"]?.ToString()
                           ?? props?["skuName"]?.ToString()
                           ?? "Unnamed Asset";

            // If name is just a GUID or technical string, try deep property dive
            if (rawName.Length > 30 && rawName.Contains("-")) {
                 var deepName = props?["name"]?.ToString() ?? props?["displayname"]?.ToString();
                 if (!string.IsNullOrEmpty(deepName) && deepName.Length < 30 && !deepName.Contains("-")) 
                     rawName = deepName;
            }
            
            // Special cases for Power Pages / Solutions
            if (type == "Solution") rawName = props?["friendlyname"]?.ToString() ?? props?["displayName"]?.ToString() ?? rawName;
            if (type == "Power Page") rawName = props?["name"]?.ToString() ?? rawName;

            // 2. STATE & HEALTH
            string state = props?["state"]?.ToString() ?? props?["status"]?.ToString() ?? props?["provisioningState"]?.ToString() ?? "Active";
            string healthStatus = "Healthy";
            string lowerState = state.ToLower();
            if (lowerState.Contains("stopped") || lowerState == "off" || lowerState == "disabled" || lowerState == "suspended") 
                healthStatus = "Disabled";
            else if (lowerState.Contains("failed") || lowerState.Contains("issue")) 
                healthStatus = "Issues";

            // 3. OWNERSHIP & LINKS
            string owner = props?["owner"]?["displayName"]?.ToString() 
                        ?? props?["createdBy"]?["displayName"]?.ToString() 
                        ?? props?["publisherDisplayName"]?.ToString() 
                        ?? props?["creator"]?["userId"]?.ToString() ?? "";

            string createdOnRaw = props?["createdTime"]?.ToString() ?? props?["createdOn"]?.ToString() ?? "";
            string modifiedOnRaw = props?["lastModifiedTime"]?.ToString() ?? props?["modifiedTime"]?.ToString() ?? props?["modifiedOn"]?.ToString() ?? "";
            string solutionId = props?["solutionId"]?.ToString() ?? props?["packageId"]?.ToString() ?? "";
            string version = props?["version"]?.ToString() ?? props?["appVersion"]?.ToString() ?? "";
            bool isManaged = props?["isManaged"]?.ToObject<bool>() ?? false;
            string appPlayUri = props?["appPlayUri"]?.ToString() ?? props?["appOpenUri"]?.ToString() ?? "";

            // 4. FIELD MAPPING
            var fields = new Dictionary<string, object>{
                { "gov_name", (healthStatus != "Healthy" ? "⚠️ " : "✅ ") + rawName },
                { "gov_displayname", rawName },
                { "gov_type", type },
                { "gov_envid", envId },
                { "gov_owner", owner },
                { "gov_state", state },
                { "gov_healthstatus", healthStatus },
                { "gov_solutionid", solutionId },
                { "gov_version", version },
                { "gov_ismanaged", isManaged },
                { "gov_playuri", appPlayUri }
            };

            // 5. STRICT ENVIRONMENT GUARD (Prevent leaked or shared assets from appearing in wrong env)
            var actualEnvId = props?["environment"]?["name"]?.ToString() 
                           ?? props?["environment"]?["id"]?.ToString()?.Split('/').LastOrDefault();
            
            if (!string.IsNullOrEmpty(actualEnvId) && !envId.Contains("0000") && 
                !string.Equals(actualEnvId, envId, StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogTrace($"   ⏭️ Skipping asset '{rawName}' - belongs to env {actualEnvId}, currently syncing {envId}");
                return;
            }

            // 6. ID ISOLATION (Prevent cross-environment collisions)
            string rawAssetId = item["name"]?.ToString() 
                          ?? item["workflowid"]?.ToString() 
                          ?? item["canvasappid"]?.ToString() 
                          ?? item["solutionid"]?.ToString()
                          ?? item["id"]?.ToString() ?? Guid.NewGuid().ToString();

            // Combined Key: Ensure 'DefaultSolution' in Env A doesn't overwrite 'DefaultSolution' in Env B
            // Use deterministic GUID to satisfy Dataverse primary key requirements
            string rawKey = $"{envId}_{rawAssetId}";
            string uniqueKey;
            using (MD5 md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(rawKey));
                uniqueKey = new Guid(hash).ToString();
            }

            fields["gov_assetid"] = rawAssetId; // Original GUID for deep linking

            // 7. METADATA BAG PROTECTION (The 'Response Twig')
            var itemStr = item.ToString();
            bool isRich = itemStr.Length > 500 || item["id"]?.ToString().Contains("/providers/Microsoft.") == true || item["tags"] != null;
            
            if (isRich) {
                 fields["gov_metadata"] = itemStr;
            }

            if (DateTime.TryParse(createdOnRaw, out DateTime createdOn)) fields["gov_createdon"] = createdOn;
            if (DateTime.TryParse(modifiedOnRaw, out DateTime modifiedOn)) fields["gov_modifiedon"] = modifiedOn;

            // 8. UPSERT
            _dv.UpsertRecord("gov_asset", uniqueKey, fields);
        }
    }
}
