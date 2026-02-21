using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;
using System.Text;
using Azure.Identity;
using Azure.Core;
using Microsoft.PowerPlatform.Management;
using Microsoft.PowerPlatform.Management.Models;
using Microsoft.Kiota.Abstractions.Authentication;
using Microsoft.Kiota.Http.HttpClientLibrary;
using Microsoft.Kiota.Serialization.Json;
using Microsoft.Kiota.Abstractions.Serialization;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using DvServiceClient = Microsoft.PowerPlatform.Dataverse.Client.ServiceClient;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;

using ManagementClient = Microsoft.PowerPlatform.Management.ServiceClient;

namespace InventoryPuller
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("🚀 OMNI-RESOURCES RICH INVENTORY ENGINE (V2)");
            Console.WriteLine("===========================================");
            Console.WriteLine($"COMMAND ARGS: {string.Join(" ", args)}");

            // 1. Load configuration
            if (!File.Exists("config.json")) { Console.WriteLine("❌ Error: config.json not found."); return; }
            var config = JObject.Parse(File.ReadAllText("config.json"));
            var ClientId = config["ClientId"]?.ToString() ?? "";
            var ClientSecret = config["ClientSecret"]?.ToString() ?? "";
            var TenantId = config["TenantId"]?.ToString() ?? "";
            var TargetUrl = config["DataverseUrl"]?.ToString() ?? "";

            bool interactive = args.Contains("--interactive");
            TokenCredential credential;
            if (interactive) {
                string PliClientId = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"; // Azure CLI ID (Usually most reliable)
                Console.WriteLine("🔑 INTERACTIVE MODE: Using Azure CLI Identity...");
                credential = new DeviceCodeCredential(new DeviceCodeCredentialOptions {
                    TenantId = TenantId,
                    ClientId = PliClientId,
                    DeviceCodeCallback = (code, cancellation) => {
                        Console.WriteLine($"\n🔑 {code.Message}\n");
                        return Task.CompletedTask;
                    }
                });
            } else {
                credential = new ClientSecretCredential(TenantId, ClientId, ClientSecret);
            }

            try 
            {
                // Setup Kiota/SDK Serialization
                ParseNodeFactoryRegistry.DefaultInstance.ContentTypeAssociatedFactories.TryAdd("application/json", new JsonParseNodeFactory());
                SerializationWriterFactoryRegistry.DefaultInstance.ContentTypeAssociatedFactories.TryAdd("application/json", new JsonSerializationWriterFactory());

                var authProvider = new UnifiedAuthProvider(credential);
                var adapter = new HttpClientRequestAdapter(authProvider) { BaseUrl = "https://api.powerplatform.com" };
                var client = new ManagementClient(adapter);

                using var httpClient = new HttpClient();
                
                // Setup SeaCass Connection - Unified Token Provider
                Microsoft.PowerPlatform.Dataverse.Client.ServiceClient dvClient;
                
                if (interactive) {
                    Console.WriteLine("🔄 Initializing Dataverse Client with Token Provider...");
                    try {
                        dvClient = new Microsoft.PowerPlatform.Dataverse.Client.ServiceClient(
                            new Uri(TargetUrl),
                            async (string authUrl) => {
                                var ctx = new Azure.Core.TokenRequestContext(new[] { $"{TargetUrl}/.default" });
                                var token = await credential.GetTokenAsync(ctx, default);
                                return token.Token;
                            },
                            useUniqueInstance: true
                        );
                    } catch (Exception ex) {
                        Console.WriteLine($"❌ Client Init Failed: {ex.Message}");
                        return;
                    }
                } else {
                    dvClient = new Microsoft.PowerPlatform.Dataverse.Client.ServiceClient($"AuthType=ClientSecret;Url={TargetUrl};ClientId={ClientId};ClientSecret={ClientSecret};");
                }
                
                if (!dvClient.IsReady) { Console.WriteLine($"❌ SeaCass Fail: {dvClient.LastError}"); return; }
                Console.WriteLine($"✅ Connected to SeaCass: {TargetUrl}");

                // 2. Discover Environments (Management Plane)
                Console.WriteLine("\n🌍 FETCHING ENVIRONMENTS...");
                var allEnvs = await GetEnvironments(httpClient, credential);
                
                // --- APPLY FILTERS ---
                // string[] filters = { "seacass", "srpottur_" };
                var environments = allEnvs; // Full Sync Mode
                
                Console.WriteLine($"✅ Processing ALL {environments.Count} environments.");

                if (args.Contains("--seacass-only") || args.Contains("--priority-only")) {
                    Console.WriteLine("🎯 PRIORITY FILTER: Targeting SeaCass and Srpottur...");
                    string[] filters = { "seacass", "srpottur_" };
                    environments = allEnvs.Where(e => {
                        var name = e["properties"]?["displayName"]?.ToString()?.ToLower() ?? "";
                        return filters.Any(f => name.Contains(f));
                    }).ToList();
                }

                List<JToken> allApps = new List<JToken>();
                List<JToken> allFlows = new List<JToken>();
                List<JToken> allSolutions = new List<JToken>();

                if (args.Contains("--licensing-only")) {
                    Console.WriteLine("🎯 LICENSING DIAGNOSTIC MODE: Skipping all environments.");
                    environments.Clear();
                }

                int envCount = 0;
                foreach (var env in environments)
                {
                    envCount++;
                    var envId = env["name"]?.ToString() ?? "";
                    var displayName = env["properties"]?["displayName"]?.ToString() ?? "Unnamed";
                    
                    Console.WriteLine($"\n[{envCount}/{environments.Count}] 🔎 Processing: {displayName} ({envId})");

                    // Mapping from the shared JSON schema
                    var props = env["properties"];
                    var linkedMeta = props?["linkedEnvironmentMetadata"];
                    
                    // Region fix: prefers azureRegion over root location
                    string region = props?["azureRegion"]?.ToString() ?? env["location"]?.ToString() ?? "";

                    // --- NEW: CAPACITY HYDRATION ---
                    // If capacity is missing from the list view, try a direct hydration call
                    if (props?["capacity"] == null) {
                        Console.WriteLine($"      ⚡ Capacity missing in list. Attempting direct hydration via BAP API...");
                        var hydrated = await HydrateEnvironmentCapacity(httpClient, credential, envId);
                        if (hydrated != null && hydrated["properties"]?["capacity"] != null) {
                            if (env["properties"] == null) env["properties"] = new JObject();
                            env["properties"]!["capacity"] = hydrated["properties"]!["capacity"];
                            props = env["properties"];
                            
                            // Log found capacity for verification
                            var caps = props?["capacity"] as JArray;
                            string capSummary = string.Join(", ", caps?.Select(c => $"{c["capacityType"]}: {c["actualConsumption"] ?? c["value"]}") ?? new[] { "None" });
                            Console.WriteLine($"      ✅ Capacity hydrated: {capSummary}");
                        } else {
                            Console.WriteLine($"      ⚠️ Hydration yielded no capacity data for {envId}.");
                        }
                    } else {
                        var caps = props?["capacity"] as JArray;
                        string capSummary = string.Join(", ", caps?.Select(c => $"{c["capacityType"]}: {c["actualConsumption"] ?? c["value"]}") ?? new[] { "None" });
                        Console.WriteLine($"      📦 Capacity present in list: {capSummary}");
                    }
                    // ---------------------------------
                    
                    // Upsert Environment Table
                    UpsertRecord(dvClient, "gov_environment", envId, new Dictionary<string, object> {
                        { "gov_name", displayName },
                        { "gov_envid", envId },
                        { "gov_displayname", displayName },
                        { "gov_type", props?["environmentSku"]?.ToString() ?? "" }, // Better use SKU as primary type
                        { "gov_region", region },
                        { "gov_sku", props?["environmentSku"]?.ToString() ?? "" },
                        { "gov_provisioningstate", props?["provisioningState"]?.ToString() ?? "" },
                        { "gov_version", linkedMeta?["version"]?.ToString() ?? "" },
                        { "gov_url", linkedMeta?["instanceUrl"]?.ToString() ?? "" },
                        { "gov_isdefault", props?["isDefault"]?.ToObject<bool>() ?? false },
                        { "gov_metadata", env.ToString() }
                    });

                    // Parse created time if available
                    if (props?["createdTime"] != null && DateTime.TryParse(props["createdTime"].ToString(), out DateTime createdDt)) {
                        UpsertRecord(dvClient, "gov_environment", envId, new Dictionary<string, object> { { "gov_createdtime", createdDt } });
                    }

                    var apps = new List<JToken>();
                    var flows = new List<JToken>();
                    var solutions = new List<JToken>();

                    if (!args.Contains("--skip-mgmt")) {
                        // A. Fetch Apps (Management API)
                        ShowProgress("Fetching Apps", 10);
                        apps = await GetApps(httpClient, credential, envId);
                        ShowProgress("Fetching Flows", 40);
                        flows = await GetFlows(httpClient, credential, envId);
                        ShowProgress("Fetching Solutions", 70);
                        solutions = await GetSolutions(client, envId);
                        ShowProgress("Completed Mgmt Sync", 100);
                        Console.WriteLine(); // New line after progress

                        Console.WriteLine($"   📦 Management Plane: {apps.Count} Apps, {flows.Count} Flows, {solutions.Count} Solutions.");

                        if (displayName.Contains("SeaCass")) {
                            Console.WriteLine("      ⭐ PRIORITY TARGET (SeaCass) - Committing Assets...");
                        }

                        allApps.AddRange(apps);
                        allFlows.AddRange(flows);
                        allSolutions.AddRange(solutions);
                    } else {
                        Console.WriteLine("   ⏩ Skipping Management Plane Sync (--skip-mgmt)");
                    }

                if (!args.Contains("--skip-mgmt")) {
                    // D. ENRICHMENT PHASE: Dataverse Probing
                    string dvUrl = env["properties"]?["linkedEnvironmentMetadata"]?["instanceUrl"]?.ToString() ?? "";
                    
                    // Cleanup orphans before sync
                    PurgeStaleAssets(dvClient, envId);

                    if (!string.IsNullOrEmpty(dvUrl))
                    {
                        string securityGroupId = env["properties"]?["linkedEnvironmentMetadata"]?["securityGroupId"]?.ToString() ?? "";
                        Console.WriteLine($"   🛰️ Probing Dataverse: {dvUrl} (SG: {securityGroupId})");
                        await ProbeAndEnrich(httpClient, credential, dvClient, envId, dvUrl, apps, flows, securityGroupId);
                    }
                    else 
                    {
                        foreach(var app in apps) PersistAsset(dvClient, (JObject)app, "Canvas App", envId);
                        foreach(var flow in flows) PersistAsset(dvClient, (JObject)flow, "Cloud Flow", envId);
                    }
                } else {
                    Console.WriteLine("   ⏩ Skipping Dataverse Enrichment (--skip-mgmt)");
                }
            }

            Console.WriteLine("\n🛠️ FETCHING GLOBAL GOVERNANCE DATA...");
                var capacity = await GetCapacity(httpClient, credential);
                var licensing = await GetLicensing(httpClient, credential);
                var admin = await GetAdmin(httpClient, credential);

                // Save Local JSON snapshots (as requested)
                await SaveToJson(environments, "environments.json");
                await SaveToJson(allApps, "apps.json");
                await SaveToJson(allFlows, "flows.json");
                await SaveToJson(allSolutions, "solutions.json");

                var capDetails = await GetCapacityDetails(httpClient, credential);
                Console.WriteLine($"   📊 Tenant Capacity: {capDetails.Count} record groups retrieved.");
                await SaveToJson(capDetails, "capacity_details.json");
                await SaveToJson(capacity, "capacity.json");
                await SaveToJson(licensing, "licensing.json");
                await SaveToJson(admin, "admin.json");

                // New: Persist to Dataverse for the Dashboard
                PersistGovernanceOverview(dvClient, capacity, licensing, admin);
 
                Console.WriteLine("\n🏆 TOTAL TENANT SYNC COMPLETED.");
            }
            catch (Exception ex) 
            {
                Console.WriteLine($"\n❌ ERROR: {ex.Message}");
                Console.WriteLine(ex.ToString());
            }
        }

        static void ShowProgress(string task, int percent)
        {
            int width = 30;
            int progress = (int)((float)percent / 100 * width);
            string bar = new string('█', progress) + new string('░', width - progress);
            Console.Write($"\r      ⏳ {task,-25} [{bar}] {percent}%");
        }

        #region Enrichment & Probing

        static async Task ProbeAndEnrich(HttpClient http, TokenCredential cred, DvServiceClient dv, string envId, string dvUrl, List<JToken> bapApps, List<JToken> bapFlows, string securityGroupId)
        {
            try {
                var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { dvUrl.TrimEnd('/') + "/.default" }), default);
                http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);

                // 1. Get Friendly App Names
                ShowProgress("Fetching App Names", 10);
                var appResp = await http.GetAsync($"{dvUrl}/api/data/v9.2/canvasapps?$select=displayname,canvasappid");
                var appMap = new Dictionary<string, string>();
                if (appResp.IsSuccessStatusCode) {
                    var json = JObject.Parse(await appResp.Content.ReadAsStringAsync());
                    foreach(var a in json["value"] ?? new JArray()) {
                        appMap[a["canvasappid"]?.ToString() ?? ""] = a["displayname"]?.ToString() ?? "";
                    }
                }

                // 2. Get Friendly Flow Names
                ShowProgress("Fetching Flow Names", 30);
                var flowResp = await http.GetAsync($"{dvUrl}/api/data/v9.2/workflows?$filter=category eq 5&$select=name,workflowid");
                var flowMap = new Dictionary<string, string>();
                if (flowResp.IsSuccessStatusCode) {
                    var json = JObject.Parse(await flowResp.Content.ReadAsStringAsync());
                    foreach(var f in json["value"] ?? new JArray()) {
                        flowMap[f["workflowid"]?.ToString() ?? ""] = f["name"]?.ToString() ?? "";
                    }
                }

                // 3. Get Solution Components & Dataverse Solutions (Expanded Publisher)
                var dvSolResp = await http.GetAsync($"{dvUrl}/api/data/v9.2/solutions?$select=friendlyname,uniquename,version,ismanaged,createdon,modifiedon,solutionid,_publisherid_value&$expand=publisherid($select=friendlyname)");
                if (dvSolResp.IsSuccessStatusCode) {
                    var solJson = JObject.Parse(await dvSolResp.Content.ReadAsStringAsync());
                    foreach(var s in solJson["value"] ?? new JArray()) {
                        PersistSolution(dv, (JObject)s, envId);
                    }
                }

                // Solution Component Mapping (for Apps/Flows)
                ShowProgress("Mapping Components", 60);
                var compResp = await http.GetAsync($"{dvUrl}/api/data/v9.2/solutioncomponents?$select=objectid,_solutionid_value&$filter=componenttype eq 29 or componenttype eq 300");
                var solCompMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (compResp.IsSuccessStatusCode) {
                    var json = JObject.Parse(await compResp.Content.ReadAsStringAsync());
                    foreach(var c in json["value"] ?? new JArray()) {
                        var objId = c["objectid"]?.ToString() ?? "";
                        var sId = c["_solutionid_value"]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(objId)) solCompMap[objId] = sId;
                    }
                }
                ShowProgress("Enriching Users", 80);
                int userCount = 0;
                var userMap = new Dictionary<string, string>(); // SystemUserId -> Email/Name
                if (!string.IsNullOrEmpty(securityGroupId)) {
                    // 4. Get Active Users & Fetch Licensing from Graph
                    var graphToken = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://graph.microsoft.com/.default" }), default);
                    // Filter for Enabled Read-Write users (0) and Application Users (5)
                    // Note: Removed isapplicationuser as it can cause BadRequest on some versions
                    var userResp = await http.GetAsync($"{dvUrl}/api/data/v9.2/systemusers?$select=fullname,internalemailaddress,azureactivedirectoryobjectid,systemuserid,createdon,modifiedon,accessmode&$filter=isdisabled eq false and (accessmode eq 0 or accessmode eq 5) and not(startswith(fullname,'%23'))");
                    
                    if (userResp.IsSuccessStatusCode) {
                        var userJson = JObject.Parse(await userResp.Content.ReadAsStringAsync());
                        foreach(var u in userJson["value"] ?? new JArray()) {
                            var userId = u["systemuserid"]?.ToString() ?? "";
                            var userName = u["fullname"]?.ToString() ?? "Unnamed User";
                            var aadId = u["azureactivedirectoryobjectid"]?.ToString() ?? "";
                            
                            var licenseList = new JArray();
                            var graphProfile = new JObject();

                            if (!string.IsNullOrEmpty(aadId)) {
                                // A. Call Microsoft Graph for Profile Info
                                var profileUrl = $"https://graph.microsoft.com/v1.0/users/{aadId}?$select=displayName,jobTitle,department,officeLocation,userPrincipalName,mail";
                                using var pReq = new HttpRequestMessage(HttpMethod.Get, profileUrl);
                                pReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", graphToken.Token);
                                var pResp = await http.SendAsync(pReq);
                                if (pResp.IsSuccessStatusCode) {
                                    graphProfile = JObject.Parse(await pResp.Content.ReadAsStringAsync());
                                }

                                // B. Call Microsoft Graph for License Details
                                var graphUrl = $"https://graph.microsoft.com/v1.0/users/{aadId}/licenseDetails";
                                using var gReq = new HttpRequestMessage(HttpMethod.Get, graphUrl);
                                gReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", graphToken.Token);
                                var gResp = await http.SendAsync(gReq);
                                if (gResp.IsSuccessStatusCode) {
                                    var gJson = JObject.Parse(await gResp.Content.ReadAsStringAsync());
                                    foreach(var l in gJson["value"] ?? new JArray()) {
                                        licenseList.Add(l["skuPartNumber"]?.ToString()?.Replace("_", " ") ?? "Unknown License");
                                    }
                                }
                            }

                            var userItem = new JObject {
                                ["name"] = userId,
                                ["properties"] = new JObject {
                                    ["displayName"] = userName,
                                    ["email"] = u["internalemailaddress"]?.ToString(),
                                    ["aadObjectId"] = aadId,
                                    ["isApplicationUser"] = (u["accessmode"]?.ToString() == "5"),
                                    ["accessMode"] = u["accessmode"]?.ToString(),
                                    ["assignedLicenses"] = licenseList,
                                    ["graphProfile"] = graphProfile,
                                    ["createdTime"] = u["createdon"]?.ToString(),
                                    ["lastModifiedTime"] = u["modifiedon"]?.ToString()
                                }
                            };
                            PersistAsset(dv, userItem, "User", envId, userName);
                            
                            // Map for Flow Owner Resolution: Key by AAD ID (which flows use) 
                            // fallback to SystemUserID if AAD ID is missing
                            string mapKey = !string.IsNullOrEmpty(aadId) ? aadId : userId;
                            userMap[mapKey] = u["internalemailaddress"]?.ToString() ?? userName;
                            userCount++;
                        }
                    } else {
                        Console.WriteLine($"      ⚠️ User Probe Failed: {userResp.StatusCode} - {await userResp.Content.ReadAsStringAsync()}");
                    }
                } else {
                    Console.WriteLine("   ⏩ Skipping User Enrichment: No Security Group assigned to environment.");
                }

                // Persist with Enrichment
                foreach(var app in bapApps) {
                    var appId = app["name"]?.ToString() ?? "";
                    string friendly = appMap.ContainsKey(appId) ? appMap[appId] : (app["properties"]?["displayName"]?.ToString() ?? "");
                    string solutionId = solCompMap.ContainsKey(appId) ? solCompMap[appId] : "";
                    PersistAsset(dv, (JObject)app, "Canvas App", envId, friendly, solutionId);
                }

                foreach(var flow in bapFlows) {
                    var flowId = flow["name"]?.ToString() ?? "";
                    string friendly = flowMap.ContainsKey(flowId) ? flowMap[flowId] : (flow["properties"]?["displayName"]?.ToString() ?? "");
                    
                    // Owner Resolution: Resolve GUID to Email
                    var creatorId = flow["properties"]?["creator"]?["userId"]?.ToString();
                    if (!string.IsNullOrEmpty(creatorId) && userMap.ContainsKey(creatorId)) {
                        // Inject email into the JSON so PersistAsset picks it up
                        if (flow["properties"]?["creator"] is JObject creatorObj) {
                            creatorObj["email"] = userMap[creatorId];
                        }
                    }

                    string solutionId = solCompMap.ContainsKey(flowId) ? solCompMap[flowId] : ""; // Default handled mainly in Dashboard, but let's leave empty if none
                    PersistAsset(dv, (JObject)flow, "Cloud Flow", envId, friendly, solutionId);
                }

                // 5. Get Power Pages (Advanced Management Plane + Dataverse Probing)
                int portalCount = 0;
                var ppTokenRequest = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.powerplatform.com/.default" }), default);
                
                // Layer 1: Power Platform Management API (Unified Discovery)
                var ppMgmtReqUrl = $"https://api.powerplatform.com/powerpages/environments/{envId}/websites?api-version=2022-03-01-preview";
                using var ppMgmtReq = new HttpRequestMessage(HttpMethod.Get, ppMgmtReqUrl);
                ppMgmtReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", ppTokenRequest.Token);
                
                var ppMgmtResp = await http.SendAsync(ppMgmtReq);

                if (ppMgmtResp.IsSuccessStatusCode) {
                    var ppJson = JObject.Parse(await ppMgmtResp.Content.ReadAsStringAsync());
                    var values = ppJson["value"] as JArray;
                    if (values != null && values.Count > 0) {
                        foreach(var p in values) {
                            var portalId = p["dataverseRecordId"]?.ToString() ?? p["id"]?.ToString() ?? "";
                            var portalName = p["name"]?.ToString() ?? "Power Page Website";
                            var siteUrl = p["websiteUrl"]?.ToString() ?? "";
                            
                            var portalItem = new JObject {
                                ["name"] = portalId,
                                ["properties"] = new JObject {
                                    ["displayName"] = portalName,
                                    ["portalId"] = portalId,
                                    ["siteUrl"] = siteUrl,
                                    ["state"] = $"Active ({p["applicationType"]})",
                                    ["status"] = p["websiteStatus"]?.ToString(),
                                    ["visibility"] = p["websiteVisibility"]?.ToString(),
                                    ["provisioningState"] = p["packageInstallationStatus"]?.ToString(),
                                    ["createdTime"] = p["createdTime"]?.ToString(),
                                    ["lastModifiedTime"] = p["lastModifiedTime"]?.ToString() // Added
                                }
                            };
                            PersistAsset(dv, portalItem, "Power Page", envId, portalName);
                            portalCount++;
                        }
                    } else {
                         // Console.WriteLine($"   ℹ️ Power Pages Mgmt API returned 0 sites for {envId}.");
                    }
                } else {
                     string err = await ppMgmtResp.Content.ReadAsStringAsync();
                     Console.WriteLine($"   ⚠️ Power Pages Mgmt API failed ({ppMgmtResp.StatusCode}): {err.Substring(0, Math.Min(err.Length, 150))}");
                }

                // Layer 2: Dataverse Probing (Fallback/Legacy Discovery)
                // Only run if Layer 1 found nothing or to ensure legacy sites are captured
                if (portalCount == 0) {
                    // v1 (Standard Data Model)
                    var portalRespV1 = await http.GetAsync($"{dvUrl}/api/data/v9.2/adx_websites?$select=adx_name,adx_websiteid,adx_primarydomainname,adx_partialurl,createdon,modifiedon");
                    if (portalRespV1.IsSuccessStatusCode) {
                        var portalJson = JObject.Parse(await portalRespV1.Content.ReadAsStringAsync());
                        foreach(var p in portalJson["value"] ?? new JArray()) {
                            var portalId = p["adx_websiteid"]?.ToString() ?? "";
                            var portalName = p["adx_name"]?.ToString() ?? "Legacy Power Page";
                            var domain = p["adx_primarydomainname"]?.ToString();
                            var partial = p["adx_partialurl"]?.ToString();
                            string siteUrl = !string.IsNullOrEmpty(domain) ? (domain.StartsWith("http") ? domain : $"https://{domain}") : "";
                            
                            var portalItem = new JObject {
                                ["name"] = portalId,
                                ["properties"] = new JObject {
                                    ["displayName"] = portalName,
                                    ["portalId"] = portalId,
                                    ["siteUrl"] = siteUrl,
                                    ["partialUrl"] = partial,
                                    ["state"] = "Active (v1)",
                                    ["createdTime"] = p["createdon"]?.ToString(),
                                    ["lastModifiedTime"] = p["modifiedon"]?.ToString()
                                }
                            };
                            PersistAsset(dv, portalItem, "Power Page", envId, portalName);
                            portalCount++;
                        }
                    }

                    // v2 (Enhanced/Modern Data Model)
                    var portalRespV2 = await http.GetAsync($"{dvUrl}/api/data/v9.2/powerpagesites?$select=name,powerpagesiteid,powerpagesiteurl,hostname,createdon,modifiedon");
                    if (portalRespV2.IsSuccessStatusCode) {
                        var portalJson = JObject.Parse(await portalRespV2.Content.ReadAsStringAsync());
                        foreach(var p in portalJson["value"] ?? new JArray()) {
                            var portalId = p["powerpagesiteid"]?.ToString() ?? "";
                            var portalName = p["name"]?.ToString() ?? "Enhanced Power Page";
                            var siteUrl = (p["powerpagesiteurl"]?.ToString() ?? p["hostname"]?.ToString()) ?? "";
                            if (!string.IsNullOrEmpty(siteUrl) && !siteUrl.StartsWith("http")) siteUrl = $"https://{siteUrl}";

                            var portalItem = new JObject {
                                ["name"] = portalId,
                                ["properties"] = new JObject { 
                                    ["displayName"] = portalName, 
                                    ["portalId"] = portalId, 
                                    ["siteUrl"] = siteUrl, 
                                    ["state"] = "Active (v2)",
                                    ["createdTime"] = p["createdon"]?.ToString(),
                                    ["lastModifiedTime"] = p["modifiedon"]?.ToString()
                                }
                            };
                            PersistAsset(dv, portalItem, "Power Page", envId, portalName);
                            portalCount++;
                        }
                    }

                    // Fallback for some environments: mspp_websites
                    var portalRespMspp = await http.GetAsync($"{dvUrl}/api/data/v9.2/mspp_websites?$select=mspp_name,mspp_websiteid,mspp_primarydomainname,createdon,modifiedon");
                    if (portalRespMspp.IsSuccessStatusCode) {
                        var portalJson = JObject.Parse(await portalRespMspp.Content.ReadAsStringAsync());
                        foreach(var p in portalJson["value"] ?? new JArray()) {
                            var portalId = p["mspp_websiteid"]?.ToString() ?? "";
                            var portalName = p["mspp_name"]?.ToString() ?? "Modern Power Page";
                            var domain = p["mspp_primarydomainname"]?.ToString() ?? "";
                            string siteUrl = !string.IsNullOrEmpty(domain) ? (domain.StartsWith("http") ? domain : $"https://{domain}") : "";
                            
                            var portalItem = new JObject {
                                ["name"] = portalId,
                                ["properties"] = new JObject { 
                                    ["displayName"] = portalName, 
                                    ["portalId"] = portalId, 
                                    ["siteUrl"] = siteUrl, 
                                    ["state"] = "Active (Modern)",
                                    ["createdTime"] = p["createdon"]?.ToString(),
                                    ["lastModifiedTime"] = p["modifiedon"]?.ToString()
                                }
                            };
                            PersistAsset(dv, portalItem, "Power Page", envId, portalName);
                            portalCount++;
                        }
                    }
                }

                Console.WriteLine($"   ✅ Enriched {bapApps.Count} Apps, {bapFlows.Count} Flows, {userCount} Users, and {portalCount} Portals.");

            } catch (Exception ex) {
                Console.WriteLine($"   ⚠️ Probe failed: {ex.Message}");
            }
        }

        static void PersistSolution(DvServiceClient dv, JObject item, string envId)
        {
            // Handle both BAP and Dataverse fields
            string rawId = item["solutionid"]?.ToString() ?? item["Id"]?.ToString() ?? item["ApplicationId"]?.ToString() ?? "";
            string cleanId = ExtractGuid(rawId);
            
            string name = item["friendlyname"]?.ToString() ?? item["ApplicationName"]?.ToString() ?? item["LocalizedName"]?.ToString() ?? "Unnamed Solution";
            
            // Check for expanded publisher name first, then fallback to BAP or GUID
            string publisher = item["publisherid"]?["friendlyname"]?.ToString() ?? 
                               item["PublisherName"]?.ToString() ?? 
                               item["_publisherid_value"]?.ToString() ?? "Unknown Publisher";

            string version = item["version"]?.ToString() ?? item["Version"]?.ToString() ?? "";
            string uniqueName = item["uniquename"]?.ToString() ?? item["UniqueName"]?.ToString() ?? "";
            string description = item["description"]?.ToString() ?? item["ApplicationDescription"]?.ToString() ?? "";
            string url = item["LearnMoreUrl"]?.ToString() ?? "";
            string state = item["ismanaged"]?.ToString() ?? item["State"]?.ToString() ?? "0";

            DateTime? solCreated = null;
            DateTime? solModified = null;
            // Support PascalCase, camelCase, and direct Dataverse lowercase
            string cStr = item["createdon"]?.ToString() ?? item["CreatedOn"]?.ToString() ?? item["createdOn"]?.ToString() ?? "";
            string mStr = item["modifiedon"]?.ToString() ?? item["ModifiedOn"]?.ToString() ?? item["modifiedOn"]?.ToString() ?? "";

            if (DateTime.TryParse(cStr, out DateTime cdt)) solCreated = cdt;
            if (DateTime.TryParse(mStr, out DateTime mdt)) solModified = mdt;

            var fields = new Dictionary<string, object> {
                { "gov_name", name },
                { "gov_displayname", name },
                { "gov_uniquename", uniqueName },
                { "gov_version", version },
                { "gov_owner", publisher },
                { "gov_description", description },
                { "gov_url", url },
                { "gov_state", state == "True" || state == "1" ? "Managed" : (state == "False" || state == "0" ? "Unmanaged" : state) },
                { "gov_envid", envId },
                { "gov_createdon", solCreated },
                { "gov_modifiedon", solModified },
                { "gov_metadata", item.ToString() }
            };

            UpsertRecord(dv, "gov_solution", cleanId, fields);
        }

        static void PersistAsset(DvServiceClient dv, JObject item, string type, string envId, string overrideName = "", string solutionId = "")
        {
            string rawId = item["name"]?.ToString() ?? item["id"]?.ToString() ?? "";
            string cleanId = ExtractGuid(rawId);
            
            var props = item["properties"];
            string name = string.IsNullOrEmpty(overrideName) ? (props?["displayName"]?.ToString() ?? "Unnamed") : overrideName;
            
            // Richer Owner Logic: Prefer email, fallback to ID
            string owner = props?["owner"]?["email"]?.ToString() 
                        ?? props?["creator"]?["email"]?.ToString() 
                        ?? props?["creator"]?["userId"]?.ToString() 
                        ?? "";

            // Mapping rich properties from your JSON
            string state = props?["state"]?.ToString() ?? props?["status"]?.ToString() ?? "";
            bool isManaged = props?["isManaged"]?.ToObject<bool>() ?? (props?["almMode"]?.ToString() == "Solution");
            
            // Canvas App Specifics
            string appType = item["appType"]?.ToString() ?? "";
            string formFactor = item["tags"]?["primaryFormFactor"]?.ToString() ?? "";
            string almMode = props?["almMode"]?.ToString() ?? "";
            bool usesPremium = props?["usesPremiumApi"]?.ToObject<bool>() ?? false;
            string dlpStatus = props?["executionRestrictions"]?["dataLossPreventionEvaluationResult"]?["status"]?.ToString() ?? "";
            string version = props?["appVersion"]?.ToString() ?? "";

            // Parse timestamps for the dashboard
            DateTime? created = null;
            DateTime? modified = null;
            if (DateTime.TryParse(props?["createdTime"]?.ToString() ?? "", out DateTime cdt)) created = cdt;
            if (DateTime.TryParse(props?["lastModifiedTime"]?.ToString() ?? "", out DateTime mdt)) modified = mdt;

            var fields = new Dictionary<string, object> {
                { "gov_name", name },
                { "gov_displayname", name },
                { "gov_type", type },
                { "gov_envid", envId },
                { "gov_owner", owner },
                { "gov_state", state },
                { "gov_ismanaged", isManaged },
                { "gov_apptype", appType },
                { "gov_formfactor", formFactor },
                { "gov_almmode", almMode },
                { "gov_usespremiumapi", usesPremium },
                { "gov_dlpstatus", dlpStatus },
                { "gov_version", version },
                { "gov_createdon", created },
                { "gov_modifiedon", modified },
                { "gov_solutionid", solutionId },
                { "gov_metadata", item.ToString() }
            };

            if (modified.HasValue) fields["gov_modifiedon"] = modified.Value;

            // Only add gov_assetid if we found a valid GUID
            if (Guid.TryParse(cleanId, out Guid assetGuid)) {
                fields["gov_assetid"] = assetGuid;
            }

            UpsertRecord(dv, "gov_asset", cleanId, fields);
        }

        static string ExtractGuid(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            
            // If it's already a GUID, return it
            if (Guid.TryParse(input, out _)) return input;
            
            // Try to find a GUID pattern in the string
            var match = System.Text.RegularExpressions.Regex.Match(input, @"[{(]?[0-9A-Fa-f]{8}[-]?([0-9A-Fa-f]{4}[-]?){3}[0-9A-Fa-f]{12}[)}]?");
            if (match.Success) return match.Value;

            // Fallback: Split and search
            var parts = input.Split(new[] { '/', '(', ')', ':' }, StringSplitOptions.RemoveEmptyEntries);
            foreach(var part in parts) {
                if (Guid.TryParse(part, out _)) return part;
            }
            return input;
        }

        #endregion

        #region API Retrieval (From User Logic)

        static async Task<List<JToken>> GetEnvironments(HttpClient client, TokenCredential cred)
        {
            var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.powerplatform.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            
            // 1. Try Unified Management API (Preview)
            string url = "https://api.powerplatform.com/environmentmanagement/environments?api-version=2022-03-01-preview&$expand=properties/capacity";
            var resp = await client.GetAsync(url);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                var envs = json["value"]?.ToList() ?? new List<JToken>();
                if (envs.Any(e => e["properties"]?["capacity"] != null)) return envs;
            }

            // 2. Try BAP Fallback (Stable 2020)
            url = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2020-10-01&$expand=properties/capacity";
            var bapToken = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.bap.microsoft.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bapToken.Token);
            resp = await client.GetAsync(url);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                var envs = json["value"]?.ToList() ?? new List<JToken>();
                if (envs.Any(e => e["properties"]?["capacity"] != null)) return envs;
            }

            // 3. Try BAP with alternative expansion
            url = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2016-11-01&$expand=permissions,properties.capacity";
            resp = await client.GetAsync(url);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                return json["value"]?.ToList() ?? new List<JToken>();
            }

            return new List<JToken>();
        }

        static async Task<JToken?> HydrateEnvironmentCapacity(HttpClient client, TokenCredential cred, string envId)
        {
            try {
                // Direct GET environment/capacity
                string url = $"https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/{envId}?api-version=2020-10-01&$expand=properties/capacity";
                var bapToken = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.bap.microsoft.com/.default" }), default);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bapToken.Token);
                var resp = await client.GetAsync(url);
                if (resp.IsSuccessStatusCode) {
                    return JObject.Parse(await resp.Content.ReadAsStringAsync());
                }
            } catch { }
            return null;
        }

        static async Task<List<JToken>> GetApps(HttpClient client, TokenCredential cred, string envId)
        {
            var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://service.powerapps.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            string url = $"https://api.powerapps.com/providers/Microsoft.PowerApps/scopes/admin/environments/{envId}/apps?api-version=2021-02-01";
            var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode) return new List<JToken>();
            var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
            return json["value"]?.ToList() ?? new List<JToken>();
        }

        static async Task<List<JToken>> GetFlows(HttpClient client, TokenCredential cred, string envId)
        {
            var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://service.flow.microsoft.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            string url = $"https://api.flow.microsoft.com/providers/Microsoft.ProcessSimple/scopes/admin/environments/{envId}/v2/flows?api-version=2016-11-01";
            var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode) return new List<JToken>();
            var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
            return json["value"]?.ToList() ?? new List<JToken>();
        }

        static async Task<List<JToken>> GetSolutions(ManagementClient client, string envId)
        {
            try {
                var result = await client.Appmanagement.Environments[envId].ApplicationPackages.GetAsync(config => {
                    config.QueryParameters.ApiVersion = "2022-03-01-preview";
                });
                var list = result?.Value ?? new List<ApplicationPackage>();
                return list.Select(pkg => {
                    var jobj = JObject.FromObject(pkg);
                    jobj["environmentId"] = envId;
                    return jobj as JToken;
                }).ToList();
            } catch { return new List<JToken>(); }
        }

        static async Task<List<JToken>> GetCapacity(HttpClient _, TokenCredential cred)
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "InventoryPuller/2.0");
            
            var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.powerplatform.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            
            // 1. Try Capacity Summary (Modern)
            string url = "https://api.powerplatform.com/licensing/tenantCapacity?api-version=2022-03-01-preview";
            var resp = await client.GetAsync(url);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                var val = json["value"]?.ToList() ?? json["tenantCapacities"]?.ToList();
                if (val != null) return val;
            }

            // 2. Try Currency Reports (User Suggestion)
            string reportUrl = "https://api.powerplatform.com/licensing/currencyReports?api-version=2022-03-01-preview";
            resp = await client.GetAsync(reportUrl);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                return json["value"]?.ToList() ?? new List<JToken>();
            }

            // 3. Fallback to BAP (Stable)
            var bapToken = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.bap.microsoft.com/.default" }), default);
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Add("User-Agent", "InventoryPuller/2.0");
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bapToken.Token);
            string bapUrl = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments/capacities?api-version=2020-10-01";
            resp = await client.GetAsync(bapUrl);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                return json["value"]?.ToList() ?? new List<JToken>();
            }

            return new List<JToken>();
        }

        static async Task<List<JToken>> GetCapacityDetails(HttpClient _, TokenCredential cred)
        {
             // Combine into Capacity for now or use specific summary if found
             return new List<JToken>();
        }

        static async Task<List<JToken>> GetLicensing(HttpClient _, TokenCredential cred)
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("User-Agent", "InventoryPuller/2.0");
            
            var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.powerplatform.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            
            // 1. Try Modern Licensing
            string url = "https://api.powerplatform.com/licensing/tenantLicenses?api-version=2022-03-01-preview";
            var resp = await client.GetAsync(url);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                return json["value"]?.ToList() ?? new List<JToken>();
            } else {
                Console.WriteLine($"      ℹ️ LicensingModern failed ({resp.StatusCode})");
            }

            // 3. Try Product Inventory
            string prodUrl = "https://api.powerplatform.com/licensing/productInventory?api-version=2022-03-01-preview";
            resp = await client.GetAsync(prodUrl);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                return json["value"]?.ToList() ?? new List<JToken>();
            }

            // 4. Fallback to BAP
            var bapToken = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.bap.microsoft.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bapToken.Token);
            string bapUrl = "https://api.bap.microsoft.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/licensing/tenantLicenses?api-version=2020-10-01";
            resp = await client.GetAsync(bapUrl);
            if (resp.IsSuccessStatusCode) {
                var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
                return json["value"]?.ToList() ?? new List<JToken>();
            } else {
                 Console.WriteLine($"      ℹ️ LicensingBap failed ({resp.StatusCode})");
            }

            return new List<JToken>();
        }

        static async Task<List<JToken>> GetAdmin(HttpClient client, TokenCredential cred)
        {
            client.DefaultRequestHeaders.Clear();
            var token = await cred.GetTokenAsync(new Azure.Core.TokenRequestContext(new[] { "https://api.powerplatform.com/.default" }), default);
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
            string url = "https://api.powerplatform.com/governance/ruleBasedPolicies?api-version=2022-03-01-preview";
            var resp = await client.GetAsync(url);
            if (!resp.IsSuccessStatusCode) {
                var err = await resp.Content.ReadAsStringAsync();
                Console.WriteLine($"   ⚠️ Admin API Check failed ({resp.StatusCode}): {err}");
                return new List<JToken>();
            }
            var json = JObject.Parse(await resp.Content.ReadAsStringAsync());
            return json["value"]?.ToList() ?? new List<JToken>();
        }

        #endregion

        #region Data Persistence & Saving

        static void PurgeStaleAssets(DvServiceClient dv, string envId)
        {
            try {
                // Delete assets for this environment that have NULL metadata ( Orphans from failed syncs )
                var query = new Microsoft.Xrm.Sdk.Query.QueryExpression("gov_asset") {
                    ColumnSet = new Microsoft.Xrm.Sdk.Query.ColumnSet("gov_assetid")
                };
                query.Criteria.AddCondition("gov_envid", Microsoft.Xrm.Sdk.Query.ConditionOperator.Equal, envId);
                query.Criteria.AddCondition("gov_metadata", Microsoft.Xrm.Sdk.Query.ConditionOperator.Null);
                
                var results = dv.RetrieveMultiple(query);
                if (results.Entities.Count > 0) {
                    Console.WriteLine($"   🧹 Purging {results.Entities.Count} stale/broken assets...");
                    foreach (var entity in results.Entities) {
                        dv.Delete("gov_asset", entity.Id);
                    }
                }
            } catch { /* Silent skip */ }
        }

        static void UpsertRecord(DvServiceClient dv, string entityName, string externalId, Dictionary<string, object> fields)
        {
            try {
                Entity record = new Entity(entityName);
                if (Guid.TryParse(externalId, out Guid idGuid)) record.Id = idGuid;
                foreach (var field in fields) record[field.Key] = field.Value;
                dv.Execute(new UpsertRequest { Target = record });
            } catch (Exception ex) {
                Console.WriteLine($"   ⚠️ Upsert failed for {externalId}: {ex.Message}");
            }
        }

        static void PersistGovernanceOverview(DvServiceClient dv, List<JToken> capacity, List<JToken> licensing, List<JToken> admin)
        {
            try {
                var tenantId = "00000000-0000-0000-0000-000000000000";
                
                // Smart Scavenging: If licensing is empty, try to extract from capacity entitlements
                var finalLicensing = new List<JToken>(licensing);
                if (finalLicensing.Count == 0 && capacity.Count > 0) {
                    var scavenged = new HashSet<string>();
                    foreach (var cap in capacity) {
                        var entitlements = cap["capacityEntitlements"];
                        if (entitlements != null) {
                            foreach (var ent in entitlements) {
                                var licenses = ent["licenses"];
                                if (licenses != null) {
                                    foreach (var lic in licenses) {
                                        var sku = lic["skuId"]?.ToString();
                                        if (sku != null && !scavenged.Contains(sku)) {
                                            finalLicensing.Add(lic);
                                            scavenged.Add(sku);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                var meta = new JObject {
                    ["capacity"] = new JArray(capacity),
                    ["licensing"] = new JArray(finalLicensing),
                    ["governance"] = new JArray(admin),
                    ["lastSync"] = DateTime.UtcNow
                };

                var fields = new Dictionary<string, object> {
                    { "gov_name", "Tenant Sentinel Overview" },
                    { "gov_displayname", "Global Tenant Governance" },
                    { "gov_type", "Tenant" },
                    { "gov_metadata", meta.ToString(Formatting.None) },
                    { "gov_isdefault", true }
                };

                UpsertRecord(dv, "gov_environment", tenantId, fields);
                Console.WriteLine("✅ Global Governance data (with scavenged licenses) pushed to Dataverse.");
            } catch (Exception ex) {
                Console.WriteLine($"   ⚠️ Failed to persist governance overview: {ex.Message}");
            }
        }

        static async Task SaveToJson<T>(List<T> data, string fileName)
        {
            if (data == null || data.Count == 0) return;
            var json = JsonConvert.SerializeObject(data, Formatting.Indented);
            await File.WriteAllTextAsync(fileName, json);
            Console.WriteLine($"✅ Saved {data.Count} items to {fileName}");
        }

        #endregion
    }

    public class UnifiedAuthProvider : IAuthenticationProvider
    {
        private readonly TokenCredential _cred;
        public UnifiedAuthProvider(TokenCredential cred) => _cred = cred;
        public async Task AuthenticateRequestAsync(Microsoft.Kiota.Abstractions.RequestInformation request, Dictionary<string, object>? context = null, System.Threading.CancellationToken token = default)
        {
            var scopes = new[] { "https://api.powerplatform.com/.default" };
            var result = await _cred.GetTokenAsync(new Azure.Core.TokenRequestContext(scopes), token);
            request.Headers.Add("Authorization", $"Bearer {result.Token}");
        }
    }
}
