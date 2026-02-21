using System;
using System.Collections.Generic;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using InventorySyncFunction.Interfaces;
using Microsoft.Extensions.Logging;

namespace InventorySyncFunction.Services
{
    public class DataverseRepository : IDataverseRepository
    {
        private readonly ServiceClient _client;
        private readonly ILogger _logger;

        public DataverseRepository(string connectionString, ILogger logger)
        {
            _logger = logger;
            try {
                _client = new ServiceClient(connectionString);
            } catch (Exception ex) {
                _logger.LogError($"Dataverse Init Failed: {ex.Message}");
                _client = null;
            }
        }

        public bool IsReady => _client != null && _client.IsReady;
        public string LastError => _client?.LastError ?? "Not Initialized";

        public void UpsertRecord(string entity, string key, Dictionary<string, object> fields)
        {
            if (!IsReady) return;
            try
            {
                Entity rec = new Entity(entity);
                if (Guid.TryParse(key, out Guid g)) rec.Id = g;
                // Fallback for non-guid keys (alternate keys) usually handled differently, 
                // but for this MVP we assume key is guid or we'd use Create/Update. 
                // Actually, our key is 'name' (guid-like) for assets.
                
                foreach (var f in fields) rec[f.Key] = f.Value;
                _client.Execute(new UpsertRequest { Target = rec });
            }
            catch (Exception ex)
            {
                // Log but don't stop sync
                _logger.LogTrace($"Upsert Error ({entity}): {ex.Message}");
            }
        }
    }
}
