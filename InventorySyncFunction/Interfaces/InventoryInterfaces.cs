using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace InventorySyncFunction.Interfaces
{
    public interface IBapClient
    {
        Task<JArray?> FetchList(string scope, string url);
        Task<JArray?> FetchDataverse(string instanceUrl, string query);
        Task<bool> CheckAccess(string instanceUrl);
        Microsoft.PowerPlatform.Management.ServiceClient Management { get; }
    }

    public interface IDataverseRepository
    {
        bool IsReady { get; }
        string LastError { get; }
        void UpsertRecord(string entity, string key, Dictionary<string, object> fields);
    }
}
