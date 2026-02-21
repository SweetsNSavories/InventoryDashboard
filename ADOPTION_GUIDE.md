# Power Platform Governance Sentinel - Adoption & Deployment Guide

## 🚀 Overview
The **Governance Sentinel** is a high-performance inventory and observability engine for Microsoft Power Platform. It syncs environments, apps, flows, solutions, and tenant-level metrics into a centralized Dataverse-backed dashboard.

## 📁 Repository Structure
*   `GovSolutionsInventory/`: The PCF (Power Apps Control Framework) project for the dashboard UI.
*   `InventoryPuller/`: A C# console application that performs the "Omni-Sync" from Management APIs to Dataverse.
*   `InventorySyncFunction/`: Azure Function implementation for scheduled background syncs.
*   `Solutions/`: Contains the exported Dataverse Solution (`gov_solutionInventory`) in both Managed and Unmanaged formats.
*   `Docs/`: Technical schema documentation and data dictionary.

## 🛠️ Prerequisites
1.  **Administrative Access**: You must be a **Power Platform Administrator** or **Global Administrator**.
2.  **Service Principal (SPN)**: Required for automated syncs.
3.  **Azure CLI**: Installed on the machine running the manual `InventoryPuller`.

## 🔑 Authentication Strategy

### 1. Service Principal (Background Sync)
For the `InventoryPuller` or `Azure Function` to run autonomously, you need an App Registration with:
*   **API Permissions**: 
    *   PowerApps-Runtime Services: `User` (for Dataverse)
    *   Power Platform Management API: `Administrator`
*   **Dataverse Access**: The SPN must be added as an **Application User** in the target inventory environment with the `System Administrator` role.

### 2. Interactive Login (The "Full Hydration" Mode)
The project supports an `--interactive` flag. This uses **Device Auth Code** via the Azure CLI identity (`04b07795-8ddb-461a-bbee-02f9e1bf7b46`).
*   **Why?**: Certain critical tenant-level APIs (like **User Licenses** and **Environment Level Storage/Size Data**) are strictly restricted to interactive user contexts in many tenant configurations. A pure Service Principal (SPN) often receives empty or "Unauthorized" responses for these specific metrics.
*   **Usage**: Run `InventoryPuller.exe --interactive` to perform a full global sync including storage breakdown.

## 🚀 Deployment Steps

### Phase 1: Dataverse Setup
1. **Import Solution**: Import the unmanaged solution from `Solutions/InventorySolution_Unmanaged.zip` into your dedicated Governance environment.
   *   **What this does**: Importing this solution automatically creates the custom tables (`gov_environment`, `gov_solution`, `gov_asset`), registers the PCF Dashboard control, and configures the **User (SystemUser)** form.
2. **Dashboard Access**: 
   *   The Dashboard is embedded in a custom form on the **SystemUser** table named "Asset Inventory Dashboard".
   *   The PCF control is bound to the `Middle Name` (`middlename`) field for maximum compatibility.
   *   **Tip**: Once imported, you can set this form as the default for your Governance admins to provide a seamless "Full Screen" dashboard experience.

### Phase 2: Inventory Puller Configuration
1.  Navigate to `InventoryPuller/`.
2.  Create a `config.json` (Ignored by Git) with the following structure:
```json
{
  "ClientId": "YOUR_SPN_CLIENT_ID",
  "ClientSecret": "YOUR_SPN_SECRET",
  "TenantId": "YOUR_TENANT_ID",
  "DataverseUrl": "https://your-org.crm.dynamics.com"
}
```
3. Run `dotnet run` to start the sync.

## 🛠️ Operational Setup Helpers

### Auto-Mapping Environments
To fetch "Rich Data" (deep inventory) from an environment, the SPN must be a System Admin in that environment's Dataverse instance.
We've included `Map-SpnToEnvironment.ps1` to automate this:
1.  Open the script and populate `YOUR_CLIENT_ID` and `YOUR_CLIENT_SECRET`.
2.  Add your target Environment IDs to the `$TargetEnvs` array.
3.  Run the script. It will create the Application User and grant permissions across all listed orgs in one go.

## 📊 Data Schema
The system pushes data into three primary tables:
1.  **gov_environment**: Tenant-wide container data.
2.  **gov_solution**: Packaging context.
3.  **gov_asset**: Individual Apps, Flows, and Power Pages.

See `Docs/DataverseSchemaDocumentation.html` for a full visual breakdown.

### Phase 3: Azure Function Configuration
1. Deploy the `InventorySyncFunction` project to your Azure Function App.
2. In the **Configuration** blade of the Function App, add the following App Settings:
   *   `PowerPlatform:ClientId`: YOUR_SPN_CLIENT_ID
   *   `PowerPlatform:ClientSecret`: YOUR_SPN_SECRET
   *   `PowerPlatform:TenantId`: YOUR_TENANT_ID
   *   `PowerPlatform:DataverseUrl`: https://your-org.crm.dynamics.com
   *   `PowerPlatform:InteractiveAuth`: `true` (if you want to use the email-based device login for full hydration)
   *   `PowerPlatform:AdminEmail`: The email address of the administrator who will perform the interactive login.
3. **Scheduling & Daily Workflow**:
   *   **The Timer**: By default, the function is scheduled for a daily run (e.g., `0 0 14 * * *` for 9 AM EST). You can adjust this "cron" expression in `SyncFunction.cs` to fit your team's schedule.
   *   **Morning Ritual**: 
      1. Every morning (at the scheduled time), the function fires and sends a "Login Required" email to the **AdminEmail**.
      2. The Administrator checks their inbox, clicks the link, and enters the code.
      3. The Function (which is patiently waiting) detects the login and immediately begins the high-priority sync for storage and license metrics.
   *   **Timeout Note**: The Azure Function will wait for approx 15-20 minutes for the login. If not completed, it will time out and attempt again at the next scheduled interval.

## ❓ Troubleshooting & FAQ

### Why do I need to login interactively with `--interactive`?
While 90% of the data (Apps, Flows, Solutions) can be fetched by a Service Principal, Microsoft currently restricts certain **Tenant-Level Admin APIs** to interactive user contexts:
*   **User License Counts**: Requires an interactive OIB (On behalf of) token in many tenant configurations.
*   **Environment Size/Storage Breakdown**: Many storage-related BAP capacity endpoints only return granular data for the logged-in admin user and will return "0" or "Access Denied" for SPNs.
*   **Tip**: Use `--interactive` once a week for a "Full Hydration" of these global metrics, and use the standard SPN sync for daily asset tracking.

### The Dashboard is blank after import!
1.  Ensure you have run the `InventoryPuller` at least once to populate the Dataverse tables.
2.  Check the `gov_environment` table in Dataverse to verify rows exist.
3.  Ensure the PCF control is published and the form is saved.

## 🛡️ Security Note
**NEVER** check in `config.json`, `local.settings.json`, or any `.zip` files containing sensitive solution data. The `.gitignore` has been pre-configured to protect your environment.
