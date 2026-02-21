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
*   **Why?**: Certain tenant-level APIs (like **User Licenses** and **Environment Storage Breakdown**) are restricted to interactive user contexts in many tenant configurations and fail with pure SPN access.
*   **Usage**: Run `InventoryPuller.exe --interactive` to perform a full global sync.

## 🚀 Deployment Steps

### Phase 1: Dataverse Setup
1.  Import the unmanaged solution from `Solutions/InventorySolution_Unmanaged.zip` into your dedicated Governance environment.
2.  Assign the `Gov Inventory Dashboard` PCF to the desired form (typically on the `SystemUser` or a custom configuration table).

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

## ❓ Troubleshooting & FAQ

### Why do I need to login interactively with `--interactive`?
While 90% of the data (Apps, Flows, Solutions) can be fetched by a Service Principal, Microsoft currently restricts certain **Tenant-Level Admin APIs** to interactive user contexts:
*   **User License Counts**: Requires an interactive OIB (On behalf of) token in many tenant configurations.
*   **Capacity Breakdown**: Some legacy BAP capacity endpoints only return data for the logged-in admin user.
*   **Tip**: Use `--interactive` once a week for a "Full Hydration" of global metrics, and use the standard SPN sync for daily asset tracking.

### The Dashboard is blank after import!
1.  Ensure you have run the `InventoryPuller` at least once to populate the Dataverse tables.
2.  Check the `gov_environment` table in Dataverse to verify rows exist.
3.  Ensure the PCF control is published and the form is saved.

## 🛡️ Security Note
**NEVER** check in `config.json`, `local.settings.json`, or any `.zip` files containing sensitive solution data. The `.gitignore` has been pre-configured to protect your environment.
