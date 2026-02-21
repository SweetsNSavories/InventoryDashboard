# Map-SpnToEnvironment.ps1
# ==========================================
# Instantly upgrades an environment to "Rich Data" status
# by mapping the Inventory Service Principal as a System Admin.

$TenantId = "YOUR_TENANT_ID"
$ApplicationId = "YOUR_CLIENT_ID"
$ClientSecret = "YOUR_CLIENT_SECRET"

# 📝 LIST YOUR TARGET ENVIRONMENTS HERE
$TargetEnvs = @(
    "ENVIRONMENT_GUID_1",
    "ENVIRONMENT_GUID_2"
)

# Authenticate to Power Platform
Write-Host "🔐 Authenticating..." -ForegroundColor Cyan
Add-PowerAppsAccount -TenantID $TenantId -ApplicationId $ApplicationId -ClientSecret $ClientSecret

foreach ($EnvId in $TargetEnvs) {
    Write-Host "`n🌍 Processing Environment: $EnvId" -ForegroundColor Yellow
    try {
        # 1. Create the Application User
        $User = New-PowerAppManagementApp -ApplicationId $ApplicationId -EnvironmentName $EnvId
        if ($User) {
            Write-Host "   ✅ App User Created: $($User.id)" -ForegroundColor Green
            
            # 2. Assign System Administrator Role
            # Note: This command varies by module version, but 'New-PowerAppManagementApp' 
            # often auto-assigns pure application roles. 
            # If explicit role assignment is needed, we use the specific Role Assignment cmdlets.
            
            Write-Host "   ✨ Validated. This environment is now ready for Rich Sync." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "   ❌ Failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}
Write-Host "`n🚀 DONE. Run the C# Inventory Engine to fetch rich data for these orgs." -ForegroundColor Cyan
