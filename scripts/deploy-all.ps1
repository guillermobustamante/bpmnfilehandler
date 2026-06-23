param(
    [string]$SubscriptionId,

    [string]$TenantId = "911a2cd5-fcc5-4ccb-ae76-b740c096c559",

    [string]$TenantHostName = "evolvegs.sharepoint.com",

    [string]$ResourceGroupName = "rg-bpmn-file-handler",

    [string]$Location = "canadacentral",

    [string]$AppName = "bpmn-file-handler-$((New-Guid).Guid.Substring(0, 8))",

    [string]$DisplayName = "BPMN File Handler",

    [string[]]$Extensions = @(".bpmn", ".drawio")
)

$ErrorActionPreference = "Stop"

$appBaseUrl = "https://$AppName.azurewebsites.net"

Write-Host "Creating Entra application..."
$entra = & "$PSScriptRoot\create-entra-app.ps1" `
    -DisplayName $DisplayName `
    -AppBaseUrl $appBaseUrl `
    -TenantId $TenantId | ConvertFrom-Json

Write-Host "Deploying Azure App Service..."
$deployArgs = @{
    ResourceGroupName = $ResourceGroupName
    AppName = $AppName
    ClientId = $entra.appId
    TenantId = $TenantId
    Location = $Location
}
if ($SubscriptionId) {
    $deployArgs.SubscriptionId = $SubscriptionId
}

$deployment = & "$PSScriptRoot\deploy-azure-appservice.ps1" @deployArgs | ConvertFrom-Json

Write-Host "Registering Microsoft 365 File Handler..."
$handler = & "$PSScriptRoot\register-file-handler.ps1" `
    -ApplicationObjectId $entra.objectId `
    -AppBaseUrl $deployment.appBaseUrl `
    -Extensions $Extensions | ConvertFrom-Json

Write-Host "Refreshing tenant file handler cache..."
$refresh = $null
try {
    $refresh = & "$PSScriptRoot\refresh-file-handler-cache.ps1" -TenantHostName $TenantHostName
}
catch {
    $refresh = @{
        status = "failed"
        message = $_.Exception.Message
        note = "File Handler registration remains valid. Microsoft 365 may take 24-48 hours to show new handlers without a successful cache refresh."
    }
}

[pscustomobject]@{
    entraApp = $entra
    deployment = $deployment
    fileHandler = $handler
    cacheRefresh = $refresh
} | ConvertTo-Json -Depth 20
