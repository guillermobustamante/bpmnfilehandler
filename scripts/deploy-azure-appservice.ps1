param(
    [Parameter(Mandatory = $true)]
    [string]$ResourceGroupName,

    [Parameter(Mandatory = $true)]
    [string]$AppName,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [string]$SubscriptionId,

    [string]$Location = "canadacentral",

    [string]$PlanName = "$AppName-plan",

    [string]$Sku = "B1",

    [string]$GraphScopes = "User.Read Files.ReadWrite.All"
)

$ErrorActionPreference = "Stop"

if ($SubscriptionId) {
    & az account set --subscription $SubscriptionId | Out-Null
}

$secretBytes = [byte[]]::new(32)
[System.Security.Cryptography.RandomNumberGenerator]::Fill($secretBytes)
$sessionSecret = [Convert]::ToBase64String($secretBytes)

& az group create --name $ResourceGroupName --location $Location | Out-Null
& az appservice plan create --name $PlanName --resource-group $ResourceGroupName --is-linux --sku $Sku | Out-Null
& az webapp create --name $AppName --resource-group $ResourceGroupName --plan $PlanName --runtime "NODE:22-lts" | Out-Null
& az webapp config set --name $AppName --resource-group $ResourceGroupName --startup-file "npm run start" | Out-Null
& az webapp config appsettings set `
    --name $AppName `
    --resource-group $ResourceGroupName `
    --settings `
        "M365_CLIENT_ID=$ClientId" `
        "M365_TENANT_ID=$TenantId" `
        "M365_GRAPH_SCOPES=$GraphScopes" `
        "SESSION_SECRET=$sessionSecret" `
        "NODE_ENV=production" `
        "SCM_DO_BUILD_DURING_DEPLOYMENT=true" `
        "ENABLE_ORYX_BUILD=true" `
        "WEBSITE_NODE_DEFAULT_VERSION=~22" | Out-Null

$stagingRoot = Join-Path $env:TEMP "bpmnfilehandler-deploy-$([guid]::NewGuid())"
$zipPath = "$stagingRoot.zip"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$excluded = @(".git", "node_modules", "dist", ".vite", "coverage")

New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

Get-ChildItem -LiteralPath $repoRoot -Force |
    Where-Object { $excluded -notcontains $_.Name } |
    ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $stagingRoot -Recurse -Force
    }

Compress-Archive -Path (Join-Path $stagingRoot "*") -DestinationPath $zipPath -Force
& az webapp deployment source config-zip --name $AppName --resource-group $ResourceGroupName --src $zipPath | Out-Null

Remove-Item -LiteralPath $stagingRoot -Recurse -Force
Remove-Item -LiteralPath $zipPath -Force

[pscustomobject]@{
    appName = $AppName
    appBaseUrl = "https://$AppName.azurewebsites.net"
    resourceGroupName = $ResourceGroupName
    location = $Location
    sku = $Sku
} | ConvertTo-Json -Depth 5
