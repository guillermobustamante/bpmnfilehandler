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

& az appservice plan show --name $PlanName --resource-group $ResourceGroupName 1>$null 2>$null
if ($LASTEXITCODE -ne 0) {
    & az appservice plan create --name $PlanName --resource-group $ResourceGroupName --is-linux --sku $Sku | Out-Null
}

& az webapp show --name $AppName --resource-group $ResourceGroupName 1>$null 2>$null
if ($LASTEXITCODE -ne 0) {
    & az webapp create --name $AppName --resource-group $ResourceGroupName --plan $PlanName --runtime "NODE:22-lts" | Out-Null
}
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
        "SCM_DO_BUILD_DURING_DEPLOYMENT=false" `
        "ENABLE_ORYX_BUILD=false" `
        "WEBSITE_NODE_DEFAULT_VERSION=~22" | Out-Null

$npmCommand = if ($IsWindows) { "npm.cmd" } else { "npm" }
& $npmCommand run build
if ($LASTEXITCODE -ne 0) {
    throw "Local production build failed."
}

$stagingRoot = Join-Path $env:TEMP "bpmnfilehandler-deploy-$([guid]::NewGuid())"
$zipPath = "$stagingRoot.zip"
$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path

New-Item -ItemType Directory -Path $stagingRoot -Force | Out-Null

Copy-Item -LiteralPath (Join-Path $repoRoot "dist") -Destination $stagingRoot -Recurse -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "package.json") -Destination $stagingRoot -Force
Copy-Item -LiteralPath (Join-Path $repoRoot "package-lock.json") -Destination $stagingRoot -Force

Push-Location $stagingRoot
try {
    & $npmCommand ci --omit=dev --ignore-scripts
    if ($LASTEXITCODE -ne 0) {
        throw "Production dependency install failed."
    }
}
finally {
    Pop-Location
}

Compress-Archive -Path (Join-Path $stagingRoot "*") -DestinationPath $zipPath -Force
& az webapp deployment source config-zip --name $AppName --resource-group $ResourceGroupName --src $zipPath | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Azure App Service zip deployment failed."
}

Remove-Item -LiteralPath $stagingRoot -Recurse -Force
Remove-Item -LiteralPath $zipPath -Force

[pscustomobject]@{
    appName = $AppName
    appBaseUrl = "https://$AppName.azurewebsites.net"
    resourceGroupName = $ResourceGroupName
    location = $Location
    sku = $Sku
} | ConvertTo-Json -Depth 5
