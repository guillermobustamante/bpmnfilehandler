param(
    [Parameter(Mandatory = $true)]
    [string]$DisplayName,

    [Parameter(Mandatory = $true)]
    [string]$AppBaseUrl,

    [string]$TenantId,

    [string[]]$Scopes = @("User.Read", "Files.ReadWrite.All"),

    [switch]$SkipAdminConsent
)

$ErrorActionPreference = "Stop"

function Invoke-AzJson {
    param([string[]]$Arguments)
    $result = & az @Arguments -o json
    if ($LASTEXITCODE -ne 0) {
        throw "Azure CLI command failed: az $($Arguments -join ' ')"
    }
    return $result | ConvertFrom-Json
}

$normalizedBaseUrl = $AppBaseUrl.TrimEnd("/")

if ($TenantId) {
    $currentTenantId = & az account show --query tenantId -o tsv
    if ($LASTEXITCODE -ne 0) {
        throw "Could not read the current Azure tenant."
    }

    if ($currentTenantId -ne $TenantId) {
        throw "Azure CLI is signed in to tenant '$currentTenantId'. Sign in to tenant '$TenantId' before creating the app."
    }
}

$app = Invoke-AzJson @("ad", "app", "create", "--display-name", $DisplayName, "--sign-in-audience", "AzureADMyOrg")
$appObjectId = $app.id
$appId = $app.appId

$graphAppId = "00000003-0000-0000-c000-000000000000"
$graphSp = Invoke-AzJson @("ad", "sp", "show", "--id", $graphAppId)
$resourceAccess = @()

foreach ($scope in $Scopes) {
    $permission = @($graphSp.oauth2PermissionScopes) | Where-Object { $_.value -eq $scope } | Select-Object -First 1
    if (-not $permission) {
        throw "Could not find Microsoft Graph delegated scope '$scope'."
    }

    $resourceAccess += @{
        id = $permission.id
        type = "Scope"
    }
}

$patch = @{
    spa = @{
        redirectUris = @(
            $normalizedBaseUrl,
            "$normalizedBaseUrl/auth.html",
            "http://localhost:5173",
            "http://localhost:5173/auth.html"
        )
    }
    requiredResourceAccess = @(
        @{
            resourceAppId = $graphAppId
            resourceAccess = $resourceAccess
        }
    )
} | ConvertTo-Json -Depth 20 -Compress

$patchFile = New-TemporaryFile
Set-Content -LiteralPath $patchFile -Value $patch -Encoding UTF8

try {
    & az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$appObjectId" --headers "Content-Type=application/json" --body "@$patchFile" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to patch Entra application."
    }
}
finally {
    Remove-Item -LiteralPath $patchFile -Force
}

& az ad sp create --id $appId | Out-Null

if (-not $SkipAdminConsent) {
    & az ad app permission admin-consent --id $appId | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Could not grant tenant admin consent for the delegated Microsoft Graph scopes."
    }
}

[pscustomobject]@{
    appId = $appId
    objectId = $appObjectId
    displayName = $DisplayName
    redirectUris = @($normalizedBaseUrl, "$normalizedBaseUrl/auth.html")
    scopes = $Scopes
    adminConsentGranted = (-not $SkipAdminConsent)
} | ConvertTo-Json -Depth 5
