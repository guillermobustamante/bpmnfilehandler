param(
    [Parameter(Mandatory = $true)]
    [string]$TenantHostName
)

$ErrorActionPreference = "Stop"

$hostName = $TenantHostName.Trim().ToLowerInvariant()
if (-not $hostName.EndsWith(".sharepoint.com")) {
    $hostName = "$hostName.sharepoint.com"
}

$token = & az account get-access-token --resource "https://$hostName" --query accessToken -o tsv
if ($LASTEXITCODE -ne 0 -or -not $token) {
    throw "Could not acquire SharePoint access token."
}

$uri = "https://$hostName/_api/v2.0/drive/apps?`$adminForceRefresh=1"
$headers = @{
    Authorization = "Bearer $token"
    Accept = "application/json"
}

$response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers
$response | ConvertTo-Json -Depth 10

