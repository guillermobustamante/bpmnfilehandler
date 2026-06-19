param(
    [Parameter(Mandatory = $true)]
    [string]$ApplicationObjectId,

    [Parameter(Mandatory = $true)]
    [string]$AppBaseUrl,

    [string]$FileHandlerId,

    [string]$DisplayName = "BPMN Diagram"
)

$ErrorActionPreference = "Stop"

$normalizedBaseUrl = $AppBaseUrl.TrimEnd("/")
$idFile = Join-Path $PSScriptRoot "..\infra\file-handler-id.txt"

if (-not $FileHandlerId) {
    if (Test-Path -LiteralPath $idFile) {
        $FileHandlerId = (Get-Content -LiteralPath $idFile -Raw).Trim()
    }
    else {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $idFile) | Out-Null
        $FileHandlerId = [guid]::NewGuid().ToString()
        Set-Content -LiteralPath $idFile -Value $FileHandlerId -Encoding UTF8
    }
}

$fileIcon = @{ svg = "$normalizedBaseUrl/assets/bpmn-file.svg" } | ConvertTo-Json -Compress
$appIcon = @{ svg = "$normalizedBaseUrl/assets/bpmn-app.svg" } | ConvertTo-Json -Compress

$actions = @(
    @{
        type = "open"
        url = "$normalizedBaseUrl/filehandler/open"
        availableOn = @{
            file = @{ extensions = @(".bpmn") }
            web = @{}
        }
    },
    @{
        type = "preview"
        url = "$normalizedBaseUrl/filehandler/preview"
        availableOn = @{
            file = @{ extensions = @(".bpmn") }
            web = @{}
        }
    }
) | ConvertTo-Json -Depth 20 -Compress

$fileHandler = @{
    id = $FileHandlerId
    type = "FileHandler"
    properties = @(
        @{ key = "version"; value = "2" },
        @{ key = "fileTypeDisplayName"; value = $DisplayName },
        @{ key = "actionMenuDisplayName"; value = "BPMN" },
        @{ key = "fileTypeIcon"; value = $fileIcon },
        @{ key = "appIcon"; value = $appIcon },
        @{ key = "actions"; value = $actions }
    )
}

$current = & az rest --method GET --uri "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId/`?`$select=addIns" -o json | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) {
    throw "Could not read application addIns."
}

$addIns = @()
if ($current.addIns) {
    $addIns += @($current.addIns | Where-Object { $_.id -ne $FileHandlerId })
}
$addIns += $fileHandler

$body = @{ addIns = $addIns } | ConvertTo-Json -Depth 30 -Compress
$bodyFile = New-TemporaryFile
Set-Content -LiteralPath $bodyFile -Value $body -Encoding UTF8

try {
    & az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId" --headers "Content-Type=application/json" --body "@$bodyFile" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "Could not update application addIns."
    }
}
finally {
    Remove-Item -LiteralPath $bodyFile -Force
}

[pscustomobject]@{
    fileHandlerId = $FileHandlerId
    appBaseUrl = $normalizedBaseUrl
    actions = @("open", "preview")
    extension = ".bpmn"
} | ConvertTo-Json -Depth 5

