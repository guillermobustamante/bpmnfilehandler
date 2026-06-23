param(
    [Parameter(Mandatory = $true)]
    [string]$ApplicationObjectId,

    [Parameter(Mandatory = $true)]
    [string]$AppBaseUrl,

    [string]$IconBaseUrl = "",

    [string[]]$Extensions = @(".bpmn", ".drawio")
)

$ErrorActionPreference = "Stop"

$normalizedBaseUrl = $AppBaseUrl.TrimEnd("/")
$normalizedIconBaseUrl = if ([string]::IsNullOrWhiteSpace($IconBaseUrl)) { "$normalizedBaseUrl/assets" } else { $IconBaseUrl.TrimEnd("/") }
$idFile = Join-Path $PSScriptRoot "..\infra\file-handler-ids.json"
$legacyIdFile = Join-Path $PSScriptRoot "..\infra\file-handler-id.txt"

function Normalize-Extension {
    param([string]$Extension)

    $trimmed = $Extension.Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return ""
    }

    if ($trimmed.StartsWith(".")) {
        return $trimmed
    }

    return ".$trimmed"
}

function Get-HandlerIds {
    param([string[]]$RequestedExtensions)

    $ids = @{}
    if (Test-Path -LiteralPath $idFile) {
        $stored = Get-Content -LiteralPath $idFile -Raw | ConvertFrom-Json
        foreach ($property in $stored.PSObject.Properties) {
            $ids[$property.Name] = [string]$property.Value
        }
    }

    if (-not $ids.ContainsKey(".bpmn") -and (Test-Path -LiteralPath $legacyIdFile)) {
        $legacyId = (Get-Content -LiteralPath $legacyIdFile -Raw).Trim()
        if (-not [string]::IsNullOrWhiteSpace($legacyId)) {
            $ids[".bpmn"] = $legacyId
        }
    }

    $changed = $false
    foreach ($extension in $RequestedExtensions) {
        if (-not $ids.ContainsKey($extension)) {
            $ids[$extension] = [guid]::NewGuid().ToString()
            $changed = $true
        }
    }

    if ($changed -or -not (Test-Path -LiteralPath $idFile)) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $idFile) | Out-Null
        $ids | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $idFile -Encoding UTF8
    }

    return $ids
}

function Get-ExtensionSpec {
    param([string]$Extension)

    switch ($Extension) {
        ".bpmn" {
            return @{
                assetPrefix = "bpmn"
                actionMenuDisplayName = "Open BPMN"
                fileTypeDisplayName = "BPMN process diagram"
                openLabel = "Open BPMN"
                openMode = "modeler"
                previewMode = "viewer"
            }
        }
        ".drawio" {
            return @{
                assetPrefix = "drawio"
                actionMenuDisplayName = "Open DrawIO"
                fileTypeDisplayName = "DrawIO diagram"
                openLabel = "Open DrawIO"
                openMode = "modeler"
                previewMode = "viewer"
            }
        }
        default {
            throw "No File Handler renderer/icon mapping exists for $Extension. Supported: .bpmn, .drawio"
        }
    }
}

function New-IconJson {
    param(
        [string]$AssetPrefix,
        [string]$IconType
    )

    @{
        svg = "$normalizedIconBaseUrl/$AssetPrefix-$IconType.svg"
        png1x = "$normalizedIconBaseUrl/$AssetPrefix-$IconType-32.png"
        "png1.5x" = "$normalizedIconBaseUrl/$AssetPrefix-$IconType-48.png"
        png2x = "$normalizedIconBaseUrl/$AssetPrefix-$IconType-64.png"
    } | ConvertTo-Json -Compress
}

function New-FileHandler {
    param(
        [string]$Extension,
        [string]$FileHandlerId
    )

    $spec = Get-ExtensionSpec -Extension $Extension
    $encodedExtension = [uri]::EscapeDataString($Extension)
    $actions = @(
        @{
            type = "preview"
            url = "$normalizedBaseUrl/filehandler/preview?extension=$encodedExtension&mode=$($spec.previewMode)"
            availableOn = @{
                file = @{ extensions = @($Extension) }
                web = @{}
            }
        },
        @{
            type = "open"
            url = "$normalizedBaseUrl/filehandler/open?extension=$encodedExtension&mode=$($spec.openMode)"
            displayName = $spec.openLabel
            shortDisplayName = $spec.openLabel
            availableOn = @{
                file = @{ extensions = @($Extension) }
                web = @{}
            }
        }
    ) | ConvertTo-Json -Depth 20 -Compress

    return @{
        id = $FileHandlerId
        type = "FileHandler"
        properties = @(
            @{ key = "version"; value = "2" },
            @{ key = "fileTypeDisplayName"; value = $spec.fileTypeDisplayName },
            @{ key = "actionMenuDisplayName"; value = $spec.actionMenuDisplayName },
            @{ key = "fileTypeIcon"; value = (New-IconJson -AssetPrefix $spec.assetPrefix -IconType "file") },
            @{ key = "appIcon"; value = (New-IconJson -AssetPrefix $spec.assetPrefix -IconType "app") },
            @{ key = "actions"; value = $actions }
        )
    }
}

$normalizedExtensions = @($Extensions | ForEach-Object { Normalize-Extension -Extension $_ } | Where-Object { $_ } | Select-Object -Unique)
if ($normalizedExtensions.Count -eq 0) {
    throw "At least one extension must be supplied."
}

$handlerIds = Get-HandlerIds -RequestedExtensions $normalizedExtensions
$fileHandlers = @()
foreach ($extension in $normalizedExtensions) {
    $fileHandlers += New-FileHandler -Extension $extension -FileHandlerId $handlerIds[$extension]
}

$current = & az rest --method GET --uri "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId/`?`$select=addIns" -o json | ConvertFrom-Json
if ($LASTEXITCODE -ne 0) {
    throw "Could not read application addIns."
}

$addIns = @()
if ($current.addIns) {
    $addIns += @($current.addIns | Where-Object { $_.type -ne "FileHandler" })
}
$addIns += $fileHandlers

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
    fileHandlerIds = $handlerIds
    appBaseUrl = $normalizedBaseUrl
    actions = @("preview", "open")
    extensions = $normalizedExtensions
} | ConvertTo-Json -Depth 10
