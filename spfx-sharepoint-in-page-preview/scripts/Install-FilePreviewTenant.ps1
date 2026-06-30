param(
  [Parameter(Mandatory = $true)]
  [string]$TenantRootUrl,

  [Parameter(Mandatory = $true)]
  [string]$PnPClientId,

  [string]$ConfigSiteUrl = "",

  [string]$AdminPageName = "FilePreviewAdmin.aspx",

  [string]$PackagePath = ".\sharepoint\solution\spfx-bpmn-command-set.sppkg",

  [switch]$UploadPackage,

  [switch]$CreateAdminPage
)

$ErrorActionPreference = "Stop"

$AppVersion = "1.5.21"
$SolutionPackageName = "spfx-bpmn-command-set.sppkg"
$CommandSetComponentId = "c3e13f04-c3e1-4b55-8fd5-d7557cd15752"
$AdminApplicationCustomizerComponentId = "6da1de09-f3e8-4d81-a60b-0bf2c8c65be4"
$AdminWebPartComponentId = "58ccbc3f-a8dd-40ed-b4f3-c4a647338da8"
$ConfigListTitle = "M365 File Preview Settings"
$ConfigFieldName = "BpfConfigJson"
$ConfigItemTitle = "Configuration"
$AdminPageTitle = "File Preview Admin"
$IconAssetsSourcePath = Join-Path $PSScriptRoot "..\sharepoint\assets\file-handler-icons"
$IconAssetsSiteRelativeFolder = "SiteAssets/M365FilePreviewIcons"

$DefaultSettingsJson = @'
{"appBaseUrl":"","extensions":[{"displayName":"BPMN process diagram","enabled":true,"extension":".bpmn","mode":"modeler","renderer":"bpmn-js"},{"displayName":"JT 3D model","enabled":false,"extension":".jt","mode":"viewer","renderer":"coming-soon"},{"displayName":"diagrams.net drawing","enabled":false,"extension":".drawio","mode":"modeler","renderer":"diagrams-net-embed"},{"displayName":"Mermaid diagram","enabled":false,"extension":".mmd","mode":"viewer","renderer":"mermaid-js"},{"displayName":"Mermaid diagram (.mermaid)","enabled":false,"extension":".mermaid","mode":"viewer","renderer":"mermaid-js"},{"displayName":"IFC building model","enabled":false,"extension":".ifc","mode":"viewer","renderer":"web-ifc"},{"displayName":"STEP CAD model (.step)","enabled":false,"extension":".step","mode":"viewer","renderer":"occt-step"},{"displayName":"STEP CAD model (.stp)","enabled":false,"extension":".stp","mode":"viewer","renderer":"occt-step"}],"fileHandlerEnabled":false,"license":{"declaredUserCount":20,"freeUserLimit":20,"key":"","tier":"Free"},"schemaVersion":1}
'@

function Ensure-Module {
  if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Install-Module PnP.PowerShell -Scope CurrentUser -Force
  }
}

function Ensure-ConfigList {
  param(
    [string]$TargetSiteUrl
  )

  Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $PnPClientId

  $list = Get-PnPList -Identity $ConfigListTitle -ErrorAction SilentlyContinue
  if ($null -eq $list) {
    New-PnPList -Title $ConfigListTitle -Template GenericList -OnQuickLaunch:$false | Out-Null
  }

  $field = Get-PnPField -List $ConfigListTitle -Identity $ConfigFieldName -ErrorAction SilentlyContinue
  if ($null -eq $field) {
    Add-PnPField -List $ConfigListTitle -DisplayName $ConfigFieldName -InternalName $ConfigFieldName -Type Note -AddToDefaultView:$false | Out-Null
  }

  $existingConfigItems = @(Get-PnPListItem -List $ConfigListTitle -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$ConfigItemTitle</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>")
  $existingConfig = $existingConfigItems | Select-Object -First 1
  if ($null -eq $existingConfig) {
    $configValues = @{ Title = $ConfigItemTitle }
    $configValues[$ConfigFieldName] = $DefaultSettingsJson
    Add-PnPListItem -List $ConfigListTitle -Values $configValues | Out-Null
  } else {
    Repair-PreviewSettingsConfig -ConfigItem $existingConfig
  }
}

function Ensure-FileHandlerIcons {
  param(
    [string]$TargetSiteUrl
  )

  Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $PnPClientId

  if (-not (Test-Path -LiteralPath $IconAssetsSourcePath)) {
    Write-Warning "File handler icon source folder was not found: $IconAssetsSourcePath"
    return
  }

  Resolve-PnPFolder -SiteRelativePath $IconAssetsSiteRelativeFolder | Out-Null
  Get-ChildItem -LiteralPath $IconAssetsSourcePath -File | ForEach-Object {
    Add-PnPFile -Path $_.FullName -Folder $IconAssetsSiteRelativeFolder -Values @{ Title = $_.Name } | Out-Null
  }

  Write-Host "Uploaded file handler icons to $TargetSiteUrl/$IconAssetsSiteRelativeFolder"
}

function Repair-PreviewSettingsConfig {
  param(
    [object]$ConfigItem
  )

  $rawConfig = [string](Get-ListItemFieldValue -Item $ConfigItem -FieldName $ConfigFieldName)
  if ([string]::IsNullOrWhiteSpace($rawConfig)) {
    Set-PnPListItem -List $ConfigListTitle -Identity $ConfigItem.Id -Values @{ $ConfigFieldName = $DefaultSettingsJson } | Out-Null
    return
  }

  try {
    $settings = $rawConfig | ConvertFrom-Json
  } catch {
    Write-Warning "Existing file preview settings JSON is invalid. Leaving it unchanged. Details: $($_.Exception.Message)"
    return
  }

  $changed = $false
  # Required: must exist and be enabled. Do not change enabled state for optional extensions
  # (diagrams.net transfers file XML to an external service — admin must opt-in explicitly).
  $requiredExtensions = @(
    @{ displayName = "BPMN process diagram"; enabled = $true; extension = ".bpmn"; mode = "modeler"; renderer = "bpmn-js" }
  )
  $optionalExtensions = @(
    @{ displayName = "diagrams.net drawing"; enabled = $false; extension = ".drawio"; mode = "modeler"; renderer = "diagrams-net-embed" },
    @{ displayName = "Mermaid diagram"; enabled = $false; extension = ".mmd"; mode = "viewer"; renderer = "mermaid-js" },
    @{ displayName = "Mermaid diagram (.mermaid)"; enabled = $false; extension = ".mermaid"; mode = "viewer"; renderer = "mermaid-js" },
    @{ displayName = "IFC building model"; enabled = $false; extension = ".ifc"; mode = "viewer"; renderer = "web-ifc" },
    @{ displayName = "STEP CAD model (.step)"; enabled = $false; extension = ".step"; mode = "viewer"; renderer = "occt-step" },
    @{ displayName = "STEP CAD model (.stp)"; enabled = $false; extension = ".stp"; mode = "viewer"; renderer = "occt-step" }
  )

  if ($null -eq $settings.extensions) {
    $settings | Add-Member -MemberType NoteProperty -Name "extensions" -Value @() -Force
    $changed = $true
  }

  foreach ($requiredExtension in $requiredExtensions) {
    $extensionSettings = $settings.extensions | Where-Object { [string]$_.extension -eq $requiredExtension.extension } | Select-Object -First 1
    if ($null -eq $extensionSettings) {
      $settings.extensions = @($settings.extensions) + [pscustomobject]$requiredExtension
      $changed = $true
      continue
    }

    if ($extensionSettings.enabled -ne $true) {
      $extensionSettings.enabled = $true
      $changed = $true
    }
    if ([string]$extensionSettings.mode -ne [string]$requiredExtension.mode) {
      $extensionSettings.mode = $requiredExtension.mode
      $changed = $true
    }
    if ([string]$extensionSettings.renderer -ne [string]$requiredExtension.renderer) {
      $extensionSettings.renderer = $requiredExtension.renderer
      $changed = $true
    }
  }

  foreach ($optionalExtension in $optionalExtensions) {
    $extensionSettings = $settings.extensions | Where-Object { [string]$_.extension -eq $optionalExtension.extension } | Select-Object -First 1
    if ($null -eq $extensionSettings) {
      $settings.extensions = @($settings.extensions) + [pscustomobject]$optionalExtension
      $changed = $true
    }
    # If already present, respect the admin's enabled/disabled choice — do not override.
  }

  if ($changed) {
    $updatedConfig = $settings | ConvertTo-Json -Depth 20 -Compress
    Set-PnPListItem -List $ConfigListTitle -Identity $ConfigItem.Id -Values @{ $ConfigFieldName = $updatedConfig } | Out-Null
    Write-Host "Updated file preview settings configuration."
  }
}

function Normalize-FieldToken {
  param(
    [string]$Value
  )

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return ""
  }

  return (($Value -replace "[^A-Za-z0-9]", "").ToLowerInvariant())
}

function Resolve-ListFieldInternalName {
  param(
    [object[]]$Fields,
    [string[]]$Candidates,
    [string]$LogicalName,
    [switch]$Required
  )

  foreach ($candidate in $Candidates) {
    $normalizedCandidate = Normalize-FieldToken -Value $candidate
    foreach ($field in $Fields) {
      $fieldNames = @($field.InternalName, $field.StaticName, $field.Title)
      foreach ($fieldName in $fieldNames) {
        if ((Normalize-FieldToken -Value ([string]$fieldName)) -eq $normalizedCandidate) {
          return $field.InternalName
        }
      }
    }
  }

  if ($Required) {
    $availableFields = ($Fields | ForEach-Object { "$($_.Title) [$($_.InternalName)]" }) -join ", "
    throw "Could not find required Tenant Wide Extensions field '$LogicalName'. Available fields: $availableFields"
  }

  return $null
}

function Add-ListItemValue {
  param(
    [hashtable]$Values,
    [string]$FieldName,
    [object]$Value
  )

  if (-not [string]::IsNullOrWhiteSpace($FieldName)) {
    $Values[$FieldName] = $Value
  }
}

function Get-ListItemFieldValue {
  param(
    [object]$Item,
    [string]$FieldName
  )

  if ($null -eq $Item -or [string]::IsNullOrWhiteSpace($FieldName)) {
    return $null
  }

  try {
    return $Item[$FieldName]
  } catch {
    if ($null -ne $Item.FieldValues -and $Item.FieldValues.ContainsKey($FieldName)) {
      return $Item.FieldValues[$FieldName]
    }
  }

  return $null
}

# NOTE: With skipFeatureDeployment:true, the sppkg auto-registers components via
# ClientSideInstance.xml when deployed to the app catalog. Use Ensure-TenantWideCommand
# only to update properties (e.g., configSiteUrl) on existing registrations, or if
# the automatic registration from ClientSideInstance.xml did not apply as expected.
function Ensure-TenantWideCommand {
  param(
    [string]$AppCatalogSiteUrl,
    [string]$TargetConfigSiteUrl
  )

  Connect-PnPOnline -Url $AppCatalogSiteUrl -Interactive -ClientId $PnPClientId

  $commandSetProperties = @{ configSiteUrl = $TargetConfigSiteUrl; showSettingsCommand = $false } | ConvertTo-Json -Compress
  $adminLauncherProperties = @{ configSiteUrl = $TargetConfigSiteUrl; showOnAllSites = $false } | ConvertTo-Json -Compress
  $tenantWideListTitle = "Tenant Wide Extensions"
  $tenantWideFields = Get-PnPField -List $tenantWideListTitle
  $componentIdField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("ComponentId", "TenantWideExtensionComponentId", "ClientSideComponentId", "ClientSideComponentIdOverride") -LogicalName "ComponentId" -Required
  $propertiesField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("ComponentProperties", "TenantWideExtensionComponentProperties", "ClientSideComponentProperties", "Properties") -LogicalName "ComponentProperties" -Required
  $locationField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("Location", "TenantWideExtensionLocation") -LogicalName "Location" -Required
  $listTemplateField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("ListTemplate", "TenantWideExtensionListTemplate", "ListTemplateId") -LogicalName "ListTemplate"
  $sequenceField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("Sequence", "TenantWideExtensionSequence") -LogicalName "Sequence"
  $disabledField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("Disabled", "TenantWideExtensionDisabled") -LogicalName "Disabled"

  $tenantWideItems = @(Get-PnPListItem -List $tenantWideListTitle -PageSize 1000)
  $commandSetItems = @($tenantWideItems | Where-Object { [string](Get-ListItemFieldValue -Item $_ -FieldName $componentIdField) -eq $CommandSetComponentId })
  $adminLauncherItems = @($tenantWideItems | Where-Object { [string](Get-ListItemFieldValue -Item $_ -FieldName $componentIdField) -eq $AdminApplicationCustomizerComponentId })

  Ensure-TenantWideCommandLocation `
    -TenantWideListTitle $tenantWideListTitle `
    -TenantWideItems $commandSetItems `
    -ComponentId $CommandSetComponentId `
    -ComponentIdField $componentIdField `
    -PropertiesField $propertiesField `
    -LocationField $locationField `
    -ListTemplateField $listTemplateField `
    -SequenceField $sequenceField `
    -DisabledField $disabledField `
    -ComponentProperties $commandSetProperties `
    -Location "ClientSideExtension.ListViewCommandSet.CommandBar" `
    -Title "M365 File Preview Command Set - Command Bar"

  Ensure-TenantWideCommandLocation `
    -TenantWideListTitle $tenantWideListTitle `
    -TenantWideItems $commandSetItems `
    -ComponentId $CommandSetComponentId `
    -ComponentIdField $componentIdField `
    -PropertiesField $propertiesField `
    -LocationField $locationField `
    -ListTemplateField $listTemplateField `
    -SequenceField $sequenceField `
    -DisabledField $disabledField `
    -ComponentProperties $commandSetProperties `
    -Location "ClientSideExtension.ListViewCommandSet.ContextMenu" `
    -Title "M365 File Preview Command Set - Context Menu"

  Ensure-TenantWideCommandLocation `
    -TenantWideListTitle $tenantWideListTitle `
    -TenantWideItems $adminLauncherItems `
    -ComponentId $AdminApplicationCustomizerComponentId `
    -ComponentIdField $componentIdField `
    -PropertiesField $propertiesField `
    -LocationField $locationField `
    -ListTemplateField $null `
    -SequenceField $sequenceField `
    -DisabledField $disabledField `
    -ComponentProperties $adminLauncherProperties `
    -Location "ClientSideExtension.ApplicationCustomizer" `
    -Title "M365 File Preview Admin Launcher"

  Write-TenantWideRegistrationSummary `
    -TenantWideListTitle $tenantWideListTitle `
    -ComponentIdField $componentIdField `
    -PropertiesField $propertiesField `
    -LocationField $locationField `
    -ListTemplateField $listTemplateField `
    -DisabledField $disabledField
}

function Ensure-TenantWideCommandLocation {
  param(
    [string]$TenantWideListTitle,
    [object[]]$TenantWideItems,
    [string]$ComponentId,
    [string]$ComponentIdField,
    [string]$PropertiesField,
    [string]$LocationField,
    [string]$ListTemplateField,
    [string]$SequenceField,
    [string]$DisabledField,
    [string]$ComponentProperties,
    [string]$Location,
    [string]$Title
  )

  $matchingItems = @($TenantWideItems | Where-Object { [string](Get-ListItemFieldValue -Item $_ -FieldName $LocationField) -eq $Location })
  $matchingItem = $matchingItems | Select-Object -First 1

  $tenantWideValues = @{ Title = $Title }
  Add-ListItemValue -Values $tenantWideValues -FieldName $ComponentIdField -Value $ComponentId
  Add-ListItemValue -Values $tenantWideValues -FieldName $PropertiesField -Value $ComponentProperties
  Add-ListItemValue -Values $tenantWideValues -FieldName $LocationField -Value $Location
  Add-ListItemValue -Values $tenantWideValues -FieldName $ListTemplateField -Value 101
  Add-ListItemValue -Values $tenantWideValues -FieldName $SequenceField -Value 1
  Add-ListItemValue -Values $tenantWideValues -FieldName $DisabledField -Value $false

  if ($null -ne $matchingItem) {
    Set-PnPListItem -List $TenantWideListTitle -Identity $matchingItem.Id -Values $tenantWideValues | Out-Null
  } else {
    Add-PnPListItem -List $TenantWideListTitle -Values $tenantWideValues | Out-Null
  }
}

function Write-TenantWideRegistrationSummary {
  param(
    [string]$TenantWideListTitle,
    [string]$ComponentIdField,
    [string]$PropertiesField,
    [string]$LocationField,
    [string]$ListTemplateField,
    [string]$DisabledField
  )

  Write-Host "Tenant Wide Extensions registrations:"
  Get-PnPListItem -List $TenantWideListTitle -PageSize 1000 |
    Where-Object {
      [string](Get-ListItemFieldValue -Item $_ -FieldName $ComponentIdField) -eq $CommandSetComponentId -or
      [string](Get-ListItemFieldValue -Item $_ -FieldName $ComponentIdField) -eq $AdminApplicationCustomizerComponentId
    } |
    ForEach-Object {
      $listTemplate = if ($ListTemplateField) { [string](Get-ListItemFieldValue -Item $_ -FieldName $ListTemplateField) } else { "" }
      $disabled = if ($DisabledField) { [string](Get-ListItemFieldValue -Item $_ -FieldName $DisabledField) } else { "" }
      Write-Host ("  Id={0}; Title={1}; ComponentId={2}; Location={3}; ListTemplate={4}; Disabled={5}; Properties={6}" -f $_.Id, (Get-ListItemFieldValue -Item $_ -FieldName "Title"), (Get-ListItemFieldValue -Item $_ -FieldName $ComponentIdField), (Get-ListItemFieldValue -Item $_ -FieldName $LocationField), $listTemplate, $disabled, (Get-ListItemFieldValue -Item $_ -FieldName $PropertiesField))
    }
}

function Get-AdminPageCreationName {
  if ($AdminPageName.EndsWith(".aspx", [System.StringComparison]::OrdinalIgnoreCase)) {
    return $AdminPageName.Substring(0, $AdminPageName.Length - 5)
  }

  return $AdminPageName
}

function Remove-AdminPageIfBroken {
  param(
    [string]$AdminSiteUrl
  )

  try {
    $page = Get-PnPPage -Identity $AdminPageName -ErrorAction Stop
    if ($null -ne $page -and $null -ne $page.Controls) {
      return
    }
  } catch {
    if ($_.Exception.Message -notmatch "NoComponentId|Object reference") {
      return
    }
  }

  try {
    $web = Get-PnPWeb -Includes ServerRelativeUrl
    $webRelativeUrl = [string]$web.ServerRelativeUrl
    $pageUrl = if ([string]::IsNullOrWhiteSpace($webRelativeUrl) -or $webRelativeUrl -eq "/") {
      "/SitePages/$AdminPageName"
    } else {
      "$webRelativeUrl/SitePages/$AdminPageName"
    }

    Remove-PnPFile -ServerRelativeUrl $pageUrl -Force -Recycle -ErrorAction Stop
    Write-Warning "Removed a broken File Preview Admin page so it can be recreated cleanly: $AdminSiteUrl/SitePages/$AdminPageName"
  } catch {
    Write-Warning "Could not remove the broken File Preview Admin page automatically. Delete $AdminSiteUrl/SitePages/$AdminPageName manually, then rerun this script. Details: $($_.Exception.Message)"
  }
}

function Ensure-AdminPage {
  param(
    [string]$AdminSiteUrl,
    [string]$TargetConfigSiteUrl
  )

  Connect-PnPOnline -Url $AdminSiteUrl -Interactive -ClientId $PnPClientId

  Remove-AdminPageIfBroken -AdminSiteUrl $AdminSiteUrl

  $adminPage = Get-PnPPage -Identity $AdminPageName -ErrorAction SilentlyContinue
  if ($null -eq $adminPage) {
    $adminPageCreationName = Get-AdminPageCreationName
    try {
      Add-PnPPage -Name $adminPageCreationName -Title $AdminPageTitle -LayoutType Article | Out-Null
    } catch {
      Write-Warning "Could not create the File Preview Admin page. The command set and configuration were installed, but the admin page must be created later. Details: $($_.Exception.Message)"
      return $false
    }
  }

  $availableComponent = Get-PnPAvailableClientSideComponents -Page $AdminPageName | Where-Object { $_.Id -eq $AdminWebPartComponentId } | Select-Object -First 1
  if ($null -eq $availableComponent) {
    Write-Warning "The File Preview Admin web part is not available yet. Wait a few minutes after package deployment, then rerun this script."
    return $false
  }

  $page = Get-PnPPage -Identity $AdminPageName
  $existingWebPart = $page.Controls | Where-Object { $_.WebPartId -eq $AdminWebPartComponentId } | Select-Object -First 1
  if ($null -eq $existingWebPart) {
    Add-PnPPageWebPart -Page $AdminPageName -Component $availableComponent -Section 1 -Column 1 -WebPartProperties @{ configSiteUrl = $TargetConfigSiteUrl } | Out-Null
  }

  Set-PnPPage -Identity $AdminPageName -Title $AdminPageTitle -Publish | Out-Null
  return $true
}

Ensure-Module

Connect-PnPOnline -Url $TenantRootUrl -Interactive -ClientId $PnPClientId
$tenantSettings = Invoke-PnPSPRestMethod -Url "$TenantRootUrl/_api/SP_TenantSettings_Current" -Method Get
$AppCatalogSiteUrl = $tenantSettings.CorporateCatalogUrl
if ([string]::IsNullOrWhiteSpace($AppCatalogSiteUrl)) {
  throw "Tenant app catalog URL could not be resolved. Create the tenant App Catalog first, wait a few minutes, then run this script again."
}

if ([string]::IsNullOrWhiteSpace($ConfigSiteUrl)) {
  $ConfigSiteUrl = $AppCatalogSiteUrl
}

if ($UploadPackage) {
  Connect-PnPOnline -Url $AppCatalogSiteUrl -Interactive -ClientId $PnPClientId
  Add-PnPApp -Path $PackagePath -Scope Tenant -Overwrite -Publish | Out-Null
  Write-Host "Uploaded package $SolutionPackageName version $AppVersion."
}

Ensure-ConfigList -TargetSiteUrl $ConfigSiteUrl
Ensure-FileHandlerIcons -TargetSiteUrl $ConfigSiteUrl
Ensure-TenantWideCommand -AppCatalogSiteUrl $AppCatalogSiteUrl -TargetConfigSiteUrl $ConfigSiteUrl
Remove-AdminPageIfBroken -AdminSiteUrl $AppCatalogSiteUrl
$adminPageReady = if ($CreateAdminPage) {
  Ensure-AdminPage -AdminSiteUrl $AppCatalogSiteUrl -TargetConfigSiteUrl $ConfigSiteUrl
} else {
  $false
}

Write-Host "File Preview tenant setup complete."
Write-Host "Version:" $AppVersion
Write-Host "App Catalog:" $AppCatalogSiteUrl
Write-Host "Configuration site:" $ConfigSiteUrl
Write-Host "File handler icon base URL:" "$ConfigSiteUrl/$IconAssetsSiteRelativeFolder"
if ($adminPageReady) {
  Write-Host "Admin page:" "$AppCatalogSiteUrl/SitePages/$AdminPageName"
} else {
  Write-Host "Admin launcher:" "$TenantRootUrl/?m365FilePreviewAdmin=1"
}

