param(
  [Parameter(Mandatory = $true)]
  [string]$TenantRootUrl,

  [string]$ConfigSiteUrl = "",

  [string]$AdminPageName = "FilePreviewAdmin.aspx",

  [string]$PackagePath = ".\spfx-bpmn-command-set\sharepoint\solution\spfx-bpmn-command-set.sppkg",

  [switch]$UploadPackage
)

$ErrorActionPreference = "Stop"

$AppVersion = "1.4.14"
$ClientID = "a53c4564-ac7f-48c8-ab09-c447875beb17"
$SolutionPackageName = "spfx-bpmn-command-set.sppkg"
$CommandSetComponentId = "c3e13f04-c3e1-4b55-8fd5-d7557cd15752"
$AdminWebPartComponentId = "58ccbc3f-a8dd-40ed-b4f3-c4a647338da8"
$ConfigListTitle = "M365 File Preview Settings"
$ConfigFieldName = "BpfConfigJson"
$ConfigItemTitle = "Configuration"
$AdminPageTitle = "File Preview Admin"

$DefaultSettingsJson = @'
{"appBaseUrl":"","extensions":[{"displayName":"BPMN process diagram","enabled":true,"extension":".bpmn","mode":"modeler","renderer":"bpmn-js"},{"displayName":"JT 3D model","enabled":false,"extension":".jt","mode":"viewer","renderer":"coming-soon"},{"displayName":"diagrams.net drawing","enabled":false,"extension":".drawio","mode":"modeler","renderer":"diagrams-net-embed"},{"displayName":"STEP CAD model","enabled":false,"extension":".step","mode":"viewer","renderer":"coming-soon"}],"fileHandlerEnabled":false,"license":{"declaredUserCount":20,"freeUserLimit":20,"key":"","tier":"Free"},"schemaVersion":1}
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

  Connect-PnPOnline -Url $TargetSiteUrl -Interactive -ClientId $ClientID

  $list = Get-PnPList -Identity $ConfigListTitle -ErrorAction SilentlyContinue
  if ($null -eq $list) {
    New-PnPList -Title $ConfigListTitle -Template GenericList -OnQuickLaunch:$false | Out-Null
  }

  $field = Get-PnPField -List $ConfigListTitle -Identity $ConfigFieldName -ErrorAction SilentlyContinue
  if ($null -eq $field) {
    Add-PnPField -List $ConfigListTitle -DisplayName $ConfigFieldName -InternalName $ConfigFieldName -Type Note -AddToDefaultView:$false | Out-Null
  }

  $existingConfig = Get-PnPListItem -List $ConfigListTitle -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$ConfigItemTitle</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>"
  if ($existingConfig.Count -eq 0) {
    $configValues = @{ Title = $ConfigItemTitle }
    $configValues[$ConfigFieldName] = $DefaultSettingsJson
    Add-PnPListItem -List $ConfigListTitle -Values $configValues | Out-Null
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

function Ensure-TenantWideCommand {
  param(
    [string]$AppCatalogSiteUrl,
    [string]$TargetConfigSiteUrl
  )

  Connect-PnPOnline -Url $AppCatalogSiteUrl -Interactive -ClientId $ClientID

  $componentProperties = @{ configSiteUrl = $TargetConfigSiteUrl; showSettingsCommand = $false } | ConvertTo-Json -Compress
  $tenantWideListTitle = "Tenant Wide Extensions"
  $tenantWideFields = Get-PnPField -List $tenantWideListTitle
  $componentIdField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("ComponentId", "TenantWideExtensionComponentId", "ClientSideComponentId", "ClientSideComponentIdOverride") -LogicalName "ComponentId" -Required
  $propertiesField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("ComponentProperties", "TenantWideExtensionComponentProperties", "ClientSideComponentProperties", "Properties") -LogicalName "ComponentProperties" -Required
  $locationField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("Location", "TenantWideExtensionLocation") -LogicalName "Location" -Required
  $listTemplateField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("ListTemplate", "TenantWideExtensionListTemplate", "ListTemplateId") -LogicalName "ListTemplate"
  $sequenceField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("Sequence", "TenantWideExtensionSequence") -LogicalName "Sequence"
  $disabledField = Resolve-ListFieldInternalName -Fields $tenantWideFields -Candidates @("Disabled", "TenantWideExtensionDisabled") -LogicalName "Disabled"

  $tenantWideItems = Get-PnPListItem -List $tenantWideListTitle -PageSize 1000 | Where-Object { [string]$_[$componentIdField] -eq $CommandSetComponentId }

  $tenantWideValues = @{ Title = "M365 File Preview Command Set" }
  Add-ListItemValue -Values $tenantWideValues -FieldName $componentIdField -Value $CommandSetComponentId
  Add-ListItemValue -Values $tenantWideValues -FieldName $propertiesField -Value $componentProperties
  Add-ListItemValue -Values $tenantWideValues -FieldName $locationField -Value "ClientSideExtension.ListViewCommandSet.CommandBar"
  Add-ListItemValue -Values $tenantWideValues -FieldName $listTemplateField -Value 101
  Add-ListItemValue -Values $tenantWideValues -FieldName $sequenceField -Value 1
  Add-ListItemValue -Values $tenantWideValues -FieldName $disabledField -Value $false

  if ($tenantWideItems.Count -gt 0) {
    Set-PnPListItem -List $tenantWideListTitle -Identity $tenantWideItems[0].Id -Values $tenantWideValues | Out-Null
  } else {
    Add-PnPListItem -List $tenantWideListTitle -Values $tenantWideValues | Out-Null
  }
}

function Ensure-AdminPage {
  param(
    [string]$AdminSiteUrl,
    [string]$TargetConfigSiteUrl
  )

  Connect-PnPOnline -Url $AdminSiteUrl -Interactive -ClientId $ClientID

  $adminPage = Get-PnPPage -Identity $AdminPageName -ErrorAction SilentlyContinue
  if ($null -eq $adminPage) {
    Add-PnPPage -Name $AdminPageName -Title $AdminPageTitle -LayoutType SingleWebPartAppPage | Out-Null
  }

  $availableComponent = Get-PnPAvailableClientSideComponents -Page $AdminPageName | Where-Object { $_.Id -eq $AdminWebPartComponentId } | Select-Object -First 1
  if ($null -eq $availableComponent) {
    Write-Warning "The File Preview Admin web part is not available yet. Wait a few minutes after package deployment, then rerun this script."
    return
  }

  $page = Get-PnPPage -Identity $AdminPageName
  $existingWebPart = $page.Controls | Where-Object { $_.WebPartId -eq $AdminWebPartComponentId } | Select-Object -First 1
  if ($null -eq $existingWebPart) {
    Add-PnPPageWebPart -Page $AdminPageName -Component $availableComponent -Section 1 -Column 1 -WebPartProperties @{ configSiteUrl = $TargetConfigSiteUrl } | Out-Null
  }

  Set-PnPPage -Identity $AdminPageName -Title $AdminPageTitle -Publish | Out-Null
}

Ensure-Module

Connect-PnPOnline -Url $TenantRootUrl -Interactive -ClientId $ClientID
$tenantSettings = Invoke-PnPSPRestMethod -Url "$TenantRootUrl/_api/SP_TenantSettings_Current" -Method Get
$AppCatalogSiteUrl = $tenantSettings.CorporateCatalogUrl
if ([string]::IsNullOrWhiteSpace($AppCatalogSiteUrl)) {
  throw "Tenant app catalog URL could not be resolved. Create the tenant App Catalog first, wait a few minutes, then run this script again."
}

if ([string]::IsNullOrWhiteSpace($ConfigSiteUrl)) {
  $ConfigSiteUrl = $AppCatalogSiteUrl
}

if ($UploadPackage) {
  Connect-PnPOnline -Url $AppCatalogSiteUrl -Interactive -ClientId $ClientID
  Add-PnPApp -Path $PackagePath -Scope Tenant -Overwrite -Publish | Out-Null
  Write-Host "Uploaded package $SolutionPackageName version $AppVersion."
}

Ensure-ConfigList -TargetSiteUrl $ConfigSiteUrl
Ensure-TenantWideCommand -AppCatalogSiteUrl $AppCatalogSiteUrl -TargetConfigSiteUrl $ConfigSiteUrl
Ensure-AdminPage -AdminSiteUrl $AppCatalogSiteUrl -TargetConfigSiteUrl $ConfigSiteUrl

Write-Host "File Preview tenant setup complete."
Write-Host "Version:" $AppVersion
Write-Host "App Catalog:" $AppCatalogSiteUrl
Write-Host "Configuration site:" $ConfigSiteUrl
Write-Host "Admin page:" "$AppCatalogSiteUrl/SitePages/$AdminPageName"
