# Microsoft 365 File Preview SPFx Package

This SPFx package provides the SharePoint-first preview experience for engineering, process, and architecture files.

The package includes:

- An `Open preview` command when one enabled and available file extension is selected.
- A same-page SharePoint dialog for supported files.
- A bundled `bpmn-js` renderer for BPMN files.
- Optional `.drawio` preview through the diagrams.net embedded renderer.
- SharePoint REST file read/save behavior with no Azure dependency for SharePoint preview.
- A tenant root `File Preview Admin` launcher for administrators.
- An optional `Tenant file preview settings` command that can be enabled through component properties.
- A `File Preview Admin` web part for a central admin page or single part app page.
- Central settings in the `M365 File Preview Settings` SharePoint list.
- Optional Microsoft 365 File Handler registration for OneDrive/native launch scenarios.

## Configuration Scope

Tenant defaults live in one central SharePoint configuration site. The AppSource-friendly default is the tenant root site unless a customer sets a specific `configSiteUrl`.

Default behavior:

- The `.sppkg` declaratively deploys the document-library command set through SPFx tenant-wide deployment.
- The `.sppkg` declaratively deploys an application customizer that shows `File Preview Admin` on the tenant root site for administrators.
- Opening the tenant root launcher initializes missing configuration assets before opening settings.
- Existing configuration is preserved on upgrades; initialization does not overwrite an existing configuration item.
- The document-library settings command is hidden by default so first-run configuration is not tied to a library.
- The `File Preview Admin` web part can be placed on a central SharePoint admin page or used as a single part app page.
- The library command set reads the same central settings list.

The tenant root launcher is the AppSource-friendly first-run path. The settings command and admin web part are secondary entry points and do not store settings in a document library.

## Default Configuration

```json
{
  "schemaVersion": 1,
  "appBaseUrl": "",
  "fileHandlerEnabled": false,
  "license": {
    "tier": "Free",
    "declaredUserCount": 20,
    "freeUserLimit": 20,
    "key": ""
  },
  "extensions": [
    {
      "extension": ".bpmn",
      "displayName": "BPMN process diagram",
      "enabled": true,
      "renderer": "bpmn-js",
      "mode": "modeler"
    },
    {
      "extension": ".drawio",
      "displayName": "diagrams.net drawing",
      "enabled": false,
      "renderer": "diagrams-net-embed",
      "mode": "viewer"
    },
    {
      "extension": ".jt",
      "displayName": "JT 3D model",
      "enabled": false,
      "renderer": "coming-soon",
      "mode": "viewer"
    },
    {
      "extension": ".step",
      "displayName": "STEP CAD model",
      "enabled": false,
      "renderer": "coming-soon",
      "mode": "viewer"
    }
  ]
}
```

The free tier is intended for consulting companies or internal teams with 20 total users or less. Paid tiers use an admin-entered license key in the settings panel.

## Runtime Behavior

BPMN SharePoint command path:

```text
SharePoint document library -> SPFx command set -> SPFx dialog -> bundled bpmn-js renderer -> SharePoint REST file read/save
```

No Azure web app is required for the SharePoint same-page modal renderer.

Draw.io SharePoint command path:

```text
SharePoint document library -> SPFx command set -> SPFx dialog -> https://embed.diagrams.net iframe -> SharePoint REST file read/download
```

The `.drawio` renderer is disabled by default because it uses an external embedded runtime. SharePoint remains the storage location, but the drawing XML is loaded into `https://embed.diagrams.net` in the user's browser session. Admins must explicitly enable `.drawio` in `File Preview Admin`.

## Optional Scripts

The SharePoint preview experience does not require PowerShell. The admin UI includes only an optional Microsoft 365 File Handler registration script.

The optional script includes:

- Current tenant ID when available from SharePoint page context.
- Enabled extension list.
- File Handler endpoint URL.
- Azure CLI commands to create/update the Entra application `addIns` File Handler metadata.

Native Microsoft 365 File Handler registration is optional. It requires an HTTPS handler endpoint that can receive Microsoft 365 File Handler launch requests; SPFx alone cannot receive those POST launch requests.

The repository still contains `scripts/Install-FilePreviewTenant.ps1` as an advanced troubleshooting tool for direct tenant administration, but it is not part of the AppSource first-run path.

## Local Debug

The local SPFx dev server serves the manifest from:

```text
https://localhost:4321/temp/build/manifests.js
```

Use this SharePoint debug URL to test the command set and settings command:

```text
https://evolvegs.sharepoint.com/sites/EvolveGSAIHub/Shared%20Documents/Forms/AllItems.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&noredir=true&loadSPFX=true&customActions=%7B%22c3e13f04-c3e1-4b55-8fd5-d7557cd15752%22%3A%7B%22location%22%3A%22ClientSideExtension.ListViewCommandSet.CommandBar%22%2C%22properties%22%3A%7B%22showSettingsCommand%22%3Atrue%7D%7D%7D
```

When prompted by SharePoint, load debug scripts. Select one `.bpmn` file and use `Open preview`. To test `.drawio`, enable it first in `File Preview Admin`, then select one `.drawio` file and use `Open preview`.

## Build

```powershell
cd spfx-bpmn-command-set
npm install
npm run build
```

The package is created at:

```text
spfx-bpmn-command-set/sharepoint/solution/spfx-bpmn-command-set.sppkg
```

## Deployment

Tenant app catalog deployment requires SharePoint app catalog permissions.

```powershell
m365 spo app add --filePath "spfx-bpmn-command-set/sharepoint/solution/spfx-bpmn-command-set.sppkg" --overwrite --appCatalogScope tenant
m365 spo app deploy --name "spfx-bpmn-command-set.sppkg" --appCatalogScope tenant --skipFeatureDeployment
```

After deployment:

1. Go to the tenant root site, for example `https://contoso.sharepoint.com`.
2. Use the `File Preview Admin` launcher.
3. Configure enabled extensions, licensing, and optional File Handler metadata.
4. Save settings.

The launcher initializes missing configuration assets automatically and preserves existing settings on upgrade.

## Upgrade

The current package version is `1.4.0.0`.

To upgrade:

1. Upload `spfx-bpmn-command-set/sharepoint/solution/spfx-bpmn-command-set.sppkg` to the tenant App Catalog.
2. Choose replace when SharePoint detects the existing package.
3. Confirm the App Catalog shows version `1.4.0.0`.
4. Deploy the package tenant-wide.
5. Open the tenant root site and use the `File Preview Admin` launcher.
6. Confirm settings still reflect the previous configuration.
7. Save settings only if you want to change the existing configuration.
