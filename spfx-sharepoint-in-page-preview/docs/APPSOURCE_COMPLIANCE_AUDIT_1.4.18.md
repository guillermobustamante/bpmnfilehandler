# 1. Executive Summary

Static audit result: **not ready for AppSource submission**.

I found **4 Critical store blockers**, **5 Warnings**, and **3 Optimizations** in the `1.4.18.0` SPFx package. The package builds and the supplied `.sppkg` contains the expected version, but it currently violates AppSource/SPFx marketplace requirements around DOM isolation, marketplace developer metadata, root-site independence, and hardcoded external-service assumptions.

Primary Microsoft references used:

- SPFx Marketplace checklist: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-checklist
- SPFx overview and DOM support boundary: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/sharepoint-framework-overview
- SPFx tenant-wide extension deployment: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions
- Microsoft 365 app publishing checklist: https://learn.microsoft.com/en-us/partner-center/marketplace-offers/checklist
- Microsoft Marketplace certification policies: https://learn.microsoft.com/en-us/legal/marketplace/certification-policies

# 2. Store Compliance & Technical Audit

### Critical (Store Blocker) - Direct DOM Manipulation Outside SPFx Boundary
* **Category:** AppSource Compliance / Security
* **Location:** `src/extensions/filePreviewAdmin/FilePreviewAdminApplicationCustomizer.ts` lines 77-124, 133-143, 361-391, 418-439; `src/extensions/bpmnOpenCommandSet/BpmnViewerDialog.ts` lines 262-389, 735-738, 786-809; `src/extensions/bpmnOpenCommandSet/DrawioViewerDialog.ts` lines 218-345, 663-665, 726-749
* **Description:** The application customizer appends styles to `document.head`, appends command buttons directly to `document.body`, polls and queries SharePoint list rows with global `document.querySelectorAll`, and infers selected files from SharePoint page markup. Both viewer dialogs create portal hosts in `document.body`, hide the SPFx dialog `domElement`, traverse and restyle `.ms-Modal`, `.ms-Dialog-main`, `.ms-Layer-content`, focus-trap, and overlay elements outside the component-owned DOM. This takes a dependency on SharePoint page DOM and Fabric dialog internals.
* **AppSource Rule (If Applicable):** Microsoft states that SPFx web parts/customizers must only manipulate the DOM element provided through `domElement`, and directly manipulating the page DOM will result in rejection. Microsoft also states that SharePoint page HTML DOM is not an API and should not be depended on.
* **Remediation:** Keep all UI inside SPFx-provided surfaces. For the admin launcher, render through a supported placeholder or remove the floating launcher and use the admin web part as the first-run entry point. For dialogs, render inside `this.domElement` instead of a `document.body` portal and avoid restyling Fluent/SharePoint ancestors. For file selection, rely on `BaseListViewCommandSet` `selectedRows`.

```ts
// Application customizer: use a supported placeholder instead of document.body.
const placeholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
if (placeholder) {
  this.launcherElement = document.createElement('button');
  this.launcherElement.textContent = 'File Preview Admin';
  placeholder.domElement.appendChild(this.launcherElement);
}

// Dialog: keep the renderer inside BaseDialog.domElement.
public render(): void {
  this.domElement.classList.add('bpf-preview-dialog');
  this.domElement.innerHTML = renderPreviewMarkup(this.fileName);
}
```

### Critical (Store Blocker) - Required Marketplace Developer Metadata Is Missing
* **Category:** AppSource Compliance
* **Location:** `config/package-solution.json` lines 10-15; packaged `sharepoint/solution/spfx-bpmn-command-set.sppkg` `AppManifest.xml` developer properties
* **Description:** `privacyUrl` and `termsOfUseUrl` are empty, and `mpnId` is the placeholder value `Undefined-1.23.0`. The packaged `.sppkg` carries those empty/placeholder values into `AppManifest.xml`.
* **AppSource Rule (If Applicable):** The SPFx Marketplace checklist requires valid developer metadata including organization name, website URL, privacy statement URL, and terms of use URL. Microsoft Marketplace policies also require relevant offer information and a privacy policy that details customer data collection, use, and storage.
* **Remediation:** Publish production privacy and terms pages before submission, then populate the manifest with valid HTTPS URLs and the real Partner ID if available.

```json
"developer": {
  "name": "Evolve Global Solutions",
  "websiteUrl": "https://evolvegs.ca",
  "privacyUrl": "https://evolvegs.ca/privacy",
  "termsOfUseUrl": "https://evolvegs.ca/terms",
  "mpnId": "<real-partner-id>"
}
```

### Critical (Store Blocker) - Root-Site Configuration Dependency Breaks Non-Root Validation
* **Category:** AppSource Compliance / Architecture
* **Location:** `src/extensions/bpmnOpenCommandSet/BpmnOpenCommandSetCommandSet.ts` lines 185-187 and 246-248; `src/extensions/filePreviewAdmin/FilePreviewAdminApplicationCustomizer.ts` lines 337-358 and 395-398; `src/webparts/filePreviewAdmin/FilePreviewAdminWebPart.ts` lines 121-135 and 245-247; `src/extensions/bpmnOpenCommandSet/previewSettings.ts` lines 332-335
* **Description:** Tenant-wide component properties do not set `configSiteUrl`, so command-set and customizer runtime defaults derive the configuration location from the tenant root site. Error text also tells users to open the tenant root site. This creates a root-site dependency for non-root site collection validation and for tenants where the root site is locked down, replaced, or unavailable to the target users.
* **AppSource Rule (If Applicable):** The SPFx Marketplace checklist requires testing in both root and non-root sites to verify the app has no dependencies on specific site URLs.
* **Remediation:** Make the configuration location explicit and tenant-independent. Use the app catalog site or a dedicated admin site as the default, write `configSiteUrl` into `ClientSideInstance.xml`, and fail with a clear setup state instead of silently assuming the tenant root.

```xml
<ClientSideComponentInstance
    Title="Open file"
    Location="ClientSideExtension.ListViewCommandSet.CommandBar"
    ListTemplateId="101"
    Sequence="1"
    Properties="{&quot;configSiteUrl&quot;:&quot;{appCatalogSiteUrlOrConfiguredAdminSite}&quot;,&quot;showSettingsCommand&quot;:false}"
    ComponentId="c3e13f04-c3e1-4b55-8fd5-d7557cd15752" />
```

### Critical (Store Blocker) - Hardcoded Hosted Endpoint and Fixed Admin Client ID
* **Category:** AppSource Compliance / Security
* **Location:** `src/extensions/bpmnOpenCommandSet/PreviewSettingsDialog.ts` lines 712-727; `scripts/Install-FilePreviewTenant.ps1` lines 18-23 and 46; packaged admin UI generated cleanup script
* **Description:** The cleanup script generated in the admin dialog falls back to `https://bpmn-file-handler-2f18b433.azurewebsites.net` when no endpoint is configured. The PowerShell installer also embeds a fixed `$ClientID` for PnP sign-in. Even though a client ID is not a secret, fixed external-service identifiers and a hardcoded Azure endpoint create tenant-specific/vendor-specific assumptions in customer-facing code and are difficult to justify during enterprise security review.
* **AppSource Rule (If Applicable):** Marketplace guidance requires offers to work as described without hidden tenant-specific dependencies, and all displayed URLs must be working, accurate, and safe. The SPFx checklist also requires explicit disclosure of required settings.
* **Remediation:** Remove fallback production endpoints from generated scripts. Require the administrator to enter an endpoint explicitly, or generate placeholders. If a helper script is retained, make the authentication app/client ID a parameter and document who owns it.

```ts
const endpoint = settings.appBaseUrl;
const endpointForScript = endpoint || '<file-handler-endpoint-url>';
```

```powershell
param(
  [Parameter(Mandatory = $true)]
  [string]$TenantRootUrl,
  [Parameter(Mandatory = $true)]
  [string]$PnPClientId
)
Connect-PnPOnline -Url $TenantRootUrl -Interactive -ClientId $PnPClientId
```

### Warning (Functional Risk) - Mixed Tenant-Wide Deployment Mechanisms
* **Category:** AppSource Compliance / Maintainability
* **Location:** `config/package-solution.json` lines 34-38; `sharepoint/assets/ClientSideInstance.xml` lines 3-21; `sharepoint/assets/elements.xml` lines 3-26; `scripts/Install-FilePreviewTenant.ps1` lines 220-324
* **Description:** The package has `skipFeatureDeployment: true` and includes `ClientSideInstance.xml`, which is the tenant-wide extension path. It also includes `elements.xml` custom actions, and the optional installer directly writes entries to the Tenant Wide Extensions list. This creates multiple activation paths for the same component IDs and locations, raising the risk of duplicate commands, inconsistent properties, and hard-to-reproduce deployment behavior.
* **AppSource Rule (If Applicable):** Microsoft documents `ClientSideInstance.xml` as the automation mechanism used during app catalog activation when `skipFeatureDeployment` is true.
* **Remediation:** Choose one deployment path for AppSource. For tenant-wide AppSource deployment, keep `skipFeatureDeployment: true` plus `ClientSideInstance.xml`, remove `elements.xml` custom actions, and do not require a post-install script to edit Tenant Wide Extensions. If site-scoped activation is required, set up a separate package.

### Warning (Functional Risk) - External diagrams.net Renderer Transfers Customer File XML
* **Category:** Security / Data Handling
* **Location:** `src/extensions/bpmnOpenCommandSet/DrawioViewerDialog.ts` lines 6-8, 63, 137-190; `src/extensions/bpmnOpenCommandSet/PreviewSettingsDialog.ts` lines 84-87; `scripts/Install-FilePreviewTenant.ps1` lines 31-33
* **Description:** `.drawio` preview loads `https://embed.diagrams.net` in an iframe and posts SharePoint file XML to that external origin. The runtime default in `previewSettings.ts` disables `.drawio`, but the installer default JSON enables `.drawio`, creating inconsistent behavior and a potential unapproved external data flow. The origin check in `onMessage` is good, but the data transfer still requires explicit customer disclosure and admin consent.
* **AppSource Rule (If Applicable):** Marketplace policies require privacy disclosures for collection, use, and storage of customer data, and offer links/URLs must be accurate and safe.
* **Remediation:** Keep `.drawio` disabled in every default path, add an explicit admin acknowledgement before enabling external renderers, document diagrams.net as a subprocess/external processor in privacy terms, and include this data flow in certification notes.

### Warning (Functional Risk) - License Key Stored in Plaintext SharePoint List JSON
* **Category:** Security / Data Handling
* **Location:** `src/extensions/bpmnOpenCommandSet/PreviewSettingsDialog.ts` lines 90-109 and 242-268; `src/extensions/bpmnOpenCommandSet/previewSettings.ts` lines 15-20, 116-120, 179-187
* **Description:** The license key is read from a password input but then stored directly inside `BpfConfigJson` in the `M365 File Preview Settings` SharePoint list. A password input only masks the value in the browser; it does not protect the value at rest from site collection administrators, list owners, export tools, backups, or audit reviewers.
* **AppSource Rule (If Applicable):** Marketplace privacy/security review expects accurate disclosure and appropriate handling of customer data. If the license key is a secret, this storage model is insufficient.
* **Remediation:** Treat the key as a non-secret license token and document that explicitly, or move validation to a backend licensing service and store only tenant ID, license tier, expiration, and an opaque validation status/hash in SharePoint.

### Warning (Functional Risk) - Generated Admin Scripts Perform Tenant/Entra Changes From In-App UI
* **Category:** Security / Enterprise Governance
* **Location:** `src/extensions/bpmnOpenCommandSet/PreviewSettingsDialog.ts` lines 130-145 and 650-705; `scripts/Install-FilePreviewTenant.ps1` lines 428-453
* **Description:** The admin dialog generates copyable Azure CLI/PowerShell scripts that create/update Entra application addIns and remove existing File Handler registrations. This is powerful tenant administration behavior exposed inside the SPFx app UI. It may be valid for advanced setup, but it raises governance concerns because the store-delivered app is instructing admins to execute privileged scripts outside SharePoint.
* **AppSource Rule (If Applicable):** The SPFx checklist requires configuration requirements to be clearly stated, and Marketplace policies require accurate documentation and safe customer-facing links/instructions.
* **Remediation:** Move tenant/Entra scripts to external administrator documentation or a separately distributed enterprise deployment guide. Keep the SPFx UI to status, validation, and links to documentation. If scripts remain in the app, add a signed-script download flow, versioned docs, and explicit warnings about required roles and tenant impact.

### Warning (Functional Risk) - Full-Screen Dialog Behavior Depends on Fluent UI Internals
* **Category:** Code Quality / Stability
* **Location:** `src/extensions/bpmnOpenCommandSet/BpmnViewerDialog.ts` lines 262-383; `src/extensions/bpmnOpenCommandSet/DrawioViewerDialog.ts` lines 218-339
* **Description:** The dialogs force full-screen behavior by traversing up to `.ms-Modal`, `.ms-Dialog-main`, `.ms-Dialog-inner`, `.ms-Dialog-content`, `.ms-Modal-scrollableContent`, and `.ms-Layer-content`, then applying `!important` inline styles. Even aside from AppSource DOM isolation, this is brittle against Fluent UI and SharePoint host updates.
* **AppSource Rule (If Applicable):** SharePoint page DOM and CSS styles are not supported APIs, and Microsoft requires components to work without errors across supported hosts.
* **Remediation:** Use a supported dialog layout within `BaseDialog.domElement`, or route full-screen editing to a dedicated SharePoint page/web part where your component owns the layout. Avoid relying on host CSS class names.

### Optimization (Best Practice) - Selection Polling Scans Page DOM Every 750ms
* **Category:** Performance
* **Location:** `src/extensions/filePreviewAdmin/FilePreviewAdminApplicationCustomizer.ts` lines 219-223 and 418-439
* **Description:** The application customizer polls every 750ms and scans global SharePoint row/checkbox selectors to infer the selected document. This can add sustained CPU work to document libraries and is fragile because it depends on current SharePoint list markup.
* **AppSource Rule (If Applicable):** The SPFx checklist recommends verifying that the application does not cause sustained CPU usage or browser unresponsiveness.
* **Remediation:** Remove this watcher and expose file preview only through the ListView Command Set, which already receives selection state through supported SPFx APIs.

### Optimization (Best Practice) - No Automated Tests Found in Production Build
* **Category:** Code Quality
* **Location:** Build output from `npm run build`; project test configuration
* **Description:** The production build completed, but Jest reported `No tests found`. This is not itself a store blocker, but the solution includes high-risk behavior: SharePoint REST reads/writes, settings normalization, external iframe messaging, and XML save flows. These should have automated coverage before AppSource validation and enterprise deployment.
* **AppSource Rule (If Applicable):** Microsoft asks publishers to verify that the app works without errors and provides the advertised functionality across supported browsers and sites.
* **Remediation:** Add unit tests for `normalizeSettings`, file URL conversion, renderer selection, and postMessage parsing. Add Playwright or SPFx integration tests for command visibility and first-run admin setup in root and non-root sites.

### Optimization (Best Practice) - Duplicated Inline SVG and `innerHTML` Rendering
* **Category:** Code Quality / Security Hardening
* **Location:** `src/extensions/bpmnOpenCommandSet/BpmnViewerDialog.ts` lines 40-79 and 770-783; `src/extensions/bpmnOpenCommandSet/DrawioViewerDialog.ts` lines 38-64 and 714-723; `src/webparts/filePreviewAdmin/FilePreviewAdminWebPart.ts` lines 39-95
* **Description:** The code renders sizeable UI strings and duplicated SVG path maps through `innerHTML`. Current user-controlled values are generally escaped, which reduces immediate XSS risk, but repeated string templates make future regressions more likely and complicate accessibility/theming.
* **AppSource Rule (If Applicable):** The SPFx checklist expects polished, reliable UI and recommends Microsoft 365 design alignment.
* **Remediation:** Centralize icon rendering, prefer DOM construction or a React component model, and keep all dynamic text assigned through `textContent` or escaped attributes. This also makes theme support and keyboard/focus behavior easier to test.

### Optimization (Best Practice) - Full Page Web Part Missing Recommended Full-Page Icon Metadata
* **Category:** AppSource Compliance / UX
* **Location:** `src/webparts/filePreviewAdmin/FilePreviewAdminWebPart.manifest.json` lines 10-31
* **Description:** The web part supports `SharePointFullPage`, but the manifest does not include a `fullPageAppIconImageUrl`. Microsoft lists a properly sized full-page image as a recommended marketplace check for single part app pages.
* **AppSource Rule (If Applicable):** SPFx Marketplace recommended check: full page image should be sized properly when components are exposed as single part app pages.
* **Remediation:** Add a 193x158 asset and reference it in the web part manifest, or remove `SharePointFullPage` support if the admin web part should only be embedded on a normal page.

