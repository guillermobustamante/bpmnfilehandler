# 1. Executive Summary

Static audit result: **not ready for AppSource submission**.

I found **3 Critical store blockers**, **5 Warnings**, and **2 Optimizations**. The strongest blocker is SPFx contract conformance: multiple components directly create, append, query, and restyle DOM outside the SPFx-owned `domElement`, which Microsoft says will result in rejection. I verified against Microsoft Marketplace/SPFx publishing guidance, including SPFx v1.11+ requirements, developer metadata requirements, DOM isolation rules, and tenant-wide deployment behavior.

Sources:

- Microsoft Marketplace certification policies: https://learn.microsoft.com/en-us/legal/marketplace/certification-policies
- SPFx Marketplace checklist: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-checklist
- SPFx tenant-wide deployment docs: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions

# 2. Store Compliance & Technical Audit

### Critical (Store Blocker) - Direct DOM Manipulation Outside SPFx Boundary

* **Category:** AppSource Compliance / Security
* **Location:** `FilePreviewAdminApplicationCustomizer.ts` lines 77-124, 361-391, 426-439; `BpmnViewerDialog.ts` lines 262-389, 735-738, 805-809; `DrawioViewerDialog.ts` lines 218-345, 663-665, 745-749
* **Description:** The code appends styles to `document.head`, appends buttons/portal hosts to `document.body`, queries SharePoint page rows with `document.querySelectorAll`, and mutates Microsoft dialog/layer/overlay elements outside the component-owned DOM. Microsoft’s SPFx Marketplace checklist states web parts/customizers must only manipulate the provided `domElement`; direct page DOM manipulation causes rejection.
* **AppSource Rule:** SPFx Contract Conformance / DOM isolation.
* **Remediation:** Keep UI inside `this.domElement` or official SPFx placeholders/dialog content. Remove body portals and SharePoint chrome restyling.

```ts
// Application customizer: render inside a top/bottom placeholder, not document.body.
const placeholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom);
if (placeholder) {
  placeholder.domElement.appendChild(this.launcherElement);
}

// Dialog: use this.domElement as root; do not append a host to document.body.
const root = this.domElement;
root.classList.add('bpf-preview-dialog');
root.innerHTML = renderDialogMarkup();
```

### Critical (Store Blocker) - Required Marketplace Developer Metadata Is Missing

* **Category:** AppSource Compliance
* **Location:** `spfx-sharepoint-in-page-preview/config/package-solution.json` lines 10-15
* **Description:** `privacyUrl` and `termsOfUseUrl` are empty, and `mpnId` is a placeholder value: `"Undefined-1.23.0"`. Microsoft’s checklist marks privacy statement and terms URLs as mandatory developer metadata.
* **AppSource Rule:** Solution package must contain valid developer metadata.
* **Remediation:** Populate production URLs and a real Partner ID or omit only if intentionally unavailable.

```json
"developer": {
  "name": "Evolve Global Solutions",
  "websiteUrl": "https://evolvegs.ca",
  "privacyUrl": "https://evolvegs.ca/privacy",
  "termsOfUseUrl": "https://evolvegs.ca/terms",
  "mpnId": "<real-partner-id>"
}
```

### Critical (Store Blocker) - Hardcoded Tenant and Hosted-Service Artifacts

* **Category:** AppSource Compliance / Security
* **Location:** `spfx-sharepoint-in-page-preview/config/serve.json` lines 7, 18; `microsoft-365-file-handler-hosted-azure-app/infra/deployment.json` lines 2-13; `PreviewSettingsDialog.ts` line 717
* **Description:** The repo contains a real tenant hostname, tenant ID, Azure subscription ID, app service name, Entra object IDs, client ID, and a hardcoded Azure fallback endpoint. Store validation and enterprise source review commonly reject tenant-specific dependencies or identifiers in distributable code and docs.
* **AppSource Rule:** Must work across tenants and root/non-root sites without specific URL dependencies.
* **Remediation:** Remove `infra/deployment.json` from distributable source, replace serve URLs with placeholders, and remove the Azure fallback from generated cleanup scripts.

```ts
const endpoint = settings.appBaseUrl || '<handler-endpoint-url>';
```

### Warning (Functional Risk) - Optional External Renderer Transfers File XML to diagrams.net

* **Category:** Security / AppSource Compliance
* **Location:** `DrawioViewerDialog.ts` lines 6-8, 181-190; `PreviewSettingsDialog.ts` lines 84-87; `Install-FilePreviewTenant.ps1` lines 31-32, 108-110
* **Description:** `.drawio` files are loaded into `https://embed.diagrams.net` via iframe/postMessage. The SPFx default disables `.drawio`, but the tenant installer enables it by default. This is a material external data flow and must be explicitly disclosed in privacy, certification notes, and admin consent UX.
* **Remediation:** Keep `.drawio` disabled by default everywhere, add an admin acknowledgement before enabling it, and document the third-party data flow in the privacy statement and Notes for Certification.

### Warning (Functional Risk) - Plaintext License Key Stored in SharePoint List

* **Category:** Security / Data Handling
* **Location:** `PreviewSettingsDialog.ts` lines 107-108, 263-267; `previewSettings.ts` lines 185-187
* **Description:** Paid-tier license keys are stored as JSON in a SharePoint list field. A password input hides it visually but does not protect it at rest from users/admins with list access.
* **Remediation:** Treat it as a non-secret license token, or move validation to a backend service and store only a tenant-scoped license status/opaque hash in SharePoint.

### Warning (Functional Risk) - Hosted File Handler Uses Broad Graph Permissions and Browser Local Storage

* **Category:** Security / AppSource Compliance
* **Location:** `microsoft-365-file-handler-hosted-azure-app/server/index.ts` lines 81-84; `microsoft-365-file-handler-hosted-azure-app/scripts/create-entra-app.ps1` line 10; `microsoft-365-file-handler-hosted-azure-app/src/auth/msal.ts` lines 19-21
* **Description:** The optional hosted app defaults to delegated `Files.ReadWrite.All` and MSAL `localStorage`. High permissions must be justified during submission, and persistent token cache increases exposure on shared or compromised devices.
* **AppSource Rule:** High permissions need justification and manifest-driven configuration.
* **Remediation:** Use least-privilege scopes where possible, provide written justification, and prefer `sessionStorage` unless persistent SSO is required.

### Warning (Functional Risk) - Root-Site Configuration Assumption Can Fail Non-Root Validation

* **Category:** AppSource Compliance / Code Quality
* **Location:** `BpmnOpenCommandSetCommandSet.ts` lines 185-187; `FilePreviewAdminApplicationCustomizer.ts` lines 337-339; `previewSettings.ts` lines 332-335
* **Description:** Runtime defaults derive configuration from the tenant root site, and error text tells admins to open the tenant root. Microsoft requires validation in non-root modern sites and recommends testing both root and non-root to avoid URL dependencies.
* **Remediation:** Default to the current site or App Catalog config site for first run, and require `configSiteUrl` to be supplied by package properties for tenant-wide deployment.

### Warning (Functional Risk) - Mixed Tenant-Wide Deployment Mechanisms

* **Category:** AppSource Compliance / Maintainability
* **Location:** `package-solution.json` lines 34-38; `elements.xml` lines 3-26; `ClientSideInstance.xml` lines 3-21; `Install-FilePreviewTenant.ps1` lines 220-324
* **Description:** The package includes both `elements.xml` custom actions and `ClientSideInstance.xml`, while the installer script also writes Tenant Wide Extensions list entries. This creates duplicate/ambiguous activation paths and makes deployment behavior harder to certify.
* **Remediation:** Use one declarative AppSource path: `skipFeatureDeployment: true` plus `ClientSideInstance.xml`, or site-scoped feature activation, but not both plus a post-install script.

### Optimization (Best Practice) - Polling the Page DOM Every 750ms

* **Category:** Performance
* **Location:** `FilePreviewAdminApplicationCustomizer.ts` lines 219-223, 418-439
* **Description:** The application customizer scans global SharePoint DOM repeatedly to infer selected files. This is brittle and can add overhead on large library pages.
* **Remediation:** Rely on the ListView Command Set selection APIs for file actions and remove the floating preview launcher from the application customizer.

### Optimization (Best Practice) - Generated SVG Icons and Inline Markup Should Be Centralized

* **Category:** Code Quality
* **Location:** `BpmnViewerDialog.ts` lines 770-783; `DrawioViewerDialog.ts` lines 714-723
* **Description:** Icon SVG strings are duplicated across dialogs and injected with `innerHTML`. Current values are static, but centralizing them reduces future XSS and maintenance risk.
* **Remediation:** Move icons to a shared helper returning `HTMLElement` nodes, or use a trusted icon library/component consistently.
