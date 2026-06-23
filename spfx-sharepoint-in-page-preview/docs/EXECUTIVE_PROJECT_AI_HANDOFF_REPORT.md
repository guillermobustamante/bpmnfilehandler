# Executive Project & AI Handoff Report

## 1. Executive Summary

SPFx SharePoint In-Page Preview is a SharePoint Framework solution that lets users open supported process and diagram files directly from SharePoint document libraries without leaving the page. The current package version is `1.4.18.0`, and the deployable package is `sharepoint/solution/spfx-bpmn-command-set.sppkg`.

The project solves a common Microsoft 365 content problem: specialized engineering, process, and architecture files are often stored in SharePoint, but users must download them or launch separate tools to inspect them. This solution adds a document-library command for same-page preview/editing, central tenant settings, and optional native Microsoft 365 File Handler registration guidance.

Strategically, the project positions Evolve Global Solutions to offer a lightweight, SharePoint-first file preview framework for process and technical document workflows. It has meaningful business value for organizations standardizing on Microsoft 365 because it keeps files, permissions, and user workflows inside SharePoint. However, the current code is suitable for controlled tenant deployment, not AppSource submission, until the compliance risks documented in `docs/APPSOURCE_COMPLIANCE_AUDIT_1.4.18.md` are resolved.

Source boundary: this report is based on the current post-split SPFx codebase at `spfx-sharepoint-in-page-preview`. The older SharePoint documentation links were not accessible from this session because they require Microsoft tenant authentication, and pre-split code is intentionally excluded.

## 2. Industry & Audience Analysis

* **Industry Context:** Microsoft 365 customers increasingly expect business-specific file experiences to stay inside SharePoint, Teams, and OneDrive instead of forcing users into unmanaged desktop/download workflows. SPFx is the supported SharePoint Online customization model, and this project aligns with tenant-admin-controlled deployment, current-user permissions, and SharePoint-hosted client-side assets. The solution is also aligned with the broader market demand for lightweight process/document visualization in regulated collaboration environments, but AppSource readiness requires stricter DOM isolation, metadata, privacy, and multi-tenant setup controls.
* **Target Audience:** Primary end users are process analysts, operations teams, engineering teams, enterprise architects, consultants, and Microsoft 365 users who store BPMN or diagram assets in SharePoint document libraries. Tenant administrators are a secondary audience: they configure enabled extensions, licensing, optional File Handler endpoints, and central settings. Business value to users is faster review of BPMN and diagram files, fewer downloads, and less context switching. Business value to admins is a centrally governed preview experience that inherits SharePoint permissions.
* **Competitive Advantage:** The solution is SharePoint-first and does not require an Azure web app for the core BPMN preview path. BPMN rendering is bundled through `bpmn-js`, and file reads/writes use SharePoint REST with the current user's permissions. The architecture separates the AppSource-oriented SPFx package from the hosted Microsoft 365 File Handler/Azure app, reducing runtime hosting obligations for the SharePoint preview scenario. The extensible settings model already anticipates additional file types such as `.drawio`, `.jt`, and `.step`, giving the product a path from BPMN preview into a broader engineering/process file framework.

## 3. Architecture & Technical Map

* **Tech Stack:** TypeScript, SharePoint Framework `1.23.0`, Heft-based SPFx build tooling, Node.js `>=22.14.0 <23.0.0`, SharePoint REST via `@microsoft/sp-http`, ListView Command Set via `@microsoft/sp-listview-extensibility`, Application Customizer via `@microsoft/sp-application-base`, Web Part via `@microsoft/sp-webpart-base`, SPFx Dialog via `@microsoft/sp-dialog`, BPMN rendering/editing via `bpmn-js 18.18.0`, optional diagrams.net iframe renderer at `https://embed.diagrams.net`, and optional PnP PowerShell/Azure CLI scripts for tenant administration and native File Handler registration.
* **System Flow:**
1. Tenant admin uploads `sharepoint/solution/spfx-bpmn-command-set.sppkg` to the SharePoint App Catalog and deploys it tenant-wide.
2. `config/package-solution.json` declares `skipFeatureDeployment: true`, bundled client-side assets, and feature assets including `ClientSideInstance.xml` and `elements.xml`.
3. `ClientSideInstance.xml` registers the command set in document library command bar and context menu locations, plus an admin application customizer.
4. In a document library, `BpmnOpenCommandSetCommandSet` observes SPFx list selection state and displays `Open BPMN`, `Open DrawIO`, or a generic `Open file` command when one supported file is selected.
5. The command set loads central settings through `PreviewSettingsService`, using a SharePoint list named `M365 File Preview Settings` and a note field named `BpfConfigJson`.
6. If the selected extension is enabled and uses `bpmn-js`, `BpmnViewerDialog` opens, reads file metadata/content through `SharePointFileService`, renders the BPMN model, supports validation, reload, download, zoom, and save when in modeler mode.
7. If the selected extension is enabled and uses `diagrams-net-embed`, `DrawioViewerDialog` opens an iframe to diagrams.net, posts the XML payload to that iframe, and saves returned XML back to SharePoint when editable.
8. Tenant admins use `FilePreviewAdminApplicationCustomizer`, `FilePreviewAdminWebPart`, or the optional settings command to initialize and edit central settings.
9. `PreviewSettingsDialog` controls enabled extensions, license tier/user count/key, optional File Handler endpoint, and generated scripts for native Microsoft 365 File Handler registration or cleanup.
10. `scripts/Install-FilePreviewTenant.ps1` is an advanced deployment/troubleshooting path that can upload the package, create the config list, upload icon assets, write Tenant Wide Extensions entries, and optionally create an admin page.
* **External Dependencies:** SharePoint Online App Catalog, SharePoint REST APIs, tenant-wide extension list, current user's SharePoint permissions, `bpmn-js`, diagrams.net embed for optional `.drawio` files, PnP PowerShell, Azure CLI for optional Entra app `addIns` File Handler metadata, Evolve website metadata in package manifest, and static icon assets under `sharepoint/assets/file-handler-icons`.

## 4. Project Health & Current State

* **What is working:** The SPFx solution builds in a clean dependency environment and the generated `.sppkg` has version `1.4.18.0`. Core SharePoint package metadata, component manifests, command-set registration, bundled client-side assets, BPMN preview/editing path, central SharePoint settings service, admin web part, admin application customizer, optional diagrams.net renderer, file read/write service, icon assets, and deployment documentation are present. The SharePoint-first BPMN path does not require Azure hosting.
* **What is pending:** AppSource compliance remediation is pending. Required marketplace metadata is incomplete. The admin and dialog UI need to be brought inside supported SPFx DOM boundaries. Root-site configuration assumptions need replacement with a tenant-independent setup model. Mixed tenant-wide deployment mechanisms need simplification. The optional File Handler/admin scripts need governance cleanup. Automated tests are not present; the last production build reported no Jest tests. The local OneDrive `node_modules` folder was previously locked/incomplete, so reliable builds may require a clean folder outside OneDrive or a fresh dependency install.
* **Critical Risks:** The highest risk is AppSource rejection due to direct DOM manipulation outside `this.domElement`. Additional executive risks include external `.drawio` XML transfer to diagrams.net if enabled, plaintext license key storage inside SharePoint list JSON, hardcoded fallback Azure endpoint in generated cleanup script, fixed PnP client ID in installer script, and root-site dependency for configuration. There is also deployment ambiguity because both `ClientSideInstance.xml`, `elements.xml`, and the PowerShell installer can register overlapping tenant-wide extensions.

## 5. Architectural Constraints (Do Not Touch List)

Do not merge this SPFx package back with the hosted Microsoft 365 File Handler/Azure app. The split is intentional: this repo is the SharePoint/AppSource candidate and must remain focused on SharePoint-hosted SPFx assets. Hosted File Handler POST launch endpoints belong outside this package.

Do not change existing component IDs unless intentionally creating a breaking new SharePoint app. Preserve `COMMAND_SET_COMPONENT_ID` `c3e13f04-c3e1-4b55-8fd5-d7557cd15752`, `ADMIN_WEB_PART_COMPONENT_ID` `58ccbc3f-a8dd-40ed-b4f3-c4a647338da8`, `ADMIN_APPLICATION_CUSTOMIZER_COMPONENT_ID` `6da1de09-f3e8-4d81-a60b-0bf2c8c65be4`, solution ID `2ab9f11d-b745-4c32-b87a-460301fccf91`, and feature ID `45deb6cb-3879-401d-8133-4cb2cd5fe682`. These IDs are how SharePoint recognizes upgrades instead of installing unrelated apps.

Do not remove the current SharePoint REST permission model without a deliberate security review. File content operations currently run as the signed-in user through `SPHttpClient`, which preserves SharePoint authorization, auditing, and site-level access controls. Moving file access to Graph or a backend service would change consent, permissions, and customer security posture.

Do not make `.drawio` enabled by default in the SPFx runtime path. It uses an external iframe and posts file XML to diagrams.net. Keeping it disabled by default is a privacy and enterprise governance tradeoff; admins must opt in after disclosure.

Do not treat `M365 File Preview Settings`, `BpfConfigJson`, and `Configuration` as casual strings. These names are the current persistence contract for tenant settings. If they change, migration logic is required so existing deployed tenants do not lose settings.

Do not replace `bpmn-js` with a custom BPMN renderer without a strong reason. BPMN rendering/modeling is domain-specific and complex; `bpmn-js` is the proven engine currently carrying the core value proposition.

Do not rely on pre-split code or old hosted Azure artifacts when making AppSource fixes. The authoritative current system is the post-split SPFx package. Pre-split materials may explain historical intent but should not be copied into this repo unless they match the current SharePoint-first architecture.

Do not keep expanding global DOM workarounds. Current DOM portal and SharePoint chrome styling helped produce a full-screen user experience, but they are also the largest AppSource blocker. Future work should move toward supported SPFx placeholders, web part surfaces, and dialog content owned by `this.domElement`.

Do not change package version inconsistently. Release updates must keep `package.json`, `package-lock.json`, `config/package-solution.json`, feature version, `APP_VERSION`, docs, and installer version aligned.

## 6. AI Handoff Initialization 

You are assuming support of this project, a post-split SharePoint Framework package named `spfx-sharepoint-in-page-preview` for in-page preview of BPMN and other supported files in SharePoint document libraries. Work only in `C:\Users\yetro\OneDrive - Evolve Global Solutions\EvolveGSInc\Projects\BPMN file handler\bpmnfilehandler\spfx-sharepoint-in-page-preview\` unless explicitly told otherwise. Do not use pre-split code as implementation source.

Current package version is `1.4.18.0`; deployable package path is `sharepoint/solution/spfx-bpmn-command-set.sppkg`. The project uses SPFx `1.23.0`, Node `22.x`, TypeScript, Heft, `bpmn-js`, SharePoint REST, a ListView Command Set, an Application Customizer, and an Admin Web Part. Main source files are:

- `src/extensions/bpmnOpenCommandSet/BpmnOpenCommandSetCommandSet.ts` for document-library command behavior.
- `src/extensions/bpmnOpenCommandSet/BpmnViewerDialog.ts` for SharePoint-hosted BPMN rendering/editing.
- `src/extensions/bpmnOpenCommandSet/DrawioViewerDialog.ts` for optional diagrams.net rendering.
- `src/extensions/bpmnOpenCommandSet/PreviewSettingsDialog.ts` for tenant admin settings and optional script generation.
- `src/extensions/bpmnOpenCommandSet/previewSettings.ts` for settings persistence in the `M365 File Preview Settings` SharePoint list.
- `src/extensions/bpmnOpenCommandSet/sharePointFileService.ts` for SharePoint file metadata/content read/write.
- `src/extensions/filePreviewAdmin/FilePreviewAdminApplicationCustomizer.ts` for the admin launcher and current floating preview launcher.
- `src/webparts/filePreviewAdmin/FilePreviewAdminWebPart.ts` for the admin page web part.
- `config/package-solution.json`, `sharepoint/assets/ClientSideInstance.xml`, and `sharepoint/assets/elements.xml` for packaging/deployment.

Immediate next steps:

1. Read `docs/APPSOURCE_COMPLIANCE_AUDIT_1.4.18.md` before changing code. Treat it as the active remediation backlog.
2. Fix AppSource blockers first: remove global DOM manipulation and body portals, populate valid developer metadata, remove root-site default assumptions, and remove hardcoded hosted endpoint/fixed admin client assumptions.
3. Simplify deployment registration: prefer tenant-wide `skipFeatureDeployment: true` plus `ClientSideInstance.xml`; avoid duplicate registration through `elements.xml` and post-install Tenant Wide Extensions edits unless the user explicitly wants a tenant-only deployment helper.
4. Preserve component IDs and the SharePoint settings contract unless migration is intentionally planned.
5. Keep the SharePoint-first BPMN path Azure-free. Optional native Microsoft 365 File Handler registration must remain secondary and clearly disclosed.
6. Add focused tests for settings normalization, file URL handling, renderer selection, postMessage parsing, and first-run configuration behavior.
7. Build in a clean folder if OneDrive locks `node_modules`; a prior successful production package was built from a temp copy because the workspace dependency folder was locked/incomplete.

When continuing, be pragmatic: keep edits scoped, verify with `npm run build`, update version metadata consistently for release packages, and report clearly whether the resulting package is suitable for tenant deployment only or ready for AppSource review.

