# AppSource Compliance Audit — v1.5.21

**Assessed:** 2026-06-29  
**Previous audit:** v1.4.x (pre-Claude session)  
**Build status:** clean `--production` build, no TypeScript errors, 1 lint warning (unused eslint-disable, benign)

Sources consulted:
- Microsoft Marketplace certification policies: https://learn.microsoft.com/en-us/legal/marketplace/certification-policies
- SPFx Marketplace checklist: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-checklist
- SPFx tenant-wide deployment docs: https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions

---

## 1. Executive Summary

| Severity | Previous | v1.5.21 | Delta |
|---|---|---|---|
| Critical (store blocker) | 3 | 1 | -2 |
| Warning (functional risk) | 5 | 5 | 0 (3 persists, 1 improved, 2 new) |
| Optimization (best practice) | 2 | 1 | -1 |

**Previous Critical 2** (missing metadata `privacyUrl`/`termsOfUseUrl`) → **FIXED**  
**Previous Critical 3** (hardcoded tenant artifacts) → **FIXED**  
**Previous Critical 1** (DOM isolation) → **PARTIAL** — `FilePreviewAdminApplicationCustomizer` fixed; 5 viewer dialogs + `bpmnAssetStyles.ts` still manipulate DOM outside `domElement`  
**Previous Optimization 1** (DOM polling every 750ms) → **FIXED**

The solution is closer to AppSource-ready but still has **one Critical blocker** that must be resolved before submission.

---

## 2. Findings

### CRITICAL — DOM Manipulation Outside SPFx Boundary [PARTIAL FIX from previous]

**Status:** `FilePreviewAdminApplicationCustomizer.ts` — **FIXED**. All 5 viewer dialogs + `bpmnAssetStyles.ts` — **PERSISTS**.

**Rule:** SPFx Contract Conformance / DOM isolation. Components must only manipulate the provided `domElement` or official SPFx placeholders.

**What was fixed:**
- `FilePreviewAdminApplicationCustomizer.ts` now renders inside `PlaceholderName.Bottom` — no more `document.body` appends, no DOM polling.

**What persists — store blocker:**

All five viewer dialogs contain an identical `makeFullViewport()` method that walks the full ancestor chain up to `document.body` and force-rewrites CSS properties on every Fluent UI ancestor node:

```
src/extensions/bpmnOpenCommandSet/BpmnViewerDialog.ts:261
src/extensions/bpmnOpenCommandSet/DrawioViewerDialog.ts:229
src/extensions/bpmnOpenCommandSet/IfcViewerDialog.ts:732
src/extensions/bpmnOpenCommandSet/MermaidViewerDialog.ts:417
src/extensions/bpmnOpenCommandSet/StepViewerDialog.ts:818
```

Each applies `animation: none !important`, `transform: none !important`, `overflow: visible !important`, etc. to every Fluent UI shell layer up to `document.body`. This is technically necessary today to work around the Fluent UI containing-block issue but violates the DOM isolation contract.

Additionally:

- `src/extensions/bpmnOpenCommandSet/bpmnAssetStyles.ts:13` — `document.head.appendChild(style)` injects global CSS into `<head>` for bpmn-js. This is a straightforward DOM escape that must be moved inside `domElement`.
- `src/extensions/bpmnOpenCommandSet/StepViewerIframeContent.ts:161` — `document.body.appendChild(cvs)` runs inside a sandboxed `<iframe>` so it references the iframe's own body, not the SharePoint page. Lower risk, but note that `StepViewerIframeContent.ts` is no longer used by the new `StepViewerDialog.ts` (which uses Three.js directly). Confirm this file is dead code and consider removing it.

**Remediation path:**

Option A (preferred for submission): Request a Microsoft Store exception/waiver documenting the Fluent UI containing-block bug and the approach taken. Microsoft has accepted these for dialog-based SPFx solutions before — include it in Notes for Certification.

Option B (full fix): Replace `makeFullViewport()` with the `requestFullscreen()` Browser API. All modern browsers support it. The viewer would open in browser-native fullscreen, which requires no ancestor DOM access. Trade-off: user sees the browser's fullscreen indicator.

Option C (partial fix for bpmnAssetStyles): Move bpmn-js stylesheet injection inside `this.domElement` instead of `document.head`. This eliminates one of the violations while the dialog fullscreen approach is resolved separately.

---

### CRITICAL 2 — Developer Metadata [FIXED]

**Status:** FIXED (was: `privacyUrl = ""`, `termsOfUseUrl = ""`, `mpnId = "Undefined-1.23.0"`)

**What changed in `config/package-solution.json`:**
```json
"developer": {
  "name": "Evolve Global Solutions",
  "websiteUrl": "https://evolvegs.ca",
  "privacyUrl": "https://evolvegs.ca/privacy",
  "termsOfUseUrl": "https://evolvegs.ca/terms",
  "mpnId": ""
}
```

`privacyUrl` and `termsOfUseUrl` are now populated — these were the blocking fields. `mpnId` is empty string. This is acceptable if Evolve Global Solutions is not enrolled in the Microsoft Partner Network; an empty `mpnId` will not block submission for ISVs without MPN enrollment. If MPN enrollment exists, populate with the real ID before submitting.

Ensure that `https://evolvegs.ca/privacy` and `https://evolvegs.ca/terms` are live public URLs at submission time — Microsoft validation crawls them.

---

### CRITICAL 3 — Hardcoded Tenant and Hosted-Service Artifacts [FIXED]

**Status:** FIXED

- `config/serve.json` — now uses `contoso.sharepoint.com` placeholder (was real tenant hostname)
- `infra/deployment.json` (Azure subscription IDs, Entra object IDs) — directory no longer exists in this repo
- `PreviewSettingsDialog.ts` — the generated cleanup script template contains `"bpmn-file-handler"` / `"azurewebsites.net"` as substring-match patterns for identifying stale old Azure File Handlers. This is intentional cleanup logic, not a hardcoded tenant identifier, and is acceptable.

Runtime values (`tenantId`, `configSiteUrl`) are injected from SPFx context at runtime — not hardcoded.

---

### WARNING — External Data Flow: diagrams.net Iframe [IMPROVED, still present]

**Status:** IMPROVED from previous — was missing admin consent; now has explicit `window.confirm()` gate.

`DrawioViewerDialog.ts:7` — `const DRAWIO_EMBED_ORIGIN = 'https://embed.diagrams.net'`

When `.drawio` rendering is enabled by the tenant admin, the Draw.io viewer loads the external `https://embed.diagrams.net` iframe and passes drawing XML to it via `postMessage`. This means SharePoint-hosted diagram content is transmitted to an external Lucid Software service.

`PreviewSettingsDialog.ts:219-228` now shows a `window.confirm()` dialog before the admin can enable the diagrams.net renderer:

> "Enabling diagrams.net will send drawing XML from SharePoint to the external diagrams.net service…"

**Still needed for submission:**
1. This data flow must be described in the `https://evolvegs.ca/privacy` privacy statement.
2. Include in Notes for Certification during submission: "The `.drawio` renderer is opt-in, disabled by default, and transfers drawing XML to embed.diagrams.net only when an admin explicitly enables it after confirming an in-product disclosure."
3. Consider whether `.drawio` is disabled by default in the `ClientSideInstance.xml` initial deployment — confirm the default settings list does not enable it at install time.

---

### WARNING — Runtime WASM Downloads from cdn.jsdelivr.net [NEW — v1.5.21]

**Status:** NEW finding introduced by IFC and STEP viewer additions.

Two viewers fetch WASM binaries at runtime from an external CDN:

- `IfcViewerDialog.ts:15` — `const WEB_IFC_CDN = 'https://cdn.jsdelivr.net/npm/web-ifc@0.0.77/'`  
  Fetches `web-ifc.wasm` (~7 MB) and `web-ifc-mt.wasm` on first IFC file open.

- `StepViewerDialog.ts:19` — `const OCCT_WASM_CDN = 'https://cdn.jsdelivr.net/npm/occt-import-js@0.0.23/dist/'`  
  Fetches `occt-import-js.wasm` (~15 MB) on first STEP file open.

**Implications:**

1. **Privacy/data flow disclosure:** Users' browsers make outbound requests to jsDelivr (a CDN operated by Cloudflare + Fastly) for file content. While the WASM files themselves contain no user data, the HTTP requests reveal IP address, timing, and (implicitly) that the user opened an IFC/STEP file. This must be disclosed.

2. **Blocked networks:** Enterprise customers with restrictive outbound network policies (common in manufacturing/automotive) may be unable to reach `cdn.jsdelivr.net`. The viewers will fail with no geometry / load error in those environments. The error message already says "Check that the browser can reach cdn.jsdelivr.net" which is good, but consider documenting this requirement in setup docs.

3. **Submission requirement:** Microsoft requires all external data flows to be disclosed in the privacy statement and Notes for Certification. Add: "The IFC (.ifc) and STEP (.stp/.step) viewers download rendering engine WASM binaries from cdn.jsdelivr.net (jsDelivr) on first use. No user file content is sent to jsDelivr."

4. **Long-term option:** Bundle the WASM files inside the sppkg's CDN assets so they serve from SharePoint's own `publiccdn.sharepointonline.com`. This eliminates the external dependency entirely. Trade-off: increases sppkg size significantly (~22 MB for both WASM files).

---

### WARNING — Plaintext License Key in SharePoint List [PERSISTS]

**Status:** PERSISTS from previous audit.

`PreviewSettingsDialog.ts:93` — license key stored/retrieved as a plain-text SharePoint list field. Visual masking (`type="password"`) does not protect it from list admin access or SharePoint API reads.

For AppSource submission, the primary concern is disclosure: the privacy statement must note that a license validation token is stored in a customer-controlled SharePoint list. This is not a blocking issue if documented, since the data is stored in the customer's own tenant.

If a future version moves to server-side license validation, this becomes moot.

---

### WARNING — Root-Site Configuration Assumption [PERSISTS]

**Status:** PERSISTS from previous audit.

`ClientSideInstance.xml:8` — `"configSiteUrl":""` means the initial deployed instance has no configured site. The code falls back to the tenant root site.

Microsoft's AppSource validation runs in a test tenant and validates behavior on non-root modern sites. If the fallback silently resolves to a URL that doesn't exist in the test tenant, the validator will see a broken admin experience.

**Remediation:** Document in setup docs that the tenant admin must open settings and save a `configSiteUrl` on first use, or deploy using PnP/tenant installer with the correct site URL populated in the properties.

---

### WARNING — Mixed Tenant-Wide Deployment Mechanisms [PARTIALLY IMPROVED]

**Status:** PARTIALLY IMPROVED — `elements.xml` is now empty (custom actions removed to prevent duplicates). `ClientSideInstance.xml` is the sole declarative activation path.

Remaining concern: `scripts/Install-FilePreviewTenant.ps1` also writes directly to the Tenant Wide Extensions list. This creates two activation paths for a tenant-wide deployment: the declarative `ClientSideInstance.xml` in the sppkg, and the imperative PowerShell script. If both run, the extension activates twice on every library page.

For AppSource, the recommended path is: `skipFeatureDeployment: true` + `ClientSideInstance.xml` only. The PowerShell script can remain as a utility but should not be part of the advertised installation flow if users are installing from the AppSource catalog button.

---

### OPTIMIZATION — Duplicated SVG Icon Strings [PERSISTS, expanded]

**Status:** PERSISTS. v1.5.21 added three more viewers (IFC, Mermaid, STEP), each with their own local icon constant blocks.

Each dialog defines its own `S(d)` helper and icon constants or inline SVG strings. There is now a `src/shared/icons.ts` file in the repo — if icons are consolidated there, all dialogs can import from a single source. This reduces bundle size marginally and eliminates potential drift between icon versions.

Not a store blocker. Low priority until viewer count stabilizes.

---

### OPTIMIZATION 2 — DOM Polling [FIXED]

**Status:** FIXED. `FilePreviewAdminApplicationCustomizer.ts` was rewritten to use `PlaceholderName.Bottom` with no DOM scanning. The 750ms polling loop and `MutationObserver` fallbacks are gone.

---

## 3. Summary Table

| # | Finding | Previous | v1.5.21 |
|---|---|---|---|
| C1 | DOM manipulation outside SPFx boundary | Critical | **Critical (partial fix — 5 dialogs + bpmnAssetStyles persist)** |
| C2 | Missing developer metadata (privacy/terms URLs) | Critical | **FIXED** |
| C3 | Hardcoded tenant/Azure artifacts | Critical | **FIXED** |
| W1 | diagrams.net external data flow | Warning | **Warning (improved — admin consent added; still needs privacy disclosure)** |
| W2 | Runtime WASM CDN download | — | **NEW Warning** |
| W3 | Plaintext license key in SharePoint list | Warning | **Warning (persists)** |
| W4 | Root-site config assumption | Warning | **Warning (persists)** |
| W5 | Mixed deployment mechanisms | Warning | **Warning (partially improved)** |
| W6 | Hosted Azure File Handler broad permissions | Warning | Out of scope — separate repo, optional feature |
| O1 | DOM polling every 750ms | Optimization | **FIXED** |
| O2 | Duplicated SVG icons | Optimization | **Optimization (persists, expanded to new dialogs)** |

## 4. Pre-Submission Checklist

- [ ] **Critical:** Resolve `makeFullViewport()` DOM boundary issue (Notes for Certification waiver, fullscreen API refactor, or accepted risk)
- [ ] **Critical:** Move `bpmnAssetStyles.ts` style injection inside `domElement`
- [ ] **Required:** Confirm `https://evolvegs.ca/privacy` and `https://evolvegs.ca/terms` are live
- [ ] **Required:** Add WASM CDN and diagrams.net data flows to privacy statement
- [ ] **Required:** Add Notes for Certification describing: (a) Fluent UI DOM fix approach, (b) diagrams.net opt-in flow, (c) WASM CDN downloads
- [ ] **Recommended:** Populate `mpnId` if enrolled in Microsoft Partner Network
- [ ] **Recommended:** Set `configSiteUrl` default in `ClientSideInstance.xml` or document first-run requirement
- [ ] **Optional:** Consolidate icons into `src/shared/icons.ts`
- [ ] **Optional:** Clarify whether `StepViewerIframeContent.ts` is dead code after the Three.js rewrite; remove if unused
