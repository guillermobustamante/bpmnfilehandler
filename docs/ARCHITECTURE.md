# BPMN Microsoft 365 File Handler Architecture

## Selected Pattern

The solution uses Microsoft 365 File Handler 2.0 for tenant-wide `.bpmn` integration in SharePoint Online and OneDrive for Business.

- `preview` opens the BPMN viewer/editor inside the SharePoint/OneDrive iframe, matching the native file preview surface.
- The `open` action is intentionally not registered because Microsoft 365 launches file-handler `open` actions in a new browser tab.
- Authentication uses Microsoft Entra ID delegated permissions only. The app acts as the signed-in user.
- File reads and writes use Microsoft Graph `DriveItem` metadata and `/content` endpoints.
- Rendering and editing use `bpmn-js`.

## Runtime Flow

1. SharePoint or OneDrive invokes `/filehandler/preview` with an `application/x-www-form-urlencoded` POST.
2. The Node host validates the activation payload and stores the selected Graph item URL in a short-lived in-memory launch cache.
3. The host redirects to `/launch/{action}/{launchId}`.
4. The React app loads runtime auth config and launch context.
5. MSAL signs in the user with a login hint from the activation payload and persists auth in browser local storage.
6. The app reads BPMN XML from Microsoft Graph, renders it, and saves back through delegated Graph access.

## Security Posture

- No application permissions are required.
- Launch IDs are random, short-lived, and never placed in persistent storage.
- The app validates launch item URLs before accepting them.
- Microsoft Graph calls are made from the browser using delegated user tokens.
- Save uses `If-Match` when an eTag is available, which catches conflicting file changes.

## Operational Note

The current launch cache is process memory. Keep Azure App Service at one instance for the first production release. If you scale out, replace the in-memory cache with Redis or another shared store.
