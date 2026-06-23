# BPMN File Preview Workspace

This repository is split into two separate projects with different compliance boundaries.

## Projects

### SPFx SharePoint in-page preview

Path: `spfx-sharepoint-in-page-preview/`

This is the AppSource candidate package. It contains the SharePoint Framework command set, application customizer, admin web part, SharePoint assets, package manifest, and AppSource audit notes. It should remain independent from the Microsoft 365 File Handler hosted app.

```powershell
cd spfx-sharepoint-in-page-preview
npm install
npm run build
```

### Microsoft 365 File Handler / hosted Azure app

Path: `microsoft-365-file-handler-hosted-azure-app/`

This is the original hosted File Handler implementation. It contains the Azure App Service Node/Express server, React/Vite app, MSAL browser auth, Microsoft Graph file access, Entra app registration scripts, and File Handler manifest tooling.

```powershell
cd microsoft-365-file-handler-hosted-azure-app
npm install
npm run dev
```

## Audit

The AppSource compliance audit is saved at `spfx-sharepoint-in-page-preview/docs/APPSOURCE_COMPLIANCE_AUDIT.md`.
