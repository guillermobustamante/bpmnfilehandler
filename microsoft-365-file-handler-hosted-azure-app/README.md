# Microsoft 365 File Handler / Hosted Azure App

Original Microsoft 365 File Handler 2.0 implementation for opening `.bpmn` and `.drawio` files from SharePoint Online and OneDrive for Business through a hosted Azure App Service.

This project is intentionally separated from the SPFx AppSource candidate because it uses Entra app registration, MSAL browser auth, delegated Microsoft Graph scopes, Azure hosting, and Microsoft 365 File Handler `addIns`.

## Stack

- Azure App Service, Node.js 22, Express
- React, TypeScript, Vite
- Microsoft Entra ID and MSAL browser auth
- Microsoft Graph delegated file access
- bpmn-js viewer/modeler
- diagrams.net embedded renderer for DrawIO files

## Local Development

```powershell
cd microsoft-365-file-handler-hosted-azure-app
npm install
npm run dev
```

Set `.env` values from `.env.example` before connecting to Microsoft 365.

## Build

```powershell
npm run typecheck
npm run build
```

## Deployment

See `docs/DEPLOYMENT.md`.

## License Compliance

See `docs/THIRD_PARTY_NOTICES.md`.
