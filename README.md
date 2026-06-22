# BPMN Microsoft 365 File Handler

A Microsoft 365 File Handler 2.0 for previewing, editing, and saving `.bpmn` files from SharePoint Online and OneDrive for Business.

## Stack

- Azure App Service, Node.js 22, Express
- React, TypeScript, Vite
- Microsoft Entra ID and MSAL browser auth
- Microsoft Graph delegated file access
- bpmn-js viewer/modeler

## Local Development

```powershell
npm install
npm run dev
```

For local Vite development:

```powershell
npm run dev
npx vite --host 127.0.0.1
```

Set `.env` values from `.env.example` before connecting to Microsoft 365.

## Build

```powershell
npm run typecheck
npm run build
```

## Deployment

See [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md).

## License Compliance

See [docs/THIRD_PARTY_NOTICES.md](docs/THIRD_PARTY_NOTICES.md).
