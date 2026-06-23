# SPFx SharePoint In-Page Preview

## Summary

SharePoint Framework package for opening supported process and diagram files directly inside SharePoint document libraries. This project is the AppSource candidate and is intentionally separated from the hosted Microsoft 365 File Handler / Azure app.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js 22.x, matching `package.json`.
- SharePoint Framework 1.23 toolchain through the local project dependencies.
- SharePoint tenant App Catalog for package deployment.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| SPFx SharePoint In-Page Preview | Evolve Global Solutions |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.4.18  | June 23, 2026    | Deployment package refresh |
| 1.4.17  | June 23, 2026    | Split from hosted File Handler app |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Local Development

```powershell
cd spfx-sharepoint-in-page-preview
npm install
npm run start
```

## Build

```powershell
npm run build
```

## Features

- ListView Command Set for document library preview actions.
- Admin web part and application customizer for tenant preview settings.
- SharePoint REST file read/write using the current user's permissions.
- BPMN preview/modeling through bundled `bpmn-js`.

## AppSource Audit

See `docs/APPSOURCE_COMPLIANCE_AUDIT.md`.

## Notes

The hosted Microsoft 365 File Handler / Azure app is now in `../microsoft-365-file-handler-hosted-azure-app/` and should not be included in the AppSource SPFx submission package.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [Heft Documentation](https://heft.rushstack.io/)
