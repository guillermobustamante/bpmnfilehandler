# Deployment

## Prerequisites

- Azure CLI signed in with rights to create Azure resources and Entra applications.
- Tenant admin or application admin rights to update the Entra app manifest `addIns` property.
- SharePoint admin rights to run the tenant file handler cache refresh.
- Node.js 22 for local build validation.

## One-command Deployment

```powershell
.\scripts\deploy-all.ps1 `
  -SubscriptionId "<subscription-id>" `
  -TenantId "<tenant-id>" `
  -TenantHostName "<tenant>.sharepoint.com" `
  -ResourceGroupName "rg-bpmn-file-handler" `
  -Location "canadacentral" `
  -AppName "<globally-unique-app-name>"
```

The script creates:

- Azure App Service on Linux Node 22.
- Entra ID single-page application registration.
- Delegated Microsoft Graph scopes: `User.Read`, `Files.ReadWrite.All`.
- Tenant admin consent for those delegated scopes.
- Microsoft 365 File Handler 2.0 `addIns` registration for `.bpmn` and `.drawio`.
- Per-extension icon URLs for BPMN and DrawIO file/app icons.
- Native File Handler `preview` and `open` actions, so direct file-name click and the SharePoint/OneDrive Open menu can route to the handler.
- Tenant-wide File Handler cache refresh.

The App Service deployment package is built locally and includes the compiled app plus runtime dependencies. Remote Oryx build is disabled for deterministic deployment.

## Manual Deployment Steps

1. Create the app registration:

```powershell
.\scripts\create-entra-app.ps1 `
  -DisplayName "BPMN File Handler" `
  -AppBaseUrl "https://<app-name>.azurewebsites.net" `
  -TenantId "<tenant-id>"
```

Add `-SkipAdminConsent` only if the tenant requires consent to be granted in a separate approval workflow.

2. Deploy Azure App Service:

```powershell
.\scripts\deploy-azure-appservice.ps1 `
  -SubscriptionId "<subscription-id>" `
  -ResourceGroupName "rg-bpmn-file-handler" `
  -Location "canadacentral" `
  -AppName "<app-name>" `
  -ClientId "<app-registration-client-id>" `
  -TenantId "<tenant-id>"
```

3. Register the File Handler:

```powershell
.\scripts\register-file-handler.ps1 `
  -ApplicationObjectId "<app-registration-object-id>" `
  -AppBaseUrl "https://<app-name>.azurewebsites.net" `
  -Extensions ".bpmn",".drawio"
```

4. Refresh the tenant File Handler cache:

```powershell
.\scripts\refresh-file-handler-cache.ps1 -TenantHostName "<tenant>.sharepoint.com"
```

Microsoft caches File Handler registrations aggressively, so tenant-wide visibility can still take time after a successful registration. Some tenants reject Azure CLI's SharePoint token for this endpoint with `invalidScope`; in that case, the registration is still valid and should propagate naturally.
