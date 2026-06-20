import { PublicClientApplication, type Configuration } from "@azure/msal-browser";

export type PublicConfig = {
  appName: string;
  authority: string;
  clientId: string;
  configured: boolean;
  scopes: string[];
  tenantId: string;
};

export function createMsalClient(config: PublicConfig): PublicClientApplication {
  const msalConfig: Configuration = {
    auth: {
      authority: config.authority,
      clientId: config.clientId,
      redirectUri: window.location.origin
    },
    cache: {
      cacheLocation: "sessionStorage"
    },
    system: {
      allowPlatformBroker: false
    }
  };

  return new PublicClientApplication(msalConfig);
}
