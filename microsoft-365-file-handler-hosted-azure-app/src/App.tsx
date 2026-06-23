import { useCallback, useEffect, useMemo, useState } from "react";
import {
  BrowserAuthError,
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo
} from "@azure/msal-browser";
import { LogIn, TriangleAlert } from "lucide-react";
import { DirectLauncher } from "./components/DirectLauncher";
import { FileWorkspace } from "./components/FileWorkspace";
import { IconButton } from "./components/IconButton";
import { createMsalClient, type PublicConfig } from "./auth/msal";
import { ManualLauncher } from "./components/ManualLauncher";

const pendingLaunchStorageKey = "bpmnFileHandler.pendingLaunch";
const pendingFileUrlStorageKey = "bpmnFileHandler.pendingFileUrl";
const pendingDirectLaunchStorageKey = "bpmnFileHandler.pendingDirectLaunch";

export type ViewerMode = "modeler" | "viewer";

export type DirectLaunchOptions = {
  extension?: string;
  fileUrl: string;
  mode?: ViewerMode;
};

export type LaunchContext = {
  id: string;
  action: "open" | "preview";
  cultureName?: string;
  client?: string;
  userId?: string;
  domainHint?: string;
  extension?: string;
  itemUrls: string[];
  mode?: ViewerMode;
  createdAt: string;
  expiresAt: string;
};

type AppState =
  | { kind: "loading" }
  | { kind: "error"; message: string }
  | {
      kind: "ready";
      config: PublicConfig;
      directLaunch: DirectLaunchOptions | null;
      launch: LaunchContext | null;
      msal: PublicClientApplication;
    };

export function App() {
  const [state, setState] = useState<AppState>({ kind: "loading" });
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [authMessage, setAuthMessage] = useState<string>("");

  const route = useMemo(() => parseLaunchRoute(window.location.pathname), []);
  const requestedDirectLaunch = useMemo(() => parseDirectLaunchOptions(window.location.search), []);

  useEffect(() => {
    let cancelled = false;

    async function bootstrap() {
      try {
        const [configResponse, launchResponse] = await Promise.all([
          fetch("/api/config", { cache: "no-store" }),
          route ? fetch(`/api/launch/${encodeURIComponent(route.launchId)}`, { cache: "no-store" }) : null
        ]);

        if (!configResponse.ok) {
          throw new Error("Could not load authentication configuration.");
        }
        if (launchResponse && !launchResponse.ok) {
          const details = await launchResponse.json().catch(() => ({}));
          throw new Error(details.error || "Could not load file handler launch context.");
        }

        const config = (await configResponse.json()) as PublicConfig;
        const routeLaunch = launchResponse ? ((await launchResponse.json()) as LaunchContext) : null;
        const launch = routeLaunch || (hasMsalAuthResponse() ? readPendingLaunch() : null);
        const directLaunch = requestedDirectLaunch || (hasMsalAuthResponse() ? readPendingDirectLaunch() : null);

        if (!config.configured) {
          throw new Error("The handler is deployed without M365_CLIENT_ID.");
        }

        const msal = createMsalClient(config);
        await msal.initialize();
        const redirectResponse = await msal.handleRedirectPromise({
          navigateToLoginRequestUrl: false
        });

        if (redirectResponse) {
          clearPendingLaunch();
          clearPendingDirectLaunch();
          clearPendingFileUrl();
          if (launch && !route) {
            window.history.replaceState(null, "", `/launch/${launch.action}/${launch.id}`);
          } else if (directLaunch && !requestedDirectLaunch) {
            window.history.replaceState(null, "", buildDirectLaunchPath(directLaunch));
          }
        }

        const existingAccount = redirectResponse?.account || pickAccount(msal.getAllAccounts(), launch?.userId);
        if (existingAccount) {
          msal.setActiveAccount(existingAccount);
        }

        if (!cancelled) {
          setAccount(existingAccount);
          setState({ kind: "ready", config, directLaunch, launch, msal });
        }
      } catch (error) {
        if (!cancelled) {
          setState({
            kind: "error",
            message: error instanceof Error ? error.message : "The BPMN handler could not start."
          });
        }
      }
    }

    void bootstrap();
    return () => {
      cancelled = true;
    };
  }, [requestedDirectLaunch, route]);

  useEffect(() => {
    if (state.kind !== "ready" || account) {
      return;
    }

    const readyState = state;
    let cancelled = false;
    async function trySilentSignIn() {
      try {
        const response = await readyState.msal.ssoSilent({
          scopes: readyState.config.scopes,
          loginHint: readyState.launch?.userId
        });

        if (!cancelled) {
          readyState.msal.setActiveAccount(response.account);
          setAccount(response.account);
          setAuthMessage("");
        }
      } catch {
        if (!cancelled) {
          setAuthMessage("Sign in to continue.");
        }
      }
    }

    void trySilentSignIn();
    return () => {
      cancelled = true;
    };
  }, [account, state]);

  const signIn = useCallback(async () => {
    if (state.kind !== "ready") {
      return;
    }

    setAuthMessage("");
    try {
      if (state.launch) {
        writePendingLaunch(state.launch);
      } else {
        clearPendingLaunch();
      }
      if (state.directLaunch) {
        writePendingDirectLaunch(state.directLaunch);
      } else {
        clearPendingDirectLaunch();
        clearPendingFileUrl();
      }

      const request = {
        scopes: state.config.scopes,
        loginHint: state.launch?.userId
      };

      if (isEmbedded()) {
        const response = await state.msal.loginPopup({
          ...request,
          redirectUri: `${window.location.origin}/auth.html`
        });
        state.msal.setActiveAccount(response.account);
        setAccount(response.account);
        clearPendingLaunch();
        clearPendingDirectLaunch();
        clearPendingFileUrl();
        setAuthMessage("");
        return;
      }

      await state.msal.loginRedirect({
        ...request,
        redirectUri: window.location.origin,
        redirectStartPage: window.location.href
      });
    } catch (error) {
      setAuthMessage(error instanceof Error ? error.message : "Sign-in did not complete.");
    }
  }, [state]);

  const getAccessToken = useCallback(async () => {
    if (state.kind !== "ready") {
      throw new Error("Authentication is not ready.");
    }

    const activeAccount = state.msal.getActiveAccount() || account;
    if (!activeAccount) {
      throw new InteractionRequiredAuthError("no_account", "A user account is required.");
    }

    try {
      const token = await state.msal.acquireTokenSilent({
        account: activeAccount,
        scopes: state.config.scopes
      });
      return token.accessToken;
    } catch (error) {
      if (error instanceof BrowserAuthError || error instanceof InteractionRequiredAuthError) {
        if (state.launch) {
          writePendingLaunch(state.launch);
        }
        if (state.directLaunch) {
          writePendingDirectLaunch(state.directLaunch);
        }

        if (isEmbedded()) {
          const token = await state.msal.acquireTokenPopup({
            account: activeAccount,
            scopes: state.config.scopes,
            redirectUri: `${window.location.origin}/auth.html`
          });
          state.msal.setActiveAccount(token.account);
          setAccount(token.account);
          clearPendingLaunch();
          clearPendingDirectLaunch();
          clearPendingFileUrl();
          return token.accessToken;
        }

        await state.msal.acquireTokenRedirect({
          account: activeAccount,
          scopes: state.config.scopes,
          redirectUri: window.location.origin,
          redirectStartPage: window.location.href
        });
        throw new Error("Redirecting for Microsoft 365 access.");
      }

      throw error;
    }
  }, [account, state]);

  if (state.kind === "loading") {
    return <CenteredMessage title="Loading" />;
  }

  if (state.kind === "error") {
    return <CenteredMessage title="BPMN handler unavailable" message={state.message} />;
  }

  if (!account) {
    return (
      <CenteredMessage title="Sign in" message={authMessage || state.launch?.userId || "Use your Microsoft 365 account."}>
        <IconButton label="Sign in" onClick={signIn}>
          <LogIn size={18} />
        </IconButton>
      </CenteredMessage>
    );
  }

  if (state.launch) {
    return <FileWorkspace getAccessToken={getAccessToken} launch={state.launch} />;
  }

  if (state.directLaunch) {
    return <DirectLauncher directLaunch={state.directLaunch} getAccessToken={getAccessToken} />;
  }

  return <ManualLauncher getAccessToken={getAccessToken} />;
}

function CenteredMessage({
  children,
  message,
  title
}: {
  children?: React.ReactNode;
  message?: string;
  title: string;
}) {
  return (
    <main className="centered">
      <div className="centered__panel">
        <TriangleAlert aria-hidden="true" size={24} />
        <h1>{title}</h1>
        {message ? <p>{message}</p> : null}
        {children ? <div className="centered__actions">{children}</div> : null}
      </div>
    </main>
  );
}

function parseLaunchRoute(pathname: string): { action: "open" | "preview"; launchId: string } | null {
  const match = pathname.match(/^\/launch\/(open|preview)\/([0-9a-f-]{36})$/i);
  if (!match) {
    return null;
  }

  return {
    action: match[1].toLowerCase() as "open" | "preview",
    launchId: match[2]
  };
}

function parseDirectLaunchOptions(search: string): DirectLaunchOptions | null {
  const params = new URLSearchParams(search);
  const fileUrl = params.get("fileUrl");
  if (!fileUrl) {
    return null;
  }

  const trimmed = fileUrl.trim();
  if (!trimmed) {
    return null;
  }

  const mode = params.get("mode") === "viewer" ? "viewer" : params.get("mode") === "modeler" ? "modeler" : undefined;
  const extension = normalizeExtension(params.get("extension") || "");
  return {
    extension: extension || undefined,
    fileUrl: trimmed,
    mode
  };
}

function pickAccount(accounts: AccountInfo[], loginHint?: string): AccountInfo | null {
  if (accounts.length === 0) {
    return null;
  }

  if (!loginHint) {
    return accounts[0];
  }

  const normalizedHint = loginHint.toLowerCase();
  return (
    accounts.find((candidate) => candidate.username.toLowerCase() === normalizedHint) ||
    accounts.find((candidate) => candidate.idTokenClaims?.login_hint === normalizedHint) ||
    accounts[0]
  );
}

function hasMsalAuthResponse(): boolean {
  const authResponse = `${window.location.search}${window.location.hash}`;
  return /(?:[?#&])(code|error|error_description)=/.test(authResponse);
}

function isEmbedded(): boolean {
  try {
    return window.self !== window.top;
  } catch {
    return true;
  }
}

function writePendingLaunch(launch: LaunchContext): void {
  sessionStorage.setItem(pendingLaunchStorageKey, JSON.stringify(launch));
}

function clearPendingLaunch(): void {
  sessionStorage.removeItem(pendingLaunchStorageKey);
}

function writePendingDirectLaunch(directLaunch: DirectLaunchOptions): void {
  sessionStorage.setItem(pendingDirectLaunchStorageKey, JSON.stringify(directLaunch));
  sessionStorage.setItem(pendingFileUrlStorageKey, directLaunch.fileUrl);
}

function clearPendingFileUrl(): void {
  sessionStorage.removeItem(pendingFileUrlStorageKey);
}

function clearPendingDirectLaunch(): void {
  sessionStorage.removeItem(pendingDirectLaunchStorageKey);
}

function readPendingDirectLaunch(): DirectLaunchOptions | null {
  const raw = sessionStorage.getItem(pendingDirectLaunchStorageKey);
  if (raw) {
    try {
      const parsed = JSON.parse(raw) as DirectLaunchOptions;
      if (parsed.fileUrl?.trim()) {
        return {
          extension: normalizeExtension(parsed.extension || "") || undefined,
          fileUrl: parsed.fileUrl.trim(),
          mode: parsed.mode === "viewer" || parsed.mode === "modeler" ? parsed.mode : undefined
        };
      }
    } catch {
      clearPendingDirectLaunch();
    }
  }

  const legacyFileUrl = sessionStorage.getItem(pendingFileUrlStorageKey)?.trim();
  return legacyFileUrl ? { fileUrl: legacyFileUrl } : null;
}

function readPendingLaunch(): LaunchContext | null {
  const raw = sessionStorage.getItem(pendingLaunchStorageKey);
  if (!raw) {
    return null;
  }

  try {
    const launch = JSON.parse(raw) as LaunchContext;
    if (!launch.id || !launch.action || !Array.isArray(launch.itemUrls)) {
      clearPendingLaunch();
      return null;
    }

    if (Date.parse(launch.expiresAt) <= Date.now()) {
      clearPendingLaunch();
      return null;
    }

    return launch;
  } catch {
    clearPendingLaunch();
    return null;
  }
}

function buildDirectLaunchPath(directLaunch: DirectLaunchOptions): string {
  const params = new URLSearchParams();
  params.set("fileUrl", directLaunch.fileUrl);
  if (directLaunch.mode) {
    params.set("mode", directLaunch.mode);
  }
  if (directLaunch.extension) {
    params.set("extension", directLaunch.extension);
  }

  return `/?${params.toString()}`;
}

function normalizeExtension(value: string): string {
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) {
    return "";
  }

  return trimmed.startsWith(".") ? trimmed : `.${trimmed}`;
}
