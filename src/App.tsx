import { useCallback, useEffect, useMemo, useState } from "react";
import {
  BrowserAuthError,
  InteractionRequiredAuthError,
  PublicClientApplication,
  type AccountInfo
} from "@azure/msal-browser";
import { LogIn, TriangleAlert } from "lucide-react";
import { BpmnWorkspace } from "./components/BpmnWorkspace";
import { IconButton } from "./components/IconButton";
import { createMsalClient, type PublicConfig } from "./auth/msal";

export type LaunchContext = {
  id: string;
  action: "open" | "preview";
  cultureName?: string;
  client?: string;
  userId?: string;
  domainHint?: string;
  itemUrls: string[];
  createdAt: string;
  expiresAt: string;
};

type AppState =
  | { kind: "loading" }
  | { kind: "error"; message: string }
  | { kind: "ready"; config: PublicConfig; launch: LaunchContext; msal: PublicClientApplication };

export function App() {
  const [state, setState] = useState<AppState>({ kind: "loading" });
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [authMessage, setAuthMessage] = useState<string>("");

  const route = useMemo(() => parseLaunchRoute(window.location.pathname), []);

  useEffect(() => {
    let cancelled = false;

    async function bootstrap() {
      try {
        if (!route) {
          setState({
            kind: "error",
            message: "Open a BPMN file from SharePoint or OneDrive to start a handler session."
          });
          return;
        }

        const [configResponse, launchResponse] = await Promise.all([
          fetch("/api/config", { cache: "no-store" }),
          fetch(`/api/launch/${encodeURIComponent(route.launchId)}`, { cache: "no-store" })
        ]);

        if (!configResponse.ok) {
          throw new Error("Could not load authentication configuration.");
        }
        if (!launchResponse.ok) {
          const details = await launchResponse.json().catch(() => ({}));
          throw new Error(details.error || "Could not load file handler launch context.");
        }

        const config = (await configResponse.json()) as PublicConfig;
        const launch = (await launchResponse.json()) as LaunchContext;

        if (!config.configured) {
          throw new Error("The handler is deployed without M365_CLIENT_ID.");
        }

        const msal = createMsalClient(config);
        await msal.initialize();
        await msal.handleRedirectPromise();

        const existingAccount = pickAccount(msal.getAllAccounts(), launch.userId);
        if (existingAccount) {
          msal.setActiveAccount(existingAccount);
        }

        if (!cancelled) {
          setAccount(existingAccount);
          setState({ kind: "ready", config, launch, msal });
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
  }, [route]);

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
          loginHint: readyState.launch.userId
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
      const response = await state.msal.loginPopup({
        scopes: state.config.scopes,
        loginHint: state.launch.userId,
        prompt: "select_account"
      });

      state.msal.setActiveAccount(response.account);
      setAccount(response.account);
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
        const token = await state.msal.acquireTokenPopup({
          account: activeAccount,
          scopes: state.config.scopes
        });
        return token.accessToken;
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
      <CenteredMessage title="Sign in" message={authMessage || state.launch.userId}>
        <IconButton label="Sign in" onClick={signIn}>
          <LogIn size={18} />
        </IconButton>
      </CenteredMessage>
    );
  }

  return <BpmnWorkspace getAccessToken={getAccessToken} launch={state.launch} />;
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
