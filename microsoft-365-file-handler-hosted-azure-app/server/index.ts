import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import express, { type Request, type Response } from "express";
import helmet from "helmet";

type FileHandlerAction = "open" | "preview";

type LaunchContext = {
  id: string;
  action: FileHandlerAction;
  cultureName?: string;
  client?: string;
  userId?: string;
  domainHint?: string;
  extension?: string;
  itemUrls: string[];
  mode?: "modeler" | "viewer";
  createdAt: string;
  expiresAt: string;
};

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const app = express();
const port = Number(process.env.PORT || 3000);
const launchTtlMs = Number(process.env.LAUNCH_TTL_MS || 15 * 60 * 1000);
const launches = new Map<string, LaunchContext>();

app.disable("x-powered-by");

app.use(
  helmet({
    frameguard: false,
    crossOriginEmbedderPolicy: false,
    crossOriginOpenerPolicy: { policy: "same-origin-allow-popups" },
    contentSecurityPolicy: {
      useDefaults: true,
      directives: {
        "default-src": ["'self'"],
        "script-src": ["'self'"],
        "style-src": ["'self'", "'unsafe-inline'"],
        "img-src": ["'self'", "data:", "https:"],
        "font-src": ["'self'", "data:"],
        "connect-src": [
          "'self'",
          "https://graph.microsoft.com",
          "https://login.microsoftonline.com",
          "https://*.sharepoint.com",
          "https://*.sharepoint-df.com",
          "https://*.1drv.com"
        ],
        "frame-src": ["'self'", "https://login.microsoftonline.com", "https://embed.diagrams.net"],
        "frame-ancestors": [
          "'self'",
          "https://*.sharepoint.com",
          "https://*.sharepoint-df.com",
          "https://*.onedrive.com",
          "https://*.office.com",
          "https://*.microsoft365.com"
        ],
        "object-src": ["'none'"]
      }
    },
    referrerPolicy: { policy: "no-referrer" }
  })
);

app.use((_, res, next) => {
  res.setHeader("X-Content-Type-Options", "nosniff");
  next();
});

app.get("/healthz", (_, res) => {
  res.status(200).json({ status: "ok" });
});

app.get("/api/config", (_, res) => {
  res.setHeader("Cache-Control", "no-store");

  const tenantId = process.env.M365_TENANT_ID || "organizations";
  const clientId = process.env.M365_CLIENT_ID || "";
  const scopes = (process.env.M365_GRAPH_SCOPES || "User.Read Files.ReadWrite.All")
    .split(/[,\s]+/)
    .map((scope) => scope.trim())
    .filter(Boolean);

  res.json({
    appName: process.env.APP_DISPLAY_NAME || "BPMN File Handler",
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientId,
    configured: Boolean(clientId),
    scopes,
    tenantId
  });
});

app.post(
  "/filehandler/:action",
  express.urlencoded({ extended: false, limit: "128kb" }),
  (req: Request<{ action: string }>, res: Response) => {
    const action = req.params.action;
    if (action !== "open" && action !== "preview") {
      res.status(404).send("Unsupported file handler action.");
      return;
    }

    let itemUrls: string[];
    try {
      itemUrls = parseItemUrls(req.body.items);
    } catch (error) {
      res.status(400).send(error instanceof Error ? error.message : "Invalid launch payload.");
      return;
    }

    if (itemUrls.length !== 1) {
      res.status(400).send("This handler currently supports one file at a time.");
      return;
    }

    const id = crypto.randomUUID();
    const now = Date.now();
    const launch: LaunchContext = {
      id,
      action,
      cultureName: readFormValue(req.body.cultureName),
      client: readFormValue(req.body.client),
      userId: readFormValue(req.body.userId),
      domainHint: readFormValue(req.body.domainHint),
      extension: normalizeExtension(readFormValue(req.query.extension)),
      itemUrls,
      mode: readMode(req.query.mode),
      createdAt: new Date(now).toISOString(),
      expiresAt: new Date(now + launchTtlMs).toISOString()
    };

    launches.set(id, launch);
    res.setHeader("Cache-Control", "no-store");
    res.redirect(303, `/launch/${action}/${id}`);
  }
);

app.get("/api/launch/:id", (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  const launch = launches.get(req.params.id);

  if (!launch) {
    res.status(404).json({ error: "Launch context was not found or has expired." });
    return;
  }

  if (Date.parse(launch.expiresAt) <= Date.now()) {
    launches.delete(req.params.id);
    res.status(410).json({ error: "Launch context has expired." });
    return;
  }

  res.json(launch);
});

const clientRoot = path.resolve(__dirname, "../client");
app.use(express.static(clientRoot, { index: false, maxAge: "1h" }));

app.use((req, res, next) => {
  if (req.method !== "GET" || !req.accepts("html")) {
    next();
    return;
  }

  res.setHeader("Cache-Control", "no-store");
  res.sendFile(path.join(clientRoot, "index.html"));
});

const cleanup = setInterval(() => {
  const now = Date.now();
  for (const [id, launch] of launches) {
    if (Date.parse(launch.expiresAt) <= now) {
      launches.delete(id);
    }
  }
}, 60_000);
cleanup.unref();

app.listen(port, () => {
  console.log(`BPMN file handler listening on port ${port}`);
});

function readFormValue(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

function readMode(value: unknown): "modeler" | "viewer" | undefined {
  const mode = readFormValue(value);
  return mode === "modeler" || mode === "viewer" ? mode : undefined;
}

function normalizeExtension(value: string | undefined): string | undefined {
  const trimmed = value?.trim().toLowerCase() || "";
  if (!trimmed) {
    return undefined;
  }

  return trimmed.startsWith(".") ? trimmed : `.${trimmed}`;
}

function parseItemUrls(itemsValue: unknown): string[] {
  if (typeof itemsValue !== "string" || itemsValue.trim().length === 0) {
    throw new Error("File handler launch did not include any selected item URLs.");
  }

  const parsed = JSON.parse(itemsValue) as unknown;
  if (!Array.isArray(parsed) || parsed.some((item) => typeof item !== "string")) {
    throw new Error("File handler launch items must be a JSON string array.");
  }

  const urls = parsed.map((item) => item.trim()).filter(Boolean);
  for (const url of urls) {
    assertAllowedGraphUrl(url);
  }

  return urls;
}

function assertAllowedGraphUrl(rawUrl: string): void {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch {
    throw new Error("File handler launch item URL is not valid.");
  }

  if (url.protocol !== "https:" || url.hostname !== "graph.microsoft.com") {
    throw new Error("File handler launch item URL must target Microsoft Graph.");
  }

  if (!url.pathname.startsWith("/v1.0/")) {
    throw new Error("File handler launch item URL must use Microsoft Graph v1.0.");
  }
}
