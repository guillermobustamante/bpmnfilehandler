import { useCallback, useEffect, useRef, useState } from "react";
import { Download, ExternalLink, RefreshCw, Save } from "lucide-react";
import type { LaunchContext } from "../App";
import { getDriveItemContent, getDriveItemMetadata, putDriveItemContent, type DriveItemMetadata } from "../graph/driveItem";
import { IconButton } from "./IconButton";

const drawioEmbedOrigin = "https://embed.diagrams.net";
const drawioEmbedUrl = `${drawioEmbedOrigin}/?embed=1&proto=json&spin=1&libraries=1&noSaveBtn=1&noExitBtn=1`;

type DrawioMessage = {
  event?: string;
  message?: string;
  xml?: string;
};

type WorkspaceProps = {
  getAccessToken: () => Promise<string>;
  launch: LaunchContext;
};

export function DrawioWorkspace({ getAccessToken, launch }: WorkspaceProps) {
  const iframeRef = useRef<HTMLIFrameElement | null>(null);
  const [metadata, setMetadata] = useState<DriveItemMetadata | null>(null);
  const [xml, setXml] = useState("");
  const [error, setError] = useState("");
  const [message, setMessage] = useState("");
  const [status, setStatus] = useState("Loading");
  const [isDirty, setIsDirty] = useState(false);
  const [isBusy, setIsBusy] = useState(true);

  const editable = true;
  const itemUrl = launch.itemUrls[0];

  const sendLoadMessage = useCallback(
    (loadedXml = xml) => {
      const target = iframeRef.current?.contentWindow;
      if (!target || !loadedXml) {
        return;
      }

      target.postMessage(
        JSON.stringify({
          action: "load",
          autosave: editable ? 1 : 0,
          modified: "unsavedChanges",
          saveAndExit: "0",
          title: metadata?.name || "diagram.drawio",
          xml: loadedXml
        }),
        drawioEmbedOrigin
      );
    },
    [editable, metadata?.name, xml]
  );

  const loadFile = useCallback(async () => {
    setIsBusy(true);
    setError("");
    setMessage("");
    setStatus("Loading");

    try {
      const token = await getAccessToken();
      const [loadedMetadata, loadedXml] = await Promise.all([
        getDriveItemMetadata(itemUrl, token),
        getDriveItemContent(itemUrl, token)
      ]);

      setMetadata(loadedMetadata);
      setXml(loadedXml);
      setIsDirty(false);
      setStatus(editable ? "Ready" : "Preview");
      window.setTimeout(() => sendLoadMessage(loadedXml), 0);
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : "Could not load DrawIO file.");
      setStatus("Error");
    } finally {
      setIsBusy(false);
    }
  }, [editable, getAccessToken, itemUrl, sendLoadMessage]);

  const saveXml = useCallback(
    async (nextXml: string) => {
      if (!editable || !metadata) {
        return;
      }

      setIsBusy(true);
      setError("");
      setStatus("Saving");

      try {
        const token = await getAccessToken();
        const savedMetadata = await putDriveItemContent(itemUrl, token, nextXml, metadata.eTag);
        setMetadata(savedMetadata);
        setXml(nextXml);
        setIsDirty(false);
        setStatus("Saved");
      } catch (saveError) {
        setError(saveError instanceof Error ? saveError.message : "Could not save DrawIO file.");
        setStatus("Error");
      } finally {
        setIsBusy(false);
      }
    },
    [editable, getAccessToken, itemUrl, metadata]
  );

  useEffect(() => {
    void loadFile();
  }, [loadFile]);

  useEffect(() => {
    function onMessage(event: MessageEvent) {
      if (event.origin !== drawioEmbedOrigin || event.source !== iframeRef.current?.contentWindow) {
        return;
      }

      const drawioMessage = parseDrawioMessage(event.data);
      if (!drawioMessage) {
        return;
      }

      switch (drawioMessage.event) {
        case "init":
          sendLoadMessage();
          break;
        case "autosave":
        case "save":
          if (typeof drawioMessage.xml === "string") {
            setXml(drawioMessage.xml);
            setIsDirty(true);
            setStatus("Unsaved");
            if (editable && drawioMessage.event === "save") {
              void saveXml(drawioMessage.xml);
            }
          }
          break;
        default:
          if (drawioMessage.message) {
            setMessage(drawioMessage.message);
          }
          break;
      }
    }

    window.addEventListener("message", onMessage);
    return () => window.removeEventListener("message", onMessage);
  }, [editable, saveXml, sendLoadMessage]);

  const requestSave = useCallback(() => {
    iframeRef.current?.contentWindow?.postMessage(JSON.stringify({ action: "save" }), drawioEmbedOrigin);
  }, []);

  const downloadFile = useCallback(() => {
    const blob = new Blob([xml], { type: "application/xml" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = metadata?.name || "diagram.drawio";
    anchor.click();
    URL.revokeObjectURL(url);
  }, [metadata?.name, xml]);

  return (
    <div className="shell">
      <header className="toolbar">
        <div className="file">
          <span className="file__type">DRAWIO</span>
          <div className="file__details">
            <strong className="file__name">{metadata?.name || "DrawIO drawing"}</strong>
            {metadata ? <span className="file__meta">{formatMetadata(metadata)}</span> : null}
          </div>
          <span className={`file__status ${status === "Error" ? "file__status--error" : ""}`}>{status}</span>
        </div>
        <div className="toolbar__actions">
          <IconButton disabled={isBusy} label="Reload" onClick={() => void loadFile()}>
            <RefreshCw size={18} />
          </IconButton>
          <IconButton disabled={isBusy} label="Save" onClick={requestSave}>
            <Save size={18} />
          </IconButton>
          <IconButton disabled={isBusy || !xml} label="Download" onClick={downloadFile}>
            <Download size={18} />
          </IconButton>
          {metadata?.webUrl ? (
            <IconButton label="Open in SharePoint" onClick={() => window.open(metadata.webUrl, "_blank", "noopener,noreferrer")}>
              <ExternalLink size={18} />
            </IconButton>
          ) : null}
        </div>
      </header>
      {error ? <div className="error-bar">{error}</div> : null}
      {message ? <div className="warning-bar">{message}</div> : null}
      <main className="canvas-wrap">
        <iframe
          ref={iframeRef}
          className="drawio-frame"
          title="diagrams.net editor"
          sandbox="allow-downloads allow-forms allow-popups allow-same-origin allow-scripts"
          src={drawioEmbedUrl}
        />
      </main>
    </div>
  );
}

function parseDrawioMessage(value: unknown): DrawioMessage | undefined {
  if (typeof value === "object" && value !== null) {
    return value as DrawioMessage;
  }

  if (typeof value !== "string") {
    return undefined;
  }

  try {
    return JSON.parse(value) as DrawioMessage;
  } catch {
    return undefined;
  }
}

function formatMetadata(metadata: DriveItemMetadata): string {
  const parts: string[] = [];
  if (metadata.lastModifiedDateTime) {
    parts.push(`Modified ${new Date(metadata.lastModifiedDateTime).toLocaleString()}`);
  }
  if (metadata.lastModifiedBy?.user?.displayName) {
    parts.push(`by ${metadata.lastModifiedBy.user.displayName}`);
  }
  if (typeof metadata.size === "number") {
    parts.push(formatBytes(metadata.size));
  }

  return parts.join(" | ");
}

function formatBytes(value: number): string {
  if (value < 1024) {
    return `${value} B`;
  }

  if (value < 1024 * 1024) {
    return `${Math.round(value / 102.4) / 10} KB`;
  }

  return `${Math.round(value / 1024 / 102.4) / 10} MB`;
}
