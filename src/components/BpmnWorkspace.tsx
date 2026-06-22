import { useCallback, useEffect, useRef, useState } from "react";
import BpmnModeler from "bpmn-js/lib/Modeler";
import BpmnViewer from "bpmn-js/lib/NavigatedViewer";
import { CheckCircle2, Download, ExternalLink, Maximize2, RefreshCw, Save, ZoomIn, ZoomOut } from "lucide-react";
import type { LaunchContext } from "../App";
import { getDriveItemContent, getDriveItemMetadata, putDriveItemContent, type DriveItemMetadata } from "../graph/driveItem";
import { IconButton } from "./IconButton";

type BpmnInstance = InstanceType<typeof BpmnModeler> | InstanceType<typeof BpmnViewer>;

type WorkspaceProps = {
  getAccessToken: () => Promise<string>;
  launch: LaunchContext;
};

export function BpmnWorkspace({ getAccessToken, launch }: WorkspaceProps) {
  const canvasRef = useRef<HTMLDivElement | null>(null);
  const bpmnRef = useRef<BpmnInstance | null>(null);
  const [metadata, setMetadata] = useState<DriveItemMetadata | null>(null);
  const [xml, setXml] = useState<string>("");
  const [error, setError] = useState<string>("");
  const [status, setStatus] = useState<string>("Loading");
  const [isDirty, setIsDirty] = useState(false);
  const [isBusy, setIsBusy] = useState(true);
  const [warnings, setWarnings] = useState<string[]>([]);

  const editable = launch.mode ? launch.mode === "modeler" : launch.action === "open" || launch.action === "preview";
  const itemUrl = launch.itemUrls[0];

  const loadFile = useCallback(async () => {
    setIsBusy(true);
    setError("");
    setStatus("Loading");
    setWarnings([]);

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
    } catch (loadError) {
      setError(loadError instanceof Error ? loadError.message : "Could not load BPMN file.");
      setStatus("Error");
    } finally {
      setIsBusy(false);
    }
  }, [editable, getAccessToken, itemUrl]);

  useEffect(() => {
    void loadFile();
  }, [loadFile]);

  useEffect(() => {
    const container = canvasRef.current;
    if (!container || !xml) {
      return;
    }

    let disposed = false;
    setError("");
    container.innerHTML = "";

    const instance: BpmnInstance = editable
      ? new BpmnModeler({
          container,
          keyboard: { bindTo: window }
        })
      : new BpmnViewer({
          container
        });

    bpmnRef.current = instance;

    async function importDiagram() {
      try {
        const importResult = (await instance.importXML(xml)) as { warnings?: Array<Error | { message?: string }> };
        if (disposed) {
          return;
        }

        setWarnings(formatWarnings(importResult.warnings));

        const canvas = instance.get("canvas") as { zoom: (value: string | number, point?: string) => void };
        canvas.zoom("fit-viewport", "auto");

        if (editable) {
          const eventBus = instance.get("eventBus") as {
            on: (eventName: string, callback: () => void) => void;
          };
          eventBus.on("commandStack.changed", () => {
            setIsDirty(true);
            setStatus("Unsaved");
          });
        }
      } catch (importError) {
        setError(importError instanceof Error ? importError.message : "Could not render BPMN XML.");
        setStatus("Error");
      }
    }

    void importDiagram();

    return () => {
      disposed = true;
      bpmnRef.current = null;
      instance.destroy();
    };
  }, [editable, xml]);

  const saveFile = useCallback(async () => {
    const instance = bpmnRef.current;
    if (!editable || !metadata || !instance) {
      return;
    }

    setIsBusy(true);
    setError("");
    setStatus("Saving");

    try {
      const token = await getAccessToken();
      const result = await instance.saveXML({ format: true });
      const savedMetadata = await putDriveItemContent(itemUrl, token, result.xml || "", metadata.eTag);
      setMetadata(savedMetadata);
      setXml(result.xml || "");
      setIsDirty(false);
      setWarnings([]);
      setStatus("Saved");
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : "Could not save BPMN file.");
      setStatus("Error");
    } finally {
      setIsBusy(false);
    }
  }, [editable, getAccessToken, itemUrl, metadata]);

  const validateDiagram = useCallback(async () => {
    const instance = bpmnRef.current;
    if (!instance) {
      return;
    }

    setError("");
    try {
      const result = await instance.saveXML({ format: true });
      if (!result.xml?.trim()) {
        throw new Error("The diagram XML is empty.");
      }

      setWarnings([]);
      setStatus(isDirty ? "Unsaved" : "Valid");
    } catch (validationError) {
      setError(validationError instanceof Error ? validationError.message : "Could not validate BPMN XML.");
      setStatus("Error");
    }
  }, [isDirty]);

  const downloadFile = useCallback(async () => {
    const instance = bpmnRef.current;
    if (!instance) {
      return;
    }

    const result = await instance.saveXML({ format: true });
    const blob = new Blob([result.xml || xml], { type: "application/xml" });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = metadata?.name || "diagram.bpmn";
    anchor.click();
    URL.revokeObjectURL(url);
  }, [metadata?.name, xml]);

  const zoom = useCallback((value: string | number) => {
    const canvas = bpmnRef.current?.get("canvas") as { zoom: (value: string | number, point?: string) => void } | undefined;
    canvas?.zoom(value, "auto");
  }, []);

  return (
    <div className="shell">
      <header className="toolbar">
        <div className="file">
          <span className="file__type">BPMN</span>
          <div className="file__details">
            <strong className="file__name">{metadata?.name || "BPMN diagram"}</strong>
            {metadata ? <span className="file__meta">{formatMetadata(metadata)}</span> : null}
          </div>
          <span className={`file__status ${status === "Error" ? "file__status--error" : ""}`}>{status}</span>
        </div>
        <div className="toolbar__actions">
          <IconButton disabled={isBusy} label="Reload" onClick={() => void loadFile()}>
            <RefreshCw size={18} />
          </IconButton>
          {editable ? (
            <IconButton disabled={isBusy || !isDirty} label="Save" onClick={() => void saveFile()}>
              <Save size={18} />
            </IconButton>
          ) : null}
          <IconButton disabled={isBusy} label="Validate BPMN" onClick={() => void validateDiagram()}>
            <CheckCircle2 size={18} />
          </IconButton>
          <IconButton label="Download" onClick={() => void downloadFile()}>
            <Download size={18} />
          </IconButton>
          {metadata?.webUrl ? (
            <IconButton label="Open in SharePoint" onClick={() => window.open(metadata.webUrl, "_blank", "noopener,noreferrer")}>
              <ExternalLink size={18} />
            </IconButton>
          ) : null}
          <span className="toolbar__separator" />
          <IconButton label="Zoom out" onClick={() => zoom(0.8)}>
            <ZoomOut size={18} />
          </IconButton>
          <IconButton label="Zoom in" onClick={() => zoom(1.2)}>
            <ZoomIn size={18} />
          </IconButton>
          <IconButton label="Fit" onClick={() => zoom("fit-viewport")}>
            <Maximize2 size={18} />
          </IconButton>
        </div>
      </header>
      {error ? <div className="error-bar">{error}</div> : null}
      {warnings.length > 0 ? <div className="warning-bar">{warnings.join(" ")}</div> : null}
      <main className="canvas-wrap">
        <div ref={canvasRef} className="bpmn-canvas" />
      </main>
    </div>
  );
}

function formatWarnings(rawWarnings: Array<Error | { message?: string }> | undefined): string[] {
  if (!rawWarnings?.length) {
    return [];
  }

  return rawWarnings.slice(0, 3).map((warning) => warning.message || "BPMN import warning.");
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
