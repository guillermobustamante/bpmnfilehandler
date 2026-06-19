import { useCallback, useEffect, useRef, useState } from "react";
import BpmnModeler from "bpmn-js/lib/Modeler";
import BpmnViewer from "bpmn-js/lib/NavigatedViewer";
import { Download, ExternalLink, Maximize2, RefreshCw, Save, ZoomIn, ZoomOut } from "lucide-react";
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

  const editable = launch.action === "open";
  const itemUrl = launch.itemUrls[0];

  const loadFile = useCallback(async () => {
    setIsBusy(true);
    setError("");
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
        await instance.importXML(xml);
        if (disposed) {
          return;
        }

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
      setStatus("Saved");
    } catch (saveError) {
      setError(saveError instanceof Error ? saveError.message : "Could not save BPMN file.");
      setStatus("Error");
    } finally {
      setIsBusy(false);
    }
  }, [editable, getAccessToken, itemUrl, metadata]);

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
          <strong className="file__name">{metadata?.name || "BPMN diagram"}</strong>
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
      <main className="canvas-wrap">
        <div ref={canvasRef} className="bpmn-canvas" />
      </main>
    </div>
  );
}

