import { useEffect, useState } from "react";
import { FolderOpen, TriangleAlert } from "lucide-react";
import type { DirectLaunchOptions, LaunchContext } from "../App";
import { resolveDriveItemInput } from "../graph/driveItem";
import { BpmnWorkspace } from "./BpmnWorkspace";

type DirectLauncherProps = {
  directLaunch: DirectLaunchOptions;
  getAccessToken: () => Promise<string>;
};

export function DirectLauncher({ directLaunch, getAccessToken }: DirectLauncherProps) {
  const [launch, setLaunch] = useState<LaunchContext | null>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    let cancelled = false;

    async function start() {
      setError("");

      try {
        const accessToken = await getAccessToken();
        const resolved = await resolveDriveItemInput(
          directLaunch.fileUrl,
          accessToken,
          directLaunch.extension ? [directLaunch.extension] : undefined
        );
        const now = Date.now();

        if (!cancelled) {
          setLaunch({
            id: crypto.randomUUID(),
            action: "preview",
            extension: directLaunch.extension,
            itemUrls: [resolved.itemUrl],
            mode: directLaunch.mode || "modeler",
            createdAt: new Date(now).toISOString(),
            expiresAt: new Date(now + 60 * 60 * 1000).toISOString(),
            client: "SharePoint Command Set"
          });
        }
      } catch (openError) {
        if (!cancelled) {
          setError(openError instanceof Error ? openError.message : "Could not open this BPMN file.");
        }
      }
    }

    void start();

    return () => {
      cancelled = true;
    };
  }, [directLaunch, getAccessToken]);

  if (launch) {
    return <BpmnWorkspace getAccessToken={getAccessToken} launch={launch} />;
  }

  return (
    <main className="centered">
      <div className="centered__panel">
        {error ? <TriangleAlert aria-hidden="true" size={24} /> : <FolderOpen aria-hidden="true" size={24} />}
        <h1>{error ? "BPMN handler unavailable" : "Opening BPMN"}</h1>
        <p>{error || "Loading the selected SharePoint file."}</p>
      </div>
    </main>
  );
}
