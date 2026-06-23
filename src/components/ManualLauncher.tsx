import { useCallback, useState } from "react";
import { ExternalLink, FolderOpen } from "lucide-react";
import type { LaunchContext } from "../App";
import { resolveDriveItemInput } from "../graph/driveItem";
import { FileWorkspace } from "./FileWorkspace";
import { IconButton } from "./IconButton";

type ManualLauncherProps = {
  getAccessToken: () => Promise<string>;
};

export function ManualLauncher({ getAccessToken }: ManualLauncherProps) {
  const [input, setInput] = useState("");
  const [launch, setLaunch] = useState<LaunchContext | null>(null);
  const [error, setError] = useState("");
  const [isBusy, setIsBusy] = useState(false);

  const start = useCallback(async () => {
    setIsBusy(true);
    setError("");

    try {
      const accessToken = await getAccessToken();
      const resolved = await resolveDriveItemInput(input, accessToken, [".bpmn", ".drawio"]);
      const now = Date.now();

      setLaunch({
        id: crypto.randomUUID(),
        action: "open",
        extension: getFileExtension(resolved.metadata.name),
        itemUrls: [resolved.itemUrl],
        mode: getFileExtension(resolved.metadata.name) === ".drawio" ? "modeler" : undefined,
        createdAt: new Date(now).toISOString(),
        expiresAt: new Date(now + 60 * 60 * 1000).toISOString(),
        client: "Manual"
      });
    } catch (openError) {
      setError(openError instanceof Error ? openError.message : "Could not open this file.");
    } finally {
      setIsBusy(false);
    }
  }, [getAccessToken, input]);

  if (launch) {
    return <FileWorkspace getAccessToken={getAccessToken} launch={launch} />;
  }

  return (
    <main className="manual">
      <section className="manual__panel">
        <div className="manual__heading">
          <FolderOpen aria-hidden="true" size={24} />
          <h1>File handler manual test</h1>
        </div>
        <label className="manual__label" htmlFor="manual-url">
          SharePoint or OneDrive file link
        </label>
        <div className="manual__row">
          <input
            id="manual-url"
            className="manual__input"
            disabled={isBusy}
            onChange={(event) => setInput(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") {
                void start();
              }
            }}
            placeholder="https://evolvegs.sharepoint.com/..."
            value={input}
          />
          <IconButton disabled={isBusy || !input.trim()} label="Open file" onClick={() => void start()}>
            <ExternalLink size={18} />
          </IconButton>
        </div>
        {error ? <p className="manual__error">{error}</p> : null}
      </section>
    </main>
  );
}

function getFileExtension(fileName: string): string {
  const match = fileName.toLowerCase().match(/\.[^.]+$/);
  return match?.[0] || "";
}
