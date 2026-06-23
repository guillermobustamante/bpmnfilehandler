import type { LaunchContext } from "../App";
import { BpmnWorkspace } from "./BpmnWorkspace";
import { DrawioWorkspace } from "./DrawioWorkspace";

type FileWorkspaceProps = {
  getAccessToken: () => Promise<string>;
  launch: LaunchContext;
};

export function FileWorkspace({ getAccessToken, launch }: FileWorkspaceProps) {
  if (normalizeExtension(launch.extension || "") === ".drawio") {
    return <DrawioWorkspace getAccessToken={getAccessToken} launch={launch} />;
  }

  return <BpmnWorkspace getAccessToken={getAccessToken} launch={launch} />;
}

function normalizeExtension(value: string): string {
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) {
    return "";
  }

  return trimmed.startsWith(".") ? trimmed : `.${trimmed}`;
}
