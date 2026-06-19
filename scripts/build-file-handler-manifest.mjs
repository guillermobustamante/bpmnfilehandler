import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

const appBaseUrl = normalizeBaseUrl(process.argv[2] || process.env.APP_BASE_URL);
if (!appBaseUrl) {
  console.error("Usage: npm run manifest -- https://your-app.azurewebsites.net");
  process.exit(1);
}

const outputPath = process.argv[3] || "dist/file-handler.addins.json";
const fileHandlerId = process.env.FILE_HANDLER_ID || crypto.randomUUID();
const icons = JSON.stringify({
  svg: `${appBaseUrl}/assets/bpmn-file.svg`
});

const appIcons = JSON.stringify({
  svg: `${appBaseUrl}/assets/bpmn-app.svg`
});

const actions = JSON.stringify([
  {
    type: "open",
    url: `${appBaseUrl}/filehandler/open`,
    availableOn: {
      file: { extensions: [".bpmn"] },
      web: {}
    }
  },
  {
    type: "preview",
    url: `${appBaseUrl}/filehandler/preview`,
    availableOn: {
      file: { extensions: [".bpmn"] },
      web: {}
    }
  }
]);

const manifest = [
  {
    id: fileHandlerId,
    type: "FileHandler",
    properties: [
      { key: "version", value: "2" },
      { key: "fileTypeDisplayName", value: "BPMN Diagram" },
      { key: "actionMenuDisplayName", value: "BPMN" },
      { key: "fileTypeIcon", value: icons },
      { key: "appIcon", value: appIcons },
      { key: "actions", value: actions }
    ]
  }
];

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
console.log(`Wrote ${outputPath}`);
console.log(`File Handler ID: ${fileHandlerId}`);

function normalizeBaseUrl(value) {
  if (!value) {
    return "";
  }

  return value.replace(/\/+$/, "");
}

