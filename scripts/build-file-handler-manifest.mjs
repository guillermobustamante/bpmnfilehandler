import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

const appBaseUrl = normalizeBaseUrl(process.argv[2] || process.env.APP_BASE_URL);
if (!appBaseUrl) {
  console.error("Usage: npm run manifest -- https://your-app.azurewebsites.net");
  process.exit(1);
}

const outputPath = process.argv[3] || "dist/file-handler.addins.json";
const idFilePath = path.resolve("infra/file-handler-id.txt");
const fileHandlerId = process.env.FILE_HANDLER_ID || readStableFileHandlerId(idFilePath);
const icons = JSON.stringify({
  svg: `${appBaseUrl}/assets/bpmn-file.svg`,
  png1x: `${appBaseUrl}/assets/bpmn-file-32.png`,
  "png1.5x": `${appBaseUrl}/assets/bpmn-file-48.png`,
  png2x: `${appBaseUrl}/assets/bpmn-file-64.png`
});

const appIcons = JSON.stringify({
  svg: `${appBaseUrl}/assets/bpmn-app.svg`,
  png1x: `${appBaseUrl}/assets/bpmn-app-32.png`,
  "png1.5x": `${appBaseUrl}/assets/bpmn-app-48.png`,
  png2x: `${appBaseUrl}/assets/bpmn-app-64.png`
});

const actions = JSON.stringify([
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

function readStableFileHandlerId(idPath) {
  if (fs.existsSync(idPath)) {
    return fs.readFileSync(idPath, "utf8").trim();
  }

  const generatedId = crypto.randomUUID();
  fs.mkdirSync(path.dirname(idPath), { recursive: true });
  fs.writeFileSync(idPath, `${generatedId}\n`, "utf8");
  return generatedId;
}
