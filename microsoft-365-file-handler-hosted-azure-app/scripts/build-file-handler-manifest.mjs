import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

const appBaseUrl = normalizeBaseUrl(process.argv[2] || process.env.APP_BASE_URL);
if (!appBaseUrl) {
  console.error("Usage: npm run manifest -- https://your-app.azurewebsites.net [outputPath] [.bpmn,.drawio] [iconBaseUrl]");
  process.exit(1);
}

const outputPath = process.argv[3] || "dist/file-handler.addins.json";
const requestedExtensions = parseExtensions(process.argv[4] || process.env.FILE_HANDLER_EXTENSIONS || ".bpmn,.drawio");
const iconBaseUrl = normalizeBaseUrl(process.argv[5] || process.env.FILE_HANDLER_ICON_BASE_URL || `${appBaseUrl}/assets`);
const idFilePath = path.resolve("infra/file-handler-ids.json");
const legacyIdFilePath = path.resolve("infra/file-handler-id.txt");
const handlerIds = readStableFileHandlerIds(idFilePath, legacyIdFilePath, requestedExtensions);

const manifest = requestedExtensions.map((extension) => buildFileHandler(extension, handlerIds[extension], appBaseUrl, iconBaseUrl));

fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, `${JSON.stringify(manifest, null, 2)}\n`, "utf8");
console.log(`Wrote ${outputPath}`);
for (const handler of manifest) {
  const extension = getPropertyValue(handler.properties, "fileTypeDisplayName");
  console.log(`File Handler ID: ${handler.id} (${extension})`);
}

function buildFileHandler(extension, fileHandlerId, baseUrl, iconsBaseUrl) {
  const spec = getExtensionSpec(extension);
  const fileIcon = JSON.stringify(buildIconSet(iconsBaseUrl, spec.assetPrefix, "file"));
  const appIcon = JSON.stringify(buildIconSet(iconsBaseUrl, spec.assetPrefix, "app"));
  const actions = JSON.stringify([
    {
      type: "preview",
      url: `${baseUrl}/filehandler/preview?extension=${encodeURIComponent(spec.extension)}&mode=${spec.previewMode}`,
      availableOn: {
        file: { extensions: [spec.extension] },
        web: {}
      }
    },
    {
      type: "open",
      url: `${baseUrl}/filehandler/open?extension=${encodeURIComponent(spec.extension)}&mode=${spec.openMode}`,
      displayName: spec.openLabel,
      shortDisplayName: spec.openLabel,
      availableOn: {
        file: { extensions: [spec.extension] },
        web: {}
      }
    }
  ]);

  return {
    id: fileHandlerId,
    type: "FileHandler",
    properties: [
      { key: "version", value: "2" },
      { key: "fileTypeDisplayName", value: spec.fileTypeDisplayName },
      { key: "actionMenuDisplayName", value: spec.actionMenuDisplayName },
      { key: "fileTypeIcon", value: fileIcon },
      { key: "appIcon", value: appIcon },
      { key: "actions", value: actions }
    ]
  };
}

function getExtensionSpec(extension) {
  const specs = {
    ".bpmn": {
      extension: ".bpmn",
      assetPrefix: "bpmn",
      fileTypeDisplayName: "BPMN process diagram",
      actionMenuDisplayName: "Open BPMN",
      openLabel: "Open BPMN",
      openMode: "modeler",
      previewMode: "viewer"
    },
    ".drawio": {
      extension: ".drawio",
      assetPrefix: "drawio",
      fileTypeDisplayName: "DrawIO diagram",
      actionMenuDisplayName: "Open DrawIO",
      openLabel: "Open DrawIO",
      openMode: "modeler",
      previewMode: "viewer"
    }
  };

  const spec = specs[extension];
  if (!spec) {
    throw new Error(`No File Handler renderer/icon mapping exists for ${extension}. Supported: ${Object.keys(specs).join(", ")}`);
  }

  return spec;
}

function buildIconSet(baseUrl, assetPrefix, iconType) {
  return {
    svg: `${baseUrl}/${assetPrefix}-${iconType}.svg`,
    png1x: `${baseUrl}/${assetPrefix}-${iconType}-32.png`,
    "png1.5x": `${baseUrl}/${assetPrefix}-${iconType}-48.png`,
    png2x: `${baseUrl}/${assetPrefix}-${iconType}-64.png`
  };
}

function parseExtensions(value) {
  const extensions = value
    .split(",")
    .map((extension) => normalizeExtension(extension))
    .filter(Boolean);
  return Array.from(new Set(extensions));
}

function normalizeExtension(value) {
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) {
    return "";
  }

  return trimmed.startsWith(".") ? trimmed : `.${trimmed}`;
}

function normalizeBaseUrl(value) {
  if (!value) {
    return "";
  }

  return value.replace(/\/+$/, "");
}

function readStableFileHandlerIds(idPath, legacyIdPath, extensions) {
  const existing = readJsonObject(idPath);
  if (!existing[".bpmn"] && fs.existsSync(legacyIdPath)) {
    const legacyId = fs.readFileSync(legacyIdPath, "utf8").trim();
    if (legacyId) {
      existing[".bpmn"] = legacyId;
    }
  }

  let changed = false;
  for (const extension of extensions) {
    if (!existing[extension]) {
      existing[extension] = crypto.randomUUID();
      changed = true;
    }
  }

  if (changed || !fs.existsSync(idPath)) {
    fs.mkdirSync(path.dirname(idPath), { recursive: true });
    fs.writeFileSync(idPath, `${JSON.stringify(existing, null, 2)}\n`, "utf8");
  }

  return existing;
}

function readJsonObject(filePath) {
  if (!fs.existsSync(filePath)) {
    return {};
  }

  try {
    const parsed = JSON.parse(fs.readFileSync(filePath, "utf8"));
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  } catch {
    return {};
  }
}

function getPropertyValue(properties, key) {
  return properties.find((property) => property.key === key)?.value || key;
}
