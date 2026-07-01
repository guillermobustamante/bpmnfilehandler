'use strict';
// Copies WASM binaries and their license files into the webpack output directories
// (dist/ and release/assets/) so they are included in the sppkg and served from
// SharePoint's Office 365 CDN.
// Run after 'heft build --production' and before 'heft package-solution'.
// License notes:
//   web-ifc:        MPL-2.0  — freely distributable with attribution
//   occt-import-js: LGPL-2.1 — distributable in closed-source; .wasm must remain a
//                              separate replaceable file (satisfied by this approach)
//   OpenCASCADE:    LGPL-2.1 — same as above (occt-import-js wraps OCCT)
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const distDir = path.join(root, 'dist');

if (!fs.existsSync(distDir)) {
  console.error('copy-wasm-assets: dist/ does not exist — run heft build first.');
  process.exit(1);
}

// release/assets/ is populated by heft package-solution from dist/, but may exist
// from a previous run. Copy there too if present.
const releaseDir = path.join(root, 'release', 'assets');

const sources = [
  // web-ifc WASM (single-thread only — SharePoint Online lacks SharedArrayBuffer for MT version)
  {
    src: path.join(root, 'node_modules', 'web-ifc', 'web-ifc.wasm'),
    name: 'web-ifc.wasm'
  },
  // web-ifc license (MPL-2.0)
  {
    src: path.join(root, 'node_modules', 'web-ifc', 'LICENSE.md'),
    name: 'LICENSE.web-ifc.md'
  },
  // occt-import-js WASM (LGPL-2.1 — must remain a separate file per LGPL relinkability requirement)
  {
    src: path.join(root, 'node_modules', 'occt-import-js', 'dist', 'occt-import-js.wasm'),
    name: 'occt-import-js.wasm'
  },
  // occt-import-js license (LGPL-2.1)
  {
    src: path.join(root, 'node_modules', 'occt-import-js', 'dist', 'license.occt-import-js.txt'),
    name: 'LICENSE.occt-import-js.txt'
  },
  // OpenCASCADE Technology license (LGPL-2.1 — wrapped by occt-import-js)
  {
    src: path.join(root, 'node_modules', 'occt-import-js', 'dist', 'license.occt.txt'),
    name: 'LICENSE.occt.txt'
  }
];

let ok = true;
for (const { src, name } of sources) {
  if (!fs.existsSync(src)) {
    console.error('copy-wasm-assets: source not found:', src);
    ok = false;
    continue;
  }
  const dst = path.join(distDir, name);
  fs.copyFileSync(src, dst);
  const sizeKb = Math.round(fs.statSync(dst).size / 1024);
  console.log('Copied', name, '(' + sizeKb + ' KB) → dist/');

  if (fs.existsSync(releaseDir)) {
    fs.copyFileSync(src, path.join(releaseDir, name));
    console.log('Copied', name, '→ release/assets/');
  }
}

if (!ok) {
  process.exit(1);
}
console.log('WASM assets ready.');
