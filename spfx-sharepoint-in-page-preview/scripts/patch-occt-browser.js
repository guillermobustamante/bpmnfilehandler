// Patches occt-import-js/package.json to add browser-safe stubs for Node.js built-ins.
// Without this patch, webpack 5 fails to build because occt-import-js conditionally
// requires 'path' and 'crypto' in its Node.js code path (which is never reached in browser).
const fs = require('fs');
const path = require('path');

const pkgPath = path.join(__dirname, '..', 'node_modules', 'occt-import-js', 'package.json');

if (!fs.existsSync(pkgPath)) {
  console.log('occt-import-js not found, skipping browser field patch');
  process.exit(0);
}

const pkg = JSON.parse(fs.readFileSync(pkgPath, 'utf8'));
const browser = pkg.browser || {};
let changed = false;

['path', 'crypto', 'fs'].forEach((mod) => {
  if (browser[mod] !== false) {
    browser[mod] = false;
    changed = true;
  }
});

if (changed) {
  pkg.browser = browser;
  fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, '\t') + '\n');
  console.log('Patched occt-import-js browser field: path, crypto, fs → false');
}
