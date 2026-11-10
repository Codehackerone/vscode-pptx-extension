// Copies browser builds of runtime libraries into media/vendor so the webview can load them.
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const dest = path.join(root, 'media', 'vendor');
fs.mkdirSync(dest, { recursive: true });

const files = [['node_modules/jszip/dist/jszip.min.js', 'jszip.min.js']];
for (const [from, to] of files) {
  fs.copyFileSync(path.join(root, from), path.join(dest, to));
  console.log(`vendor: ${from} -> media/vendor/${to}`);
}
