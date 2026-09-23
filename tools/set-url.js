// Point the manifest at wherever you host the "web" folder (e.g. GitHub Pages).
// Usage: npm run set-url -- https://yourname.github.io/worksheet-builder/
const fs = require('fs');
const path = require('path');
let url = process.argv[2];
if (!url || !/^https:\/\//.test(url)) {
  console.error('Please give the https address where the web folder is hosted, e.g.\n  npm run set-url -- https://yourname.github.io/worksheet-builder/');
  process.exit(1);
}
if (!url.endsWith('/')) url += '/';
const src = fs.readFileSync(path.join(__dirname, '..', 'manifest.xml'), 'utf8');
const out = src.split('https://localhost:3000/').join(url);
fs.writeFileSync(path.join(__dirname, '..', 'manifest-hosted.xml'), out);
console.log('Created manifest-hosted.xml pointing at ' + url);
