// Local HTTPS server for Worksheet Builder (Word only loads add-ins over HTTPS).
// Run:  npm start   → serves ./web at https://localhost:3000
const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const ROOT = path.join(__dirname, 'web');
const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json', '.ico': 'image/x-icon' };

function handler(req, res) {
  let p = decodeURIComponent(new URL(req.url, 'https://localhost').pathname);
  if (p === '/') p = '/taskpane.html';
  const file = path.join(ROOT, path.normalize(p));
  if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end('Not found'); }
    res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream', 'Cache-Control': 'no-cache' });
    res.end(data);
  });
}

(async () => {
  if (process.env.PLAIN_HTTP) {
    http.createServer(handler).listen(PORT, () => console.log('Serving (plain HTTP, testing only) on http://localhost:' + PORT));
    return;
  }
  const devCerts = require('office-addin-dev-certs');
  const options = await devCerts.getHttpsServerOptions(); // installs a trusted localhost certificate the first time
  https.createServer(options, handler).listen(PORT, () => {
    console.log('Worksheet Builder is running at https://localhost:' + PORT + '/taskpane.html');
    console.log('Leave this window open while you use the add-in in Word. Press Ctrl+C to stop.');
  });
})().catch((e) => { console.error(e.message || e); process.exit(1); });
