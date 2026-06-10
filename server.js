/*
 * server.js
 * ---------
 * Tiny static HTTPS server for local Office Add-in development.
 * Serves the project folder over https://localhost:3000 using the trusted
 * dev certificate created by `office-addin-dev-certs` (run via `npm start`).
 *
 * No build step, no framework — just enough to host taskpane.html and assets
 * so Word on the web / SharePoint can load the add-in.
 */

"use strict";

const fs = require("fs");
const path = require("path");
const https = require("https");
const devCerts = require("office-addin-dev-certs");

const PORT = 3000;
const ROOT = __dirname;

// Minimal content-type lookup for the handful of file types we serve.
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".xml": "application/xml; charset=utf-8",
};

function send(res, status, body, contentType) {
  res.writeHead(status, {
    "Content-Type": contentType || "text/plain; charset=utf-8",
    // CORS headers so Office hosts can fetch the assets.
    "Access-Control-Allow-Origin": "*",
    "Cache-Control": "no-store",
  });
  res.end(body);
}

function handleRequest(req, res) {
  // Map "/" to the task pane, strip query string, and prevent path traversal.
  let urlPath = decodeURIComponent(req.url.split("?")[0]);
  if (urlPath === "/") {
    urlPath = "/src/taskpane.html";
  }

  const filePath = path.join(ROOT, path.normalize(urlPath));
  if (!filePath.startsWith(ROOT)) {
    return send(res, 403, "Forbidden");
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      return send(res, 404, "Not found: " + urlPath);
    }
    const ext = path.extname(filePath).toLowerCase();
    send(res, 200, data, MIME[ext] || "application/octet-stream");
  });
}

(async function main() {
  // Obtain the locally-trusted HTTPS certificate (installed by `npm start`).
  const httpsOptions = await devCerts.getHttpsServerOptions();

  https.createServer(httpsOptions, handleRequest).listen(PORT, () => {
    console.log(`Perfect Align dev server running at https://localhost:${PORT}/`);
    console.log(`Task pane:  https://localhost:${PORT}/src/taskpane.html`);
    console.log("Press Ctrl+C to stop.");
  });
})().catch((err) => {
  console.error("Failed to start dev server:", err);
  process.exit(1);
});
