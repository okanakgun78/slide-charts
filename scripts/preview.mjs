// Serves dist/ locally and screenshots the task pane (outside Office it shows the chart preview).
import { createServer } from "node:http";
import { readFileSync, existsSync } from "node:fs";
import { extname, join } from "node:path";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const { chromium } = require("playwright");

const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".xml": "application/xml" };
const server = createServer((req, res) => {
  const p = join("dist", decodeURIComponent(req.url.split("?")[0]));
  if (!existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { "content-type": types[extname(p)] || "application/octet-stream" });
  res.end(readFileSync(p));
}).listen(0);
const port = server.address().port;
const out = process.argv[2] || "preview.png";
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 380, height: 900 }, deviceScaleFactor: 1 });
page.on("pageerror", (e) => console.error("pageerror:", e.message));
await page.route(/appsforoffice\.microsoft\.com/, (r) => r.abort());
await page.goto(`http://localhost:${port}/taskpane.html`);
await page.waitForSelector("#preview svg");
await page.screenshot({ path: out, fullPage: true });
await browser.close();
server.close();
console.log("wrote", out);
