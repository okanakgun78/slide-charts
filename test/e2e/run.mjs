// End-to-end test of the side panel against the mock PowerPoint (test/e2e/mock-office.js).
// Usage: npm run build && node test/e2e/run.mjs [outDir]
import { createServer } from "node:http";
import { readFileSync, existsSync, writeFileSync, mkdirSync } from "node:fs";
import { extname, join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import assert from "node:assert/strict";

const require = createRequire(import.meta.url);
const { chromium } = require("playwright");
const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, "..", "..");
const out = process.argv[2] || join(root, "e2e-out");
mkdirSync(out, { recursive: true });

const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png", ".xml": "application/xml" };
const server = createServer((req, res) => {
  const p = join(root, "dist", decodeURIComponent(req.url.split("?")[0]));
  if (!existsSync(p)) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { "content-type": types[extname(p)] || "application/octet-stream" });
  res.end(readFileSync(p));
}).listen(0);
const port = server.address().port;

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 380, height: 1000 }, deviceScaleFactor: 1 });
const pageErrors = [];
page.on("pageerror", (e) => pageErrors.push(e.message));
await page.addInitScript({ path: join(here, "mock-office.js") });
await page.route(/appsforoffice\.microsoft\.com/, (r) => r.abort());
await page.goto(`http://localhost:${port}/taskpane.html`);
await page.waitForFunction(() => document.getElementById("version").textContent.length > 0);

const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__deck, (k, v) => (k === "parent" ? undefined : v))));
const charts = async () => (await deck()).slides[0].shapes.filter((s) => s.tags && s.tags.SLIDECHARTS);
const spec = async () => JSON.parse((await charts()).at(-1).tags.SLIDECHARTS).spec;
const settle = (ms = 1200) => page.waitForTimeout(ms);
const errors = () => page.evaluate(() => window.__errors);
let step = 0;
const ok = (name) => console.log(`  ✓ ${++step}. ${name}`);
const allLines = (shapes) => shapes.flatMap((s) => (s.children ? allLines(s.children) : s.type === "Line" ? [s] : []));

try {
  // 1. Insert a stacked column chart.
  await page.click('#insert button[data-type="column"][data-mode="stacked"]');
  await page.waitForFunction(() => !document.getElementById("editor").hidden);
  await settle(300);
  let c = await charts();
  assert.equal(c.length, 1, "one chart group");
  assert.equal(c[0].type, "Group");
  assert.deepEqual((await spec()).type, "column");
  for (const l of allLines((await deck()).slides[0].shapes)) assert.ok(Math.min(l.width, l.height) < 1, `line ${l.name} has no 72pt side`);
  assert.equal(await page.locator("#sheet tbody tr").count(), 3);
  ok("insert stacked column: group + tag, flat lines, sheet with 3 series");

  // 2. Type a value -> chart redraws in place.
  const firstId = c[0].id;
  await page.fill('#sheet input[data-r="1"][data-c="1"]', "100");
  await settle();
  c = await charts();
  assert.equal(c.length, 1, "still one chart");
  assert.notEqual(c[0].id, firstId, "redrawn");
  assert.equal((await spec()).series[0].values[0], 100);
  ok("typing a value redraws the chart");

  // 3. Paste a whole table from Excel into the first data cell.
  await page.focus('#sheet input[data-r="1"][data-c="1"]');
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/plain", "\tQ1\tQ2\tQ3\r\nAlpha\t1,200\t1,350\t(50)\r\nBeta\t300\t320\t410\r\n");
    document.activeElement.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  await settle();
  let s = await spec();
  assert.deepEqual(s.categories, ["Q1", "Q2", "Q3"]);
  assert.deepEqual(s.series.map((x) => x.name), ["Alpha", "Beta"]);
  assert.deepEqual(s.series[0].values, [1200, 1350, -50]);
  ok("paste from Excel replaces the table (thousands separators, bracket negatives)");

  // 4. Change chart type.
  await page.selectOption("#ed-type", "bar:clustered");
  await settle();
  s = await spec();
  assert.equal(s.type, "bar");
  assert.equal(s.mode, "clustered");
  ok("switch to clustered bar");

  // 5. User resizes the chart in PowerPoint, then edits: chart refits to the new size.
  const before = (await charts())[0];
  await page.evaluate((id) => {
    const g = window.__deck.slides[0].shapes.find((x) => x.id === id);
    const f = 0.5;
    g.children.forEach((k) => { k.left = g.left + (k.left - g.left) * f; k.width *= f; });
    g.width *= f;
  }, before.id);
  await page.selectOption("#ed-dec", "1");
  await settle();
  const after = (await charts())[0];
  assert.ok(Math.abs(after.width - before.width * 0.5) < before.width * 0.08, `refit width ${after.width} vs ${before.width * 0.5}`);
  ok("after a manual resize, the redraw fills the new size");

  // 6/7. Deselect hides the editor; selecting the chart again opens it.
  await page.evaluate(() => window.__select([]));
  await page.waitForFunction(() => document.getElementById("editor").hidden);
  await page.evaluate((id) => window.__select([id]), after.id);
  await page.waitForFunction(() => !document.getElementById("editor").hidden);
  assert.equal(await page.inputValue('#sheet thead input[data-c="1"]'), "Q1");
  ok("deselect closes the editor; reselect loads that chart's data");

  // 8. Insert into a selected placeholder: the chart takes its bounds.
  await page.evaluate(() => {
    const ph = { id: "ph1", type: "GeometricShape", left: 500, top: 300, width: 400, height: 200, tags: {} };
    window.__deck.slides[0].shapes.push(ph);
    window.__select(["ph1"]);
  });
  await settle(400);
  await page.click('#insert button[data-type="line"]');
  await settle();
  c = await charts();
  const ln = c.at(-1);
  assert.ok(ln.left >= 499 && ln.top >= 299 && ln.left + ln.width <= 901 && ln.top + ln.height <= 501, `line chart inside placeholder: ${JSON.stringify([ln.left, ln.top, ln.width, ln.height])}`);
  ok("insert with a placeholder selected uses its position and size");

  // 9. Number format setting applies to labels.
  await page.click("#settings summary");
  await page.selectOption("#set-num", "eu");
  await settle();
  const texts = (await charts()).at(-1).children.filter((k) => k.type === "TextBox").map((k) => k.text);
  assert.ok(texts.some((t) => /^\d+,\d$/.test(t)), `decimal comma in labels: ${texts.slice(0, 8)}`);
  await page.selectOption("#set-num", "en");
  await settle();
  ok("number format setting (1.234,5) applies to the selected chart");

  // 10. Combination chart: line toggle per series.
  await page.evaluate(() => window.__select([]));
  await settle(300);
  await page.click('#insert button[data-type="combo"]');
  await settle();
  assert.equal(await page.locator("#sheet input.asline").count(), 3);
  await page.locator("#sheet input.asline").nth(1).check();
  await settle();
  s = await spec();
  assert.equal(s.series.filter((x) => x.asLine).length, 2);
  ok("combination: series can be switched to lines");

  // 11. Add and delete series / categories.
  await page.click("#add-series");
  await page.fill('#sheet input[data-r="4"][data-c="0"]', "Delta");
  await page.fill('#sheet input[data-r="4"][data-c="1"]', "12");
  await settle();
  assert.equal((await spec()).series.length, 4);
  await page.click("#del-series");
  await settle();
  assert.equal((await spec()).series.length, 3);
  await page.click("#add-cat");
  await settle();
  assert.equal((await spec()).categories.length, 6);
  await page.click("#del-cat");
  await settle();
  assert.equal((await spec()).categories.length, 5);
  ok("add/delete series and categories");

  // 12. Series colour override.
  await page.locator("#sheet button.sw").first().click();
  await page.locator(".popover button.pick").last().click();
  await settle();
  assert.equal((await spec()).series[0].color, "#C00000");
  ok("series colour can be overridden (highlight red)");

  // 13. Gantt: insert sample.
  await page.evaluate(() => window.__select([]));
  await settle(300);
  await page.click('#insert button[data-type="gantt"]');
  await settle();
  s = await spec();
  assert.equal(s.type, "gantt");
  assert.equal(s.rows.length, 8);
  assert.equal(await page.isVisible("#ed-gantt"), true);
  assert.equal(await page.isVisible("#ed-cat"), false);
  let g = (await charts()).at(-1);
  assert.ok(g.children.some((k) => k.geom === "Diamond"), "milestone diamonds");
  assert.ok(g.children.some((k) => k.type === "Line" && k.line?.dash === "Dash"), "today line");
  ok("insert Gantt: sample plan with groups, milestones and a today line");

  // 14. Type a period as an end date.
  await page.fill('#gsheet input[data-r="6"][data-c="end"]', "Q2 2027");
  await settle();
  s = await spec();
  assert.equal(s.rows[6].end, "Q2 2027");
  g = (await charts()).at(-1);
  assert.ok(g.children.some((k) => k.type === "TextBox" && k.text === "Jun"), `month scale after extending the plan: ${g.children.filter((k) => k.type === "TextBox").map((k) => k.text).slice(0, 12)}`);
  ok("typing 'Q2 2027' as an end date extends the plan to 30 Jun 2027 (weeks switch to months)");

  // 15. Paste a schedule from Excel with a title row.
  await page.focus('#gsheet input[data-r="0"][data-c="label"]');
  await page.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData("text/plain", "Task\tStart\tEnd\tOwner\r\nPhase A\t\t\t\r\nDesign\t01/02/2027\t28/02/2027\tOA\r\nBuild\t01/03/2027\t30/04/2027\t\r\nLaunch\t15/05/2027\t\t\r\n");
    document.activeElement.dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  await settle();
  s = await spec();
  assert.deepEqual(s.rows.map((r) => [r.label, r.kind]), [["Phase A", "header"], ["Design", "task"], ["Build", "task"], ["Launch", "milestone"]]);
  assert.equal(s.rows[1].note, "OA");
  ok("paste schedule from Excel: title row skipped, groups/tasks/milestones inferred, note kept");

  // 16. Unreadable date: flagged in the sheet and reported, chart still drawn.
  await page.fill('#gsheet input[data-r="2"][data-c="end"]', "sometime");
  await settle();
  assert.equal(await page.locator('#gsheet input[data-r="2"][data-c="end"].bad').count(), 1);
  const st = await page.textContent("#ed-status");
  assert.match(st, /not recognised/, `status was: "${st}"`);
  await page.fill('#gsheet input[data-r="2"][data-c="end"]', "30/04/2027");
  await settle();
  ok("unreadable dates are flagged in red and reported");

  // 17. Day/month order setting.
  await page.selectOption("#set-dates", "MDY");
  await page.fill('#gsheet input[data-r="3"][data-c="start"]', "06/05/2027");
  await settle();
  g = (await charts()).at(-1);
  assert.ok(g.children.some((k) => k.type === "TextBox" && k.text === "Jun 5"), `MDY milestone label: ${g.children.filter((k) => k.type === "TextBox").map((k) => k.text)}`);
  await page.selectOption("#set-dates", "DMY");
  await settle();
  ok("date order setting: 06/05/2027 reads as June 5 in month-first mode");

  // 18. Move and delete rows.
  await page.focus('#gsheet input[data-r="3"][data-c="label"]');
  await page.click("#g-up");
  await settle();
  assert.equal((await spec()).rows[2].label, "Launch");
  await page.click("#g-del");
  await settle();
  assert.equal((await spec()).rows.length, 3);
  ok("move row up and delete row");

  assert.deepEqual(await errors(), [], "no Office.js errors");
  assert.deepEqual(pageErrors, [], "no page errors");
  ok("no Office.js or page errors");
} catch (e) {
  console.error("FAILED at step", step + 1, "-", e.message);
  console.error("Office errors:", await errors().catch(() => "?"), "page errors:", pageErrors);
  process.exitCode = 1;
} finally {
  await page.screenshot({ path: join(out, "panel.png"), fullPage: true });
  // Draw the deck as PowerPoint would roughly show it.
  const d = await deck();
  writeFileSync(join(out, "deck.json"), JSON.stringify(d, null, 1));
  const svg = renderDeck(d);
  const p2 = await browser.newPage({ viewport: { width: 980, height: 560 } });
  await p2.setContent(`<body style="margin:0">${svg}</body>`);
  await p2.screenshot({ path: join(out, "slide.png") });
  await browser.close();
  server.close();
}

function renderDeck(d) {
  const els = [];
  const draw = (s) => {
    if (s.children) return s.children.forEach(draw);
    const cx = s.left + s.width / 2, cy = s.top + s.height / 2;
    const rot = s.rotation ? ` transform="rotate(${s.rotation} ${cx} ${cy})"` : "";
    if (s.type === "Line") els.push(`<line x1="${s.left}" y1="${s.top}" x2="${s.left + s.width}" y2="${s.top + s.height}" stroke="${s.line?.color || "#000"}" stroke-width="${s.line?.weight || 1}"${s.line?.dash ? ' stroke-dasharray="4 3"' : ""}${rot}/>`);
    else if (s.type === "TextBox") {
      const a = s.align === "Left" ? "start" : s.align === "Right" ? "end" : "middle";
      const x = s.align === "Left" ? s.left : s.align === "Right" ? s.left + s.width : cx;
      const fs = s.font?.size || 12;
      const y = s.tf?.valign === "Top" ? s.top + fs : s.tf?.valign === "Bottom" ? s.top + s.height - fs * 0.25 : cy + fs * 0.35;
      els.push(`<text x="${x}" y="${y}" font-size="${fs}" fill="${s.font?.color || "#000"}" text-anchor="${a}" font-family="Helvetica" ${s.font?.bold ? 'font-weight="bold"' : ""}>${String(s.text).replace(/</g, "&lt;")}</text>`);
    } else if (s.geom === "Diamond") els.push(`<polygon points="${cx},${s.top} ${s.left + s.width},${cy} ${cx},${s.top + s.height} ${s.left},${cy}" fill="${s.fill || "none"}"/>`);
    else if (s.geom === "Ellipse") els.push(`<ellipse cx="${cx}" cy="${cy}" rx="${s.width / 2}" ry="${s.height / 2}" fill="${s.fill || "none"}"/>`);
    else if (s.tags !== undefined && s.id === "ph1") els.push(`<rect x="${s.left}" y="${s.top}" width="${s.width}" height="${s.height}" fill="none" stroke="#bbb" stroke-dasharray="3 3"/>`);
    else els.push(`<rect x="${s.left}" y="${s.top}" width="${s.width}" height="${s.height}" fill="${s.fill || "none"}"${rot}/>`);
  };
  d.slides[0].shapes.forEach(draw);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="960" height="540" viewBox="0 0 960 540" style="background:#fff;border:1px solid #ccc">${els.join("")}</svg>`;
}
