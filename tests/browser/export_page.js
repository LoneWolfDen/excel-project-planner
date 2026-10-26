// Drive docs/export.html in headless Chromium: choose a workbook, save the download.
//   node tests/browser/export_page.js <page.html> <workbook.xlsx> <pyodide dir> <out.xml>
// Pyodide is served from <pyodide dir> instead of the CDN, so the test needs no network.
const path = require("path");
const fs = require("fs");
const { chromium } = require("playwright");

const [page, workbook, pyodideDir, out] = process.argv.slice(2);
const CDN = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
const TYPES = { ".js": "text/javascript", ".mjs": "text/javascript", ".wasm": "application/wasm",
                ".json": "application/json", ".zip": "application/zip" };

(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext({ acceptDownloads: true });
  await context.route(CDN + "**", route => {
    const file = path.join(pyodideDir, route.request().url().slice(CDN.length).split("?")[0]);
    if (!fs.existsSync(file)) return route.fulfill({ status: 404, body: "missing " + file });
    route.fulfill({ status: 200, body: fs.readFileSync(file),
                    headers: { "content-type": TYPES[path.extname(file)] || "application/octet-stream",
                               "access-control-allow-origin": "*" } });
  });
  const tab = await context.newPage();
  await tab.goto("file://" + path.resolve(page));
  await tab.setInputFiles("#file", workbook);
  await tab.waitForSelector("#status.ok, #status.bad", { timeout: 120000 });
  const status = await tab.textContent("#status");
  if (await tab.$("#status.bad")) {
    console.log("STATUS " + status);
  } else {
    const [download] = await Promise.all([tab.waitForEvent("download"), tab.click("#download")]);
    await download.saveAs(out);
    console.log("STATUS " + status);
  }
  await browser.close();
})().catch(e => { console.error(e); process.exit(1); });
