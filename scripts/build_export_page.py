"""Build docs/export.html: the browser version of scripts/export_mspdi.py.

The page runs planner/mspdi.py itself, unchanged, in Pyodide (Python in the browser), so the browser and the command
line share one field mapping. The module is inlined here, so the page is a single file that works when opened
straight from disk, from GitHub Pages or from any web server. CI fails if docs/export.html is out of date:

    python scripts/build_export_page.py
"""
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/"
OUT = ROOT / "docs" / "export.html"

PAGE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Planner to MS Project</title>
<style>
  :root { --bg: #f7f8fa; --card: #ffffff; --text: #1d2433; --muted: #5b6475; --line: #d9dee7; --accent: #1f4e78;
          --accent-text: #ffffff; --ok: #2e7d32; --bad: #b3261e; --drop: #eef3f9; }
  @media (prefers-color-scheme: dark) {
    :root { --bg: #14171c; --card: #1d2128; --text: #e6e9ef; --muted: #a3abba; --line: #343a45; --accent: #6ea8dc;
            --accent-text: #0d1117; --ok: #81c784; --bad: #f28b82; --drop: #1f2a37; }
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--text); font: 16px/1.5 system-ui, -apple-system, "Segoe UI",
         Roboto, sans-serif; }
  main { max-width: 640px; margin: 0 auto; padding: 32px 16px 48px; }
  h1 { font-size: 1.5rem; margin: 0 0 4px; }
  .lead { color: var(--muted); margin: 0 0 24px; }
  .card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 20px; }
  .drop { display: block; border: 2px dashed var(--line); border-radius: 10px; padding: 28px 16px; text-align: center;
          background: var(--drop); cursor: pointer; }
  .drop.over { border-color: var(--accent); }
  .drop strong { display: block; font-size: 1.05rem; }
  .drop span { color: var(--muted); font-size: .9rem; }
  input[type=file] { position: absolute; width: 1px; height: 1px; opacity: 0; }
  #status { margin: 16px 0 0; min-height: 1.5em; }
  #status.ok { color: var(--ok); } #status.bad { color: var(--bad); }
  #download { display: none; margin-top: 12px; background: var(--accent); color: var(--accent-text); border: 0;
              border-radius: 8px; padding: 12px 18px; font: inherit; font-weight: 600; text-decoration: none; }
  ol { padding-left: 1.2em; } li { margin: 4px 0; }
  h2 { font-size: 1.05rem; margin: 28px 0 8px; }
  .small { color: var(--muted); font-size: .9rem; }
</style>
</head>
<body>
<main>
  <h1>Planner to MS Project</h1>
  <p class="lead">Turns the planner workbook into a Microsoft Project XML file, with tasks, links, calendar, holidays,
    constraints, baseline, resources and assignments.</p>
  <div class="card">
    <label class="drop" id="drop" for="file">
      <strong>Choose the planner workbook (.xlsx)</strong>
      <span>or drop it here</span>
      <input type="file" id="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
    </label>
    <p id="status" role="status" aria-live="polite"></p>
    <a id="download" download>Download the .xml</a>
  </div>
  <h2>Open it in MS Project</h2>
  <ol>
    <li>Save the workbook in Excel first, so it holds up-to-date values.</li>
    <li>In MS Project: File &gt; Open, choose the .xml, then "Open as a new project".</li>
  </ol>
  <p class="small">The conversion runs on this device: the workbook is not uploaded anywhere. The first conversion
    downloads the converter (Pyodide, about 10 MB), which the browser then keeps.</p>
</main>
<script src="__PYODIDE__pyodide.js"></script>
<script>
const SOURCE = __SOURCE__;
const status = document.getElementById("status");
const link = document.getElementById("download");
let ready = null;

function say(text, kind) { status.textContent = text; status.className = kind || ""; }

function converter() {
  if (!ready) {
    ready = loadPyodide({ indexURL: "__PYODIDE__" }).then(py => {
      py.FS.writeFile("/home/pyodide/mspdi.py", SOURCE);
      py.runPython("import sys; sys.path.insert(0, '/home/pyodide'); import mspdi");
      return py;
    });
  }
  return ready;
}

async function convert(file) {
  link.style.display = "none";
  say("Loading the converter...");
  try {
    const py = await converter();
    say("Converting " + file.name + "...");
    py.globals.set("data", new Uint8Array(await file.arrayBuffer()));
    const result = py.runPython(`
import mspdi
try:
    out = ("ok", mspdi.convert(bytes(data.to_py())))
except mspdi.ExportError as e:
    out = ("error", str(e))
out`).toJs();
    if (result[0] !== "ok") { say(result[1], "bad"); return; }
    const xml = result[1];
    const tasks = (xml.match(/<Task>/g) || []).length;
    const people = (xml.match(/<Resource>/g) || []).length;
    link.href = URL.createObjectURL(new Blob([xml], { type: "application/xml" }));
    link.download = file.name.replace(/\\.xlsx$/i, "") + ".xml";
    link.style.display = "inline-block";
    say("Done: " + tasks + " tasks, " + people + " resources.", "ok");
  } catch (e) {
    say("Something went wrong: " + e.message, "bad");
  }
}

const input = document.getElementById("file");
const drop = document.getElementById("drop");
input.addEventListener("change", () => input.files[0] && convert(input.files[0]));
drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", e => {
  e.preventDefault(); drop.classList.remove("over");
  if (e.dataTransfer.files[0]) convert(e.dataTransfer.files[0]);
});
</script>
</body>
</html>
"""


def page():
    source = (ROOT / "planner" / "mspdi.py").read_text(encoding="utf-8")
    inline = json.dumps(source).replace("</", "<\\/")      # never close the <script> early
    return PAGE.replace("__PYODIDE__", PYODIDE).replace("__SOURCE__", inline)


if __name__ == "__main__":
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(page(), encoding="utf-8")
    print("saved", OUT)
