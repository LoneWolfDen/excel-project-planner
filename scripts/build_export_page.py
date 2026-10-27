"""Build docs/export.html: the browser version of scripts/export_mspdi.py and scripts/import_mspdi.py.

The page runs the planner's own Python, unchanged, in Pyodide (Python in the browser): planner/mspdi.py turns a
workbook into MS Project XML, and planner/mspdi_import.py turns MS Project XML into a new workbook through the same
builder as scripts/build_planner.py. The browser and the command line therefore share one field mapping.

Everything the page needs besides Pyodide itself is bundled into it as one zip: the planner package and openpyxl /
et_xmlfile (MIT licensed, pinned in requirements-dev.txt, licence files included). So the page works opened straight
from disk, from GitHub Pages or any web server, and never fetches from PyPI. The zip is built deterministically
(sorted files, fixed timestamps), and CI fails if docs/export.html is out of date:

    python scripts/build_export_page.py
"""
import base64
import importlib.metadata
import io
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/"
OUT = ROOT / "docs" / "export.html"
BUNDLED = ("openpyxl", "et_xmlfile")

PAGE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Planner and MS Project</title>
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
  main { max-width: 680px; margin: 0 auto; padding: 32px 16px 48px; }
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
  #notes { margin: 8px 0 0; padding-left: 1.2em; color: var(--muted); font-size: .9rem; }
  #download { display: none; margin-top: 12px; background: var(--accent); color: var(--accent-text); border: 0;
              border-radius: 8px; padding: 12px 18px; font: inherit; font-weight: 600; text-decoration: none; }
  table { border-collapse: collapse; width: 100%; font-size: .95rem; }
  td { border-top: 1px solid var(--line); padding: 8px 6px; vertical-align: top; }
  td:first-child { font-weight: 600; white-space: nowrap; }
  h2 { font-size: 1.05rem; margin: 28px 0 8px; }
  .small { color: var(--muted); font-size: .9rem; }
</style>
</head>
<body>
<main>
  <h1>Planner and MS Project</h1>
  <p class="lead">Converts the planner workbook to a Microsoft Project file, or a Microsoft Project file to a new
    planner workbook: tasks, links, calendar, holidays, constraints, baseline, resources and assignments.</p>
  <div class="card">
    <label class="drop" id="drop" for="file">
      <strong>Choose a planner workbook (.xlsx) or an MS Project XML file (.xml)</strong>
      <span>or drop it here</span>
      <input type="file" id="file" accept=".xlsx,.xml,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/xml,application/xml">
    </label>
    <p id="status" role="status" aria-live="polite"></p>
    <ul id="notes"></ul>
    <a id="download" download>Download</a>
  </div>
  <h2>How</h2>
  <table>
    <tr><td>Workbook to MS Project</td><td>Save the workbook in Excel first (so it holds up-to-date values), choose it
      here, download the .xml. In MS Project: File &gt; Open, choose the .xml, "Open as a new project".</td></tr>
    <tr><td>MS Project to workbook</td><td>In MS Project: File &gt; Save As &gt; XML. Choose that file here and download
      the workbook. Excel calculates it when you open it. Anything the planner can't hold (a fifth predecessor, a
      lag in % ...) is listed here and noted on the task.</td></tr>
  </table>
  <p class="small">The conversion runs on this device: no file is uploaded anywhere. The first conversion downloads
    the converter (Pyodide, about 10 MB), which the browser then keeps.</p>
</main>
<script src="__PYODIDE__pyodide.js"></script>
<script>
const BUNDLE = "__BUNDLE__";
const status = document.getElementById("status");
const notes = document.getElementById("notes");
const link = document.getElementById("download");
let ready = null;

function say(text, kind) { status.textContent = text; status.className = kind || ""; }

function converter() {
  if (!ready) {
    ready = loadPyodide({ indexURL: "__PYODIDE__" }).then(py => {
      const zip = Uint8Array.from(atob(BUNDLE), c => c.charCodeAt(0));
      py.unpackArchive(zip, "zip", { extractDir: "/home/pyodide/lib" });
      py.runPython("import sys; sys.path.insert(0, '/home/pyodide/lib')");
      return py;
    });
  }
  return ready;
}

const RUN = `
import base64
from planner import mspdi, mspdi_import
try:
    if kind == "xlsx":
        xml = mspdi.convert(bytes(data.to_py()))
        out = ["ok", base64.b64encode(xml.encode("utf-8")).decode(), [], xml.count("<Task>"), xml.count("<Resource>")]
    else:
        xlsx, warnings = mspdi_import.convert(bytes(data.to_py()))
        project, _ = mspdi_import.read_project(bytes(data.to_py()))
        out = ["ok", base64.b64encode(xlsx).decode(), warnings, len(project.tasks), len(project.resources)]
except (mspdi.ExportError, mspdi_import.MSPImportError) as e:
    out = ["error", str(e), [], 0, 0]
out`;

async function convert(file) {
  link.style.display = "none";
  notes.textContent = "";
  const kind = /\\.xml$/i.test(file.name) ? "xml" : "xlsx";
  say("Loading the converter...");
  try {
    const py = await converter();
    say("Converting " + file.name + "...");
    py.globals.set("data", new Uint8Array(await file.arrayBuffer()));
    py.globals.set("kind", kind);
    const [state, payload, warnings, tasks, people] = py.runPython(RUN).toJs();
    if (state !== "ok") { say(payload, "bad"); return; }
    const bytes = Uint8Array.from(atob(payload), c => c.charCodeAt(0));
    const type = kind === "xlsx" ? "application/xml"
                                 : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    link.href = URL.createObjectURL(new Blob([bytes], { type }));
    link.download = file.name.replace(/\\.(xlsx|xml)$/i, "") + (kind === "xlsx" ? ".xml" : ".xlsx");
    link.textContent = kind === "xlsx" ? "Download the MS Project file (.xml)" : "Download the workbook (.xlsx)";
    link.style.display = "inline-block";
    for (const w of warnings) { const li = document.createElement("li"); li.textContent = w; notes.appendChild(li); }
    say("Done: " + tasks + " tasks, " + people + " resources" + (warnings.length ? ". Notes:" : "."), "ok");
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


def bundle():
    """The planner package and the bundled libraries as one deterministic zip (sorted, fixed timestamps)."""
    files = {f"planner/{p.name}": p.read_bytes() for p in sorted((ROOT / "planner").glob("*.py"))}
    for package in BUNDLED:
        dist = importlib.metadata.distribution(package)
        for f in dist.files:
            if str(f).endswith(".py") or "LICEN" in str(f).upper():
                files[str(f)] = dist.locate_file(f).read_bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for path in sorted(files):
            info = zipfile.ZipInfo(path, (1980, 1, 1, 0, 0, 0))
            info.compress_type, info.external_attr = zipfile.ZIP_DEFLATED, 0o644 << 16
            z.writestr(info, files[path])
    return out.getvalue()


def page():
    return PAGE.replace("__PYODIDE__", PYODIDE).replace("__BUNDLE__", base64.b64encode(bundle()).decode())


if __name__ == "__main__":
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text(page(), encoding="utf-8")
    print("saved", OUT, f"({OUT.stat().st_size // 1024} KB)")
