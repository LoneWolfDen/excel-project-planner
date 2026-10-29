"""Reproducible workbook files: the same build gives the same bytes.

LibreOffice gives every chart axis a random id each time it saves, and both it and openpyxl stamp the file with the
time it was created and saved; the zip entries carry that time too. So a rebuild of an unchanged workbook always
looked changed in git, and a real change was hard to see among the noise. normalise() rewrites those parts after
each save:

  * chart axis ids (c:axId and the c:crossAx that points at it) are renumbered 1001, 1002, ... in the order they
    appear in each chart: the same chart gets the same ids every time;
  * docProps/core.xml's created and modified times become FIXED_TIME (or SOURCE_DATE_EPOCH, the usual
    reproducible-builds variable, when it is set);
  * every zip entry gets the same timestamp (ZIP_TIME).

Nothing a user sees changes: Excel doesn't show axis ids, and the version is in TemplateVersion, not the file date.
tests/test_reproducible.py builds twice and compares the bytes.
"""
import datetime as dt
import os
import re
import zipfile
from pathlib import Path

FIXED_TIME = "2026-01-01T00:00:00Z"
ZIP_TIME = (1980, 1, 1, 0, 0, 0)                      # the earliest a zip entry can hold
AXIS = re.compile(r'<c:(axId|crossAx) val="(\d+)"/>')
CORE_TIME = re.compile(r'(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:(?:created|modified)>)')


def stamp():
    """The time written into docProps/core.xml."""
    epoch = os.environ.get("SOURCE_DATE_EPOCH")
    if epoch and epoch.isdigit():
        return dt.datetime.fromtimestamp(int(epoch), dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    return FIXED_TIME


def _axes(xml):
    """Renumber one chart's axis ids in order of first appearance (a crossAx names an axis of the same chart)."""
    ids = {}
    for _, old in AXIS.findall(xml):
        ids.setdefault(old, str(1001 + len(ids)))
    return AXIS.sub(lambda m: f'<c:{m.group(1)} val="{ids[m.group(2)]}"/>', xml)


def normalise(path):
    """Rewrite the saved workbook at path in place so the same content gives the same bytes."""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    when = stamp()
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/charts/chart") and item.filename.endswith(".xml"):
                data = _axes(data.decode("utf-8")).encode("utf-8")
            elif item.filename == "docProps/core.xml":
                data = CORE_TIME.sub(lambda m: m.group(1) + when + m.group(2), data.decode("utf-8")).encode("utf-8")
            info = zipfile.ZipInfo(item.filename, date_time=ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = item.external_attr
            zout.writestr(info, data)
    tmp.replace(path)
