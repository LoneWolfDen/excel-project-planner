"""Gather an edition's workbooks for a GitHub Release (.github/workflows/release.yml).

    python scripts/release_assets.py v1.0 dist release-notes.md [--root checkout-of-the-tag]

The tag names the edition (v1.0, v1.0.1 -> Release 1.0; v1.2 -> template 1.2 ...). Copies its committed workbooks
into dist/, named with the edition so files from different editions never clash (Project_Planner_1.0.xlsx), writes
SHA256SUMS.txt beside them and the edition's CHANGELOG section(s) to the notes file, and prints the key=value lines
the workflow reads: the test classes that rebuild the edition and compare it with the committed files, the
release_check edition and the release title.
"""
import hashlib
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
# tag prefix -> (title, folders and the label in their file names, release_check edition, tests that rebuild it)
EDITIONS = {
    "v1.0": ("Release 1.0", [("release-1.0", "1.0")], "1.0", ["test_release.CommittedReleaseTest"]),
    "v1.1": ("Template 1.1 and the Microsoft 365 edition", [("release-1.1", "1.1"), ("release-1.1-m365", "1.1-M365")],
             "1.1", ["test_release_11.CommittedTest", "test_m365.CommittedTest"]),
    "v1.2": ("Template 1.2", [("release-1.2", "1.2"), ("release-1.2/samples", "1.2")], "1.2",
             ["test_release_12.CommittedTest", "test_release_12.SampleBuildTest"]),
}


def edition(tag):
    """The EDITIONS entry for a tag: v1.0 and v1.0.3 both mean Release 1.0."""
    m = re.match(r"^(v\d+\.\d+)(\.\d+)?$", tag)
    if not m or m.group(1) not in EDITIONS:
        raise SystemExit(f"{tag}: not a release tag (expected one of {', '.join(EDITIONS)}, optionally .N)")
    return EDITIONS[m.group(1)]


def notes(root, folders):
    """The CHANGELOG sections whose heading names one of the folders (`release-1.0/`)."""
    text = (root / "CHANGELOG.md").read_text()
    sections = re.split(r"(?m)^(?=## )", text)
    keep = [s.strip() for s in sections if s.startswith("## ")
            and any(f"`{folder.split('/')[0]}/`" in s.splitlines()[0] for folder, _ in folders)]
    return "\n\n".join(keep) + "\n\nEach workbook works in Excel alone (no macros). Check the downloads against " \
                                "SHA256SUMS.txt.\n"


def main(argv=None):
    args = list(argv or sys.argv[1:])
    root = ROOT
    if "--root" in args:                             # the tag's checkout (the workflow runs this script from main)
        i = args.index("--root")
        root = Path(args[i + 1]).resolve()
        del args[i:i + 2]
    tag, out, notes_file = args[:3]
    title, folders, check, tests = edition(tag)
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    sums = []
    for folder, label in folders:
        for f in sorted((root / folder).glob("*.xlsx")):
            target = out / f"{f.stem}_{label}{f.suffix}"
            shutil.copyfile(f, target)
            sums.append(f"{hashlib.sha256(target.read_bytes()).hexdigest()}  {target.name}")
    if not sums:
        raise SystemExit(f"{tag}: no workbooks in {', '.join(f for f, _ in folders)}")
    (out / "SHA256SUMS.txt").write_text("\n".join(sums) + "\n")
    Path(notes_file).write_text(notes(root, folders))
    print(f"title={title} {tag}")
    print(f"check={check}")
    print(f"tests={' '.join(tests)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
