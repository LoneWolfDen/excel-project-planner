"""Run the test suite in parallel: one test class per process, several processes at once.

    python scripts/run_tests.py -j 4                # every test class, 4 at a time
    python scripts/run_tests.py --shard 2/3         # the second of three balanced shards (CI runs them side by side)
    python scripts/run_tests.py -j 4 --record       # also save each class's time in tests/durations.json

Most of the suite's time is LibreOffice recalculating workbooks, one at a time. LibreOffice can run several copies
side by side if each has its own profile, so each worker gets its own (PLANNER_LO_PROFILE) and its own range of
UNO ports (PLANNER_LO_PORT); planner.build, tests/test_planner.recalc, tests/lo_session and scripts/perf_check read
them. A class runs in a fresh process, so no class sees another's edition or module state (the same isolation the
full `unittest discover` run relies on tearDownClass for).

Shards are balanced by the times in tests/durations.json (longest first, each to the lightest shard); a class with no
recorded time counts as 60 s. The result is the same set of tests as `python -m unittest discover -s tests`.
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import unittest
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TESTS = ROOT / "tests"
DURATIONS = TESTS / "durations.json"
UNKNOWN = 60.0                    # seconds, for a class not yet in durations.json
RAN = re.compile(r"^Ran (\d+) tests?", re.M)
SKIPPED = re.compile(r"skipped=(\d+)")


def test_classes():
    """Every test class, as "module.Class" (module relative to tests/), in a stable order."""
    sys.path.insert(0, str(TESTS))
    sys.path.insert(0, str(ROOT))
    suite = unittest.defaultTestLoader.discover(str(TESTS), top_level_dir=str(TESTS))
    found = set()

    def walk(s):
        for t in s:
            if isinstance(t, unittest.TestSuite):
                walk(t)
            elif not t.id().startswith("unittest.loader"):
                found.add(t.id().rsplit(".", 1)[0])
            else:                                    # a module that failed to import: run it to see why
                found.add(t.id())
    walk(suite)
    return sorted(found)


def durations():
    try:
        return json.loads(DURATIONS.read_text())
    except (OSError, ValueError):
        return {}


def shard(classes, k, n):
    """The k-th (1-based) of n shards, balanced by recorded time: longest first, each to the lightest shard."""
    times = durations()
    bins = [[0.0, []] for _ in range(n)]
    for c in sorted(classes, key=lambda c: (-times.get(c, UNKNOWN), c)):
        lightest = min(bins, key=lambda b: b[0])
        lightest[0] += times.get(c, UNKNOWN)
        lightest[1].append(c)
    return sorted(bins[k - 1][1])


def run_class(name, worker, tmp):
    """Run one test class in its own process with the worker's LibreOffice profile and ports."""
    env = dict(os.environ, PLANNER_LO_PROFILE=str(Path(tmp) / f"lo-{worker}"),
               PLANNER_LO_PORT=str(2101 + 200 * worker), PYTHONPATH=os.pathsep.join([str(TESTS), str(ROOT)]))
    start = time.time()
    p = subprocess.run([sys.executable, "-m", "unittest", "-v", name], cwd=TESTS, env=env,
                       capture_output=True, text=True)
    return name, p.returncode, p.stderr + p.stdout, time.time() - start


def main(argv=None):
    a = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    a.add_argument("-j", "--jobs", type=int, default=os.cpu_count() or 2, help="classes run at once")
    a.add_argument("--shard", help="k/n: run only the k-th of n balanced shards")
    a.add_argument("--record", action="store_true", help="save each class's time in tests/durations.json")
    a.add_argument("names", nargs="*", help="only these classes (module.Class), e.g. test_release.EditTest")
    args = a.parse_args(argv)
    classes = args.names or test_classes()
    if args.shard:
        k, n = (int(x) for x in args.shard.split("/"))
        classes = shard(classes, k, n)
    print(f"{len(classes)} test classes, {args.jobs} at a time", flush=True)
    results = []
    started = time.time()
    with tempfile.TemporaryDirectory() as tmp, ThreadPoolExecutor(max_workers=args.jobs) as pool:
        free = list(range(args.jobs))
        futures = []

        def submit(name):
            worker = free.pop()
            f = pool.submit(run_class, name, worker, tmp)
            f.add_done_callback(lambda _f, w=worker: free.append(w))
            return f
        # keep each worker's profile private: never more classes in flight than workers
        queue = list(sorted(classes, key=lambda c: -durations().get(c, UNKNOWN)))
        while queue or futures:
            while queue and free:
                futures.append(submit(queue.pop(0)))
            done = next(f for f in futures if f.done()) if any(f.done() for f in futures) else None
            if done is None:
                time.sleep(0.2)
                continue
            futures.remove(done)
            name, code, out, secs = done.result()
            ran = int(RAN.search(out).group(1)) if RAN.search(out) else 0
            print(f"{'ok  ' if code == 0 else 'FAIL'}  {name}  ({ran} tests, {secs:.0f} s)", flush=True)
            results.append((name, code, out, secs, ran))
    failed = [r for r in results if r[1] != 0]
    for name, _, out, _, _ in failed:
        print(f"\n{'=' * 70}\n{name}\n{'=' * 70}\n{out}")
    total = sum(r[4] for r in results)
    skipped = sum(int(m) for r in results for m in SKIPPED.findall(r[2]))
    print(f"\nRan {total} tests in {len(results)} classes in {time.time() - started:.0f} s "
          f"({'OK' if not failed else f'FAILED: {len(failed)} classes'}{f', skipped={skipped}' if skipped else ''})")
    if args.record:
        times = durations()
        times.update({r[0]: round(r[3], 1) for r in results})
        DURATIONS.write_text(json.dumps(dict(sorted(times.items())), indent=1) + "\n")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
