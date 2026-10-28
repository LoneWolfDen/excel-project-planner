"""Reorder a plan's tasks so every predecessor is listed above its successor (B-1015), the one ordering rule the
schedule needs (docs/SCHEDULING_SPEC.md section 2). The workbook itself only explains a broken link (Link Check);
this does the moving, as a tool beside it, so the .xlsx never depends on it.

The outline stays intact: a task never leaves its phase. Tasks move only as whole subtrees, among their siblings.
A link between two tasks becomes a constraint between their ancestors that are siblings under the closest common
parent, then each set of siblings is sorted topologically, keeping the original order wherever the links allow.
UIDs don't change, so every link keeps pointing at the same task. When the outline makes it impossible (a phase
that needs another phase first, which needs the first), nothing is written and the message names the tasks.

The rebuild goes through the upgrade path (planner/upgrade.py): the tasks are read by header with everything else
typed in the workbook, reordered, and the latest template is built around them. The same code runs on the command
line (scripts/reorder_planner.py) and in the browser page docs/export.html.
"""
import heapq
import io
import re

ITEM = re.compile(r"^\s*(\d+)")


class ReorderError(Exception):
    """The links can't all point upwards without breaking the outline (shown as-is by the CLI and the page)."""


def predecessors(text):
    """UIDs named in a Predecessors cell (`12, 14SS+2d`); unreadable items are left to Link Check."""
    out = []
    for item in str(text or "").replace(";", ",").split(","):
        m = ITEM.match(item)
        if m:
            out.append(int(m.group(1)))
    return out


def order(tasks):
    """(tasks reordered, [names of the tasks that moved]). tasks: dicts with UID, Level, Name and PredIn, in sheet
    order. Raises ReorderError when the outline makes it impossible."""
    n = len(tasks)
    level = [max(1, int(t.get("Level") or 1)) for t in tasks]
    parent, stack = [None] * n, []
    for i in range(n):                                  # the closest row above that is one level up
        while stack and level[stack[-1]] >= level[i]:
            stack.pop()
        parent[i] = stack[-1] if stack else None
        stack.append(i)
    children = {}
    for i in range(n):
        children.setdefault(parent[i], []).append(i)
    is_summary = {i for i in range(n) if i in children}
    chain = []
    for i in range(n):                                  # each task's ancestors, itself first
        c, path = i, []
        while c is not None:
            path.append(c)
            c = parent[c]
        chain.append(path)
    by_uid = {t.get("UID"): i for i, t in enumerate(tasks)}
    after = {}                                          # sibling -> siblings that must come after it
    reasons = {}
    for s, t in enumerate(tasks):
        if s in is_summary:
            continue                                    # links on a summary are ignored by the schedule
        for uid in predecessors(t.get("PredIn")):
            p = by_uid.get(uid)
            if p is None or p == s or p in is_summary:
                continue                                # Link Check explains these; moving can't fix them
            mine = set(chain[s])
            common = next(a for a in chain[p] + [None] if a is None or a in mine)
            a = chain[p][chain[p].index(common) - 1] if common is not None else chain[p][-1]
            b = chain[s][chain[s].index(common) - 1] if common is not None else chain[s][-1]
            if a != b:
                after.setdefault(a, set()).add(b)
                reasons.setdefault((a, b), (p, s))

    def sort(siblings):
        """Kahn's algorithm, always taking the earliest original row that is free to go next."""
        mine = set(siblings)
        needs = {i: 0 for i in siblings}
        for a in siblings:
            for b in after.get(a, ()):
                if b in mine:
                    needs[b] += 1
        ready = [i for i in siblings if needs[i] == 0]
        heapq.heapify(ready)
        out = []
        while ready:
            a = heapq.heappop(ready)
            out.append(a)
            for b in sorted(after.get(a, ())):
                if b in mine:
                    needs[b] -= 1
                    if needs[b] == 0:
                        heapq.heappush(ready, b)
        if len(out) < len(siblings):
            stuck = _cycle([i for i in siblings if i not in out], after)
            pairs = [reasons[(a, b)] for a, b in zip(stuck, stuck[1:] + stuck[:1]) if (a, b) in reasons]
            detail = "; ".join(f"UID {tasks[s]['UID']} ({tasks[s].get('Name')}) needs UID {tasks[p]['UID']} "
                               f"({tasks[p].get('Name')})" for p, s in pairs[:4])
            names = ", ".join(str(tasks[i].get("Name")) for i in stuck)
            raise ReorderError(f"Can't put every predecessor above its successor without breaking the outline: "
                               f"{names} each need another of them first ({detail}). Change one of those links, "
                               "or move the tasks between phases, then try again.")
        return out

    result = []

    def walk(group):
        for i in sort(group):
            result.append(i)
            walk(children.get(i, []))

    walk(children.get(None, []))
    kept = _longest_rising(result)                     # the tasks that kept their order; the rest moved
    moved = [tasks[i].get("Name") for i in result if i not in kept]
    return [tasks[i] for i in result], moved


def _cycle(stuck, after):
    """One cycle among the siblings left over (each must come after the next): the ones to tell the user about."""
    mine = set(stuck)
    path, seen = [stuck[0]], {stuck[0]: 0}
    while True:
        nxt = min(b for b in after.get(path[-1], ()) if b in mine)
        if nxt in seen:
            return path[seen[nxt]:]
        seen[nxt] = len(path)
        path.append(nxt)


def _longest_rising(seq):
    """The longest subsequence of seq in increasing order (as a set): the rows that didn't need to move."""
    import bisect
    tails, tail_at, back = [], [], {}
    for x in seq:
        k = bisect.bisect_left(tails, x)
        back[x] = tail_at[k - 1] if k else None
        if k == len(tails):
            tails.append(x)
            tail_at.append(x)
        else:
            tails[k], tail_at[k] = x, x
    out, x = set(), tail_at[-1] if tail_at else None
    while x is not None:
        out.add(x)
        x = back[x]
    return out


def convert(data):
    """(xlsx bytes with the tasks reordered, notes) from a planner workbook's bytes."""
    from . import upgrade
    from .build import make_workbook
    project, copies, sizes, warnings, _version = upgrade.read_planner(data)
    project.tasks, moved = order(project.tasks)
    wb = make_workbook(sizes=sizes, data=project)
    upgrade.copy_inputs(wb, copies, warnings)
    out = io.BytesIO()
    wb.save(out)
    notes = ([f"Moved {len(moved)} tasks so every predecessor is above its successor: {', '.join(map(str, moved[:8]))}"
              + (" ..." if len(moved) > 8 else "") + ". UIDs are unchanged, so every link still points at the same "
              "task."] if moved else ["Every predecessor was already above its successor: nothing moved."])
    return out.getvalue(), notes + warnings
