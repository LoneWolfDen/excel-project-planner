"""An independent Python implementation of docs/SCHEDULING_SPEC.md (B-1010).

It reads only the plan's inputs (the typed Tasks columns and the settings) and works out every task's dates, late
dates, slack and critical flag the way the spec says, without the planner's formulas. ReferenceScheduleTest
compares it with what the workbook calculated, task by task: if the formulas and the spec ever disagree, one of them
is wrong. Section numbers (§) refer to the spec.
"""
import datetime as dt
import math
import re

NO_LIMIT = dt.date(9999, 12, 31)
CONSTRAINTS = ["As Soon As Possible", "As Late As Possible", "Must Start On", "Must Finish On", "Start No Earlier Than",
               "Start No Later Than", "Finish No Earlier Than", "Finish No Later Than"]
LINKS = 4
LINK = re.compile(r"^(\d+)(FS|SS|FF|SF)?([+-]\d+(?:\.\d+)?)?D?$")
# template 1.2 (B-102): a lag may also be a % of the predecessor's duration (+50%) or elapsed calendar days (+3ed)
LINK_12 = re.compile(r"^(\d+)(FS|SS|FF|SF)?(?:([+-]\d+(?:\.\d+)?)(%|ED|D)?)?$")


class Calendar:
    """§1: working days are the week mask's working days (Mon..Sun, "1" = non-working) less the holidays."""

    def __init__(self, mask, holidays):
        self.off = {i for i, c in enumerate(mask) if c == "1"}
        self.holidays = set(holidays)

    def working(self, d):
        return d.weekday() not in self.off and d not in self.holidays

    def workday(self, d, n):
        """Excel's WORKDAY.INTL: n working days after d (before, if n < 0); d itself when n = 0."""
        n = int(n)                                   # Excel truncates a fractional day count
        step = 1 if n > 0 else -1
        while n:
            d += dt.timedelta(days=step)
            if self.working(d):
                n -= step
        return d

    def networkdays(self, a, b):
        if a > b:
            return -self.networkdays(b, a)
        return sum(1 for i in range((b - a).days + 1) if self.working(a + dt.timedelta(days=i)))

    def diff(self, a, b):
        """§1: signed working days from a to b (0 = same day)."""
        return self.networkdays(a, b) - 1 if b >= a else 1 - self.networkdays(b, a)


def parse_links(text, v12=False):
    """§2: the Predecessors text: up to LINKS items "UID[type][+/-lag]" (lag in working days, d optional; template
    1.2 also % and ed). [(uid, type, lag, unit) or None for an item that can't be read]; unit "" = working days."""
    items = [i for i in str(text or "").upper().replace(";", ",").replace(" ", "").split(",") if i][:LINKS]
    out = []
    for item in items:
        if v12:
            m = LINK_12.match(item)
            unit = "" if not m or m.group(4) in (None, "D") else m.group(4)
        else:
            m, unit = LINK.match(item), ""
        out.append((int(m.group(1)), m.group(2) or "FS", float(m.group(3) or 0), unit) if m else None)
    return out


def excel_round(x):
    """Excel's ROUND(x, 0): halves away from zero (Python's round() goes to even)."""
    return math.copysign(math.floor(abs(x) + 0.5), x)


def schedule(rows, settings, stages, v12=False):
    """rows: dicts of the typed Tasks columns in order (Level, Name, Mode, Dur, Pred, CType, Cons, DL, AS, AF, Pct,
    Stage, UID, BS, BF; template 1.2 also Active and RemDur). settings: ProjectStart, StatusDate, WorkWeek, Holidays,
    RescheduleAfterStatus, CriticalSlackDays. stages: {stage name: %}. v12: template 1.2's % and elapsed lags,
    inactive tasks and Remaining Duration (§2, §3, §6). Returns one dict per active row with the calculated
    values."""
    cal = Calendar(settings["WorkWeek"], settings["Holidays"])
    start, status = settings["ProjectStart"], settings["StatusDate"]
    resched = settings["RescheduleAfterStatus"] == "Yes"
    status_wd = cal.workday(status - dt.timedelta(days=1), 1)
    rows = [r for r in rows if r.get("Name")]
    n = len(rows)
    level = [max(1, int(r.get("Level") or 0)) for r in rows]
    # §3 outline: a row is a summary when the next row is deeper; its tasks are the deeper rows right below it
    summary = [i + 1 < n and level[i + 1] > level[i] for i in range(n)]
    # §3 (template 1.2): an inactive task (Active = No) isn't scheduled, and links to it are ignored
    inactive = [v12 and not summary[i] and str(rows[i].get("Active") or "").strip().lower() == "no" for i in range(n)]
    below = []
    for i in range(n):
        j = i + 1
        while j < n and level[j] > level[i]:
            j += 1
        below.append(range(i + 1, j))
    pos = {r["UID"]: i for i, r in enumerate(rows)}
    out = [dict() for _ in rows]

    def links(i):
        """§2: the usable links of row i: (predecessor row, type, lag). Ignored: unreadable, unknown UID, itself,
        below this task, a summary, any link on a summary."""
        if summary[i]:
            return []
        usable = []
        for link in parse_links(rows[i].get("Pred"), v12):
            if link is None:
                continue
            uid, kind, lag, unit = link
            p = pos.get(uid)
            if p is None or p == i or p > i or summary[p] or inactive[p]:
                continue
            if unit == "%":                         # §2: a share of the predecessor's duration, in working days
                lag = excel_round(lag / 100 * num(rows[p].get("Dur")))
            usable.append((p, kind, lag, unit == "ED"))
        return usable

    def on_or_after(d):
        return cal.workday(d - dt.timedelta(days=1), 1)

    def driver(i, p, kind, lag, s_key, f_key, elapsed=False):
        """§4: the earliest start link (p, kind, lag) allows row i. An elapsed lag counts calendar days, then the
        first working day on or after (FS on a task with work: the first working day after)."""
        dur = num(rows[i].get("Dur"))
        back = -max(0, dur - 1)
        if elapsed:
            e = dt.timedelta(days=int(lag))
            if kind == "FS":
                return on_or_after(out[p][f_key] + e) if dur == 0 else cal.workday(out[p][f_key] + e, 1)
            if kind == "SS":
                return on_or_after(out[p][s_key] + e)
            return cal.workday(on_or_after(out[p][f_key if kind == "FF" else s_key] + e), back)
        if kind == "FS":
            return cal.workday(out[p][f_key], (0 if dur == 0 else 1) + lag)
        if kind == "SS":
            return cal.workday(out[p][s_key], lag)
        if kind == "FF":
            return cal.workday(cal.workday(out[p][f_key], lag), back)
        return cal.workday(cal.workday(out[p][s_key], lag), back)

    for i, r in enumerate(rows):                    # §4 forward pass (plan and forecast), top to bottom
        if summary[i] or inactive[i]:
            continue
        o, dur, cons = out[i], num(r.get("Dur")), r.get("Cons")
        ctype = r.get("CType") or ""
        # §5 constraints: MS Project's codes; a date with no type = Start No Earlier Than
        if not cons:
            code = 1 if ctype == CONSTRAINTS[1] else 0
        else:
            code = CONSTRAINTS.index(ctype) if ctype in CONSTRAINTS else 4
        o["code"] = code
        c_start = None if not cons else (cal.workday(cons, -(dur - 1)) if code in (3, 6, 7) and dur > 1 else cons)
        low = c_start if code in (2, 3, 4, 6) and c_start else dt.date.min
        high = c_start if code in (2, 3, 5, 7) and c_start else NO_LIMIT
        o["high"] = high
        manual = r.get("Mode") == "Manual"
        pinned = c_start or start
        lk = links(i)
        o["links"] = lk

        def earliest(s_key, f_key, floor=dt.date.min):
            drivers = [driver(i, p, kind, lag, s_key, f_key, x) for p, kind, lag, x in lk]
            day = min(max([start, low] + drivers), high)
            return cal.workday(max(day, floor) - dt.timedelta(days=1), 1)

        o["CS"] = pinned if manual else earliest("CS", "CF")
        o["CF"] = o["CS"] if dur <= 0 else cal.workday(o["CS"], dur - 1)
        # §6 actuals and the status date
        pct = r.get("Pct")
        rem = r.get("RemDur") if v12 else None
        started = r.get("AS") and not r.get("AF") and isinstance(rem, (int, float)) and not isinstance(rem, bool)
        if started and not isinstance(pct, (int, float)):
            # §6 (template 1.2): a Remaining Duration sets the progress: done / (done + remaining) working days
            done = cal.networkdays(r["AS"], status - dt.timedelta(days=1)) if status > r["AS"] else 0
            pct = done / (done + max(0, rem)) if done + max(0, rem) else 1
        pct = pct if isinstance(pct, (int, float)) else stages.get(r.get("Stage") or "", 0)
        if r.get("AS"):
            o["FS"] = r["AS"]
        elif manual:
            o["FS"] = pinned
        else:
            o["FS"] = earliest("FS", "FF", status if resched else dt.date.min)
        natural = cal.workday(o["FS"], dur - 1)
        remaining = cal.workday(status_wd, max(0, math.ceil(round(dur * (1 - min(1, pct)), 9)) - 1))
        if r.get("AF"):
            o["FF"] = r["AF"]
        elif started:                               # §6 (1.2): the remaining working days from the status date
            left = math.ceil(max(0, rem))
            o["FF"] = max(o["FS"], cal.workday(status_wd, -1) if left == 0 else cal.workday(status_wd, left - 1))
        elif dur <= 0:
            o["FF"] = o["FS"]
        elif r.get("AS") and resched:
            o["FF"] = max(natural, remaining)
        else:
            o["FF"] = natural
        o["progress"] = 1 if r.get("AF") else min(1, pct)

    leaves = [i for i in range(n) if not summary[i] and not inactive[i]]
    finish = max(out[i]["FF"] for i in leaves)
    # §7 backward pass, bottom to top: each successor limits its predecessors' late finish
    limits = {i: [] for i in leaves}
    for i in reversed(leaves):
        r, o, dur = rows[i], out[i], num(rows[i].get("Dur"))
        high = o["high"]
        cons_limit = NO_LIMIT if high >= NO_LIMIT else (high if dur <= 1 else cal.workday(high, dur - 1))
        o["LF"] = min([finish, r.get("DL") or NO_LIMIT, cons_limit] + limits[i])
        o["LS"] = o["LF"] if dur <= 0 else cal.workday(o["LF"], -(dur - 1))
        o["TS"] = cal.diff(o["FF"], o["LF"])
        for p, kind, lag, elapsed in o["links"]:
            p_dur = num(rows[p].get("Dur"))
            if elapsed:                             # the mirror of the forward pass
                e = dt.timedelta(days=int(lag))
                one = dt.timedelta(days=1)
                if kind == "FS":
                    limit = cal.workday(o["LS"] - e + (one if dur == 0 else dt.timedelta(0)), -1)
                elif kind == "FF":
                    limit = cal.workday(o["LF"] - e + one, -1)
                else:
                    at = cal.workday((o["LS"] if kind == "SS" else o["LF"]) - e + one, -1)
                    limit = at if p_dur <= 0 else cal.workday(at, p_dur - 1)
                limits[p].append(limit)
                continue
            if kind == "FS":
                limit = cal.workday(o["LS"], -((0 if dur == 0 else 1) + lag))
            elif kind == "FF":
                limit = cal.workday(o["LF"], -lag)
            else:
                at = cal.workday(o["LS"] if kind == "SS" else o["LF"], -lag)
                limit = at if p_dur <= 0 else cal.workday(at, p_dur - 1)
            limits[p].append(limit)
    # §7 free slack: the smallest gap on an outgoing link, capped at total slack
    gaps = {i: [] for i in leaves}
    for i in leaves:
        for p, kind, lag, elapsed in out[i]["links"]:
            gaps[p].append(cal.diff(driver(i, p, kind, lag, "FS", "FF", elapsed), out[i]["FS"]))
    for i in leaves:
        o = out[i]
        o["FSl"] = min([o["TS"]] + gaps[i])
        o["Critical"] = o["progress"] < 1 and o["TS"] <= settings["CriticalSlackDays"]
        # §5 As Late As Possible: shown later by its free slack; nothing else moves
        shift = (max(0, o["FSl"]) if o["code"] == 1 and rows[i].get("Mode") != "Manual" and not rows[i].get("AS")
                 else 0)
        for key in ("CS", "CF", "FS", "FF"):
            o["shown_" + key] = cal.workday(o[key], shift) if shift else o[key]
    result = []
    for i, r in enumerate(rows):                    # §8 roll-ups and the planned dates
        if inactive[i]:
            continue
        if summary[i]:
            kids = [j for j in below[i] if not summary[j] and not inactive[j]]
            pick = lambda fn, key: fn(out[j][key] for j in kids)
            v = {"Start": pick(min, "shown_FS"), "Finish": pick(max, "shown_FF"), "CS": pick(min, "shown_CS"),
                 "CF": pick(max, "shown_CF"), "Late Start": pick(min, "LS"), "Late Finish": pick(max, "LF"),
                 "Total Slack": pick(min, "TS"), "Critical": any(out[j]["Critical"] for j in kids)}
        else:
            o = out[i]
            v = {"Start": o["shown_FS"], "Finish": o["shown_FF"], "CS": o["shown_CS"], "CF": o["shown_CF"],
                 "Late Start": o["LS"], "Late Finish": o["LF"], "Total Slack": o["TS"], "Free Slack": o["FSl"],
                 "Critical": o["Critical"]}
        v["Planned Start"] = r.get("BS") or v["CS"]
        v["Planned Finish"] = r.get("BF") or v["CF"]
        v["UID"] = r["UID"]
        result.append(v)
    return result, finish


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0
