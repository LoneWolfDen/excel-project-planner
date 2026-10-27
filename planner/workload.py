"""Work totals over time, shared by the Agile sheet (sprint planned work) and the Resource Report (weekly and
sprint grids): how many hours of task work fall on a range of working days, for everyone or for one resource."""
from .tasks import people


def work_upto(d, resource=None):
    """Hours of task work scheduled up to and including working day d (counted from the project start).

    W(D) = SUMIFS(work, finish < D) + D * SUMIFS(work per day, running on D) - SUMIFS(work before start, running on D).
    Three SUMIFS are far cheaper than a SUMPRODUCT over every task, and SUMIFS needs no array formulas. They run
    once per person slot of a task (tasks.PERSON_FIELDS), each filtered by that slot's resource."""
    parts = []
    for work, wpd, before, who in zip(people("Work"), people("WPD"), people("WBefore"), people("Name")):
        res = f",{who},{resource}" if resource else ""
        running = f'tsk_StartWD,"<="&{d},tsk_FinishWD,">="&{d}{res}'
        part = (f'SUMIFS({work},tsk_FinishWD,"<"&{d},tsk_FinishWD,">0"{res})'
                f'+{d}*SUMIFS({wpd},{running})-SUMIFS({before},{running})')
        # a slot no task uses is skipped (IF only evaluates the branch it takes), so extra people cost nothing
        # until a plan has them
        parts.append(part if not parts else f"IF(MaxPeople>={len(parts) + 1},{part},0)")
    return "(" + "+".join(parts) + ")"


def work_between(first, last, resource=None):
    """Hours of task work scheduled on working days first..last (numbers counted from the project start, worked out
    once per sprint in SprintTable's _FirstWD / _LastWD): the difference of two cumulative totals (work_upto)."""
    return f"({work_upto(last, resource)}-{work_upto(f'({first}-1)', resource)})"
