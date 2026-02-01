from __future__ import annotations

import csv
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Tuple, Optional

from ortools.sat.python import cp_model


# ----------------------------
# Data models / constants
# ----------------------------

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

GONDOLA_SHIFTS = {"AM1", "AM2", "MC1_GON", "MC2", "PM1", "PM2"}
GS_SHIFTS = {"TILL1", "TILL2", "TILL3", "GATE", "FLOOR", "FLOOR2", "MC1_GS"}

PM_SHIFTS = {"PM1", "PM2"}
AM_SHIFTS = {"AM1", "AM2"}
SOFT_BLOCK_AFTER_PM = {"MC1_GON"}  # "preferably not"

# GS preference: try have at least one of these on GS each day (soft, not absolute)
GS_TEAM_LEADS = {"Jack Frith", "Romana Suhajdova"}


@dataclass(frozen=True)
class Person:
    name: str
    can_gondola: bool
    can_gs: bool


@dataclass(frozen=True)
class PersonRules:
    allowed_shifts: Optional[set[str]] = None
    forbidden_shifts: set[str] = field(default_factory=set)


@dataclass(frozen=True)
class Request:
    name: str
    day: str
    type: str   # OFF / WANT / AVOID
    shift: str  # shift id or ANY
    weight: int


# ----------------------------
# Loaders
# ----------------------------

def load_staff(path: Path) -> List[Person]:
    people: List[Person] = []
    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            people.append(
                Person(
                    name=row["name"].strip(),
                    can_gondola=str(row.get("can_gondola", "0")).strip() == "1",
                    can_gs=str(row.get("can_gs", "1")).strip() == "1",
                )
            )
    if not people:
        raise ValueError("staff.csv is empty")
    return people


def load_rules(path: Path) -> Dict[str, PersonRules]:
    rules: Dict[str, PersonRules] = {}
    if not path.exists():
        return rules

    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row["name"].strip()
            allowed_raw = (row.get("allowed_shifts") or "").strip()
            forbidden_raw = (row.get("forbidden_shifts") or "").strip()

            allowed = set(s.strip().upper() for s in allowed_raw.split("|") if s.strip()) or None
            forbidden = set(s.strip().upper() for s in forbidden_raw.split("|") if s.strip())

            rules[name] = PersonRules(allowed_shifts=allowed, forbidden_shifts=forbidden)
    return rules


def load_requests(path: Path) -> List[Request]:
    reqs: List[Request] = []
    if not path.exists():
        return reqs

    with path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            reqs.append(
                Request(
                    name=row["name"].strip(),
                    day=row["day"].strip(),
                    type=row["type"].strip().upper(),
                    shift=str(row.get("shift") or "ANY").strip().upper(),
                    weight=int(float(row.get("weight") or 10)),
                )
            )
    return reqs


def load_week_template(path: Path) -> Dict[str, List[str]]:
    """
    Your week_template.json format:
      {
        "days": ["Mon",...],
        "weekday_shifts": [...],
        "weekend_shifts": [...],
        "weekend_days": ["Sat","Sun"],
        "seasonal_optional": {"FLOOR2": true/false}
      }
    Returns: { "Mon": [...], "Tue": [...], ... }
    """
    data = json.loads(path.read_text(encoding="utf-8"))
    days = data["days"]
    weekday_shifts = [str(s).strip().upper() for s in data["weekday_shifts"]]
    weekend_shifts = [str(s).strip().upper() for s in data["weekend_shifts"]]
    weekend_days = set(data["weekend_days"])

    week: Dict[str, List[str]] = {}
    for d in days:
        week[d] = list(weekend_shifts if d in weekend_days else weekday_shifts)

    # Optional seasonal
    seasonal = data.get("seasonal_optional", {})
    if seasonal.get("FLOOR2") is True:
        for d in days:
            if "FLOOR2" not in week[d]:
                week[d].append("FLOOR2")

    return week


# ----------------------------
# Eligibility logic
# ----------------------------

def is_shift_allowed(person: Person, shift: str, pr: Optional[PersonRules]) -> bool:
    shift = shift.strip().upper()

    # Dept eligibility
    if shift in GONDOLA_SHIFTS and not person.can_gondola:
        return False
    if shift in GS_SHIFTS and not person.can_gs:
        return False

    # Per-person caveats
    if pr and pr.allowed_shifts is not None:
        return shift in pr.allowed_shifts
    if pr and pr.forbidden_shifts:
        return shift not in pr.forbidden_shifts

    return True


# ----------------------------
# Solver (best-effort with holes)
# ----------------------------

def solve_week(
    people: List[Person],
    week: Dict[str, List[str]],
    rules: Dict[str, PersonRules],
    requests: List[Request],
    random_seed: int = 0,
) -> Dict[Tuple[str, str], Optional[str]]:
    """
    Returns mapping (day, shift) -> person_name | None

    Behaviour:
      - Leaves holes (None) when needed (instead of crashing).
      - Hard: at most 1 shift per person per day (no double booking).
      - Hard: PM -> next day not AM.
      - Soft: minimise holes (very high weight), fairness, and GS Team Lead coverage.
    """
    model = cp_model.CpModel()

    person_names = [p.name for p in people]
    people_by_name = {p.name: p for p in people}

    # x[p, d, s] = 1 if person p works shift s on day d
    x: Dict[Tuple[str, str, str], cp_model.IntVar] = {}

    # build vars only for eligible assignments
    for p in people:
        pr = rules.get(p.name)
        for d, shifts in week.items():
            for s in shifts:
                s2 = str(s).strip().upper()
                if is_shift_allowed(p, s2, pr):
                    x[(p.name, d, s2)] = model.NewBoolVar(f"x_{p.name}_{d}_{s2}")

    # 1) Every shift slot is either filled by exactly one eligible person OR left unfilled.
    unfilled: Dict[Tuple[str, str], cp_model.IntVar] = {}
    for d, shifts in week.items():
        for s in shifts:
            s2 = str(s).strip().upper()
            candidates = [x[(pn, d, s2)] for pn in person_names if (pn, d, s2) in x]
            u = model.NewBoolVar(f"unfilled_{d}_{s2}")
            unfilled[(d, s2)] = u
            if candidates:
                model.Add(sum(candidates) + u == 1)
            else:
                model.Add(u == 1)

    # 2) One shift per person per day (hard)
    for pn in person_names:
        for d, shifts in week.items():
            vars_that_day = [x[(pn, d, str(s).strip().upper())] for s in shifts if (pn, d, str(s).strip().upper()) in x]
            if vars_that_day:
                model.Add(sum(vars_that_day) <= 1)

    # 3) PM -> next day not AM (hard)
    for pn in person_names:
        for i, d in enumerate(DAYS[:-1]):
            next_d = DAYS[i + 1]
            if d not in week or next_d not in week:
                continue

            pm_vars = [x[(pn, d, str(s).strip().upper())] for s in week[d] if str(s).strip().upper() in PM_SHIFTS and (pn, d, str(s).strip().upper()) in x]
            am_vars = [x[(pn, next_d, str(s).strip().upper())] for s in week[next_d] if str(s).strip().upper() in AM_SHIFTS and (pn, next_d, str(s).strip().upper()) in x]

            if pm_vars and am_vars:
                worked_pm = model.NewBoolVar(f"worked_pm_{pn}_{d}")
                model.AddMaxEquality(worked_pm, pm_vars)
                model.Add(sum(am_vars) == 0).OnlyEnforceIf(worked_pm)

    # Requests & soft penalties
    objective_terms = []

    # HARD priority: minimise unfilled slots first (huge weight)
    HOLE_PENALTY = 10000
    for u in unfilled.values():
        objective_terms.append(u * HOLE_PENALTY)

    # Fairness: try equal number of shifts per person (soft)
    total_slots = sum(len(shifts) for shifts in week.values())
    target = total_slots / max(1, len(people))

    for pn in person_names:
        vars_p = [var for (name, _, _), var in x.items() if name == pn]
        if not vars_p:
            continue
        count = model.NewIntVar(0, len(DAYS), f"count_{pn}")
        model.Add(count == sum(vars_p))

        dev = model.NewIntVar(0, len(DAYS), f"dev_{pn}")
        t = int(round(target))
        model.Add(dev >= count - t)
        model.Add(dev >= t - count)
        objective_terms.append(dev * 10)

    # Soft: avoid MC1_GON the day after PM (penalty)
    for pn in person_names:
        for i, d in enumerate(DAYS[:-1]):
            next_d = DAYS[i + 1]
            if d not in week or next_d not in week:
                continue
            pm_vars = [x[(pn, d, str(s).strip().upper())] for s in week[d] if str(s).strip().upper() in PM_SHIFTS and (pn, d, str(s).strip().upper()) in x]
            next_mc1 = [x[(pn, next_d, str(s).strip().upper())] for s in week[next_d] if str(s).strip().upper() in SOFT_BLOCK_AFTER_PM and (pn, next_d, str(s).strip().upper()) in x]
            if pm_vars and next_mc1:
                worked_pm = model.NewBoolVar(f"worked_pm_soft_{pn}_{d}")
                model.AddMaxEquality(worked_pm, pm_vars)

                pen = model.NewBoolVar(f"pen_pm_to_mc1_{pn}_{d}")
                model.Add(pen >= worked_pm + next_mc1[0] - 1)
                objective_terms.append(pen * 5)

    # Soft: GS Team Lead coverage each day (if possible)
    # Prefer at least one of GS_TEAM_LEADS on a GS shift each day.
    TL_PENALTY = 250
    for d in DAYS:
        if d not in week:
            continue
        tl_vars = []
        for tl in GS_TEAM_LEADS:
            for s in week[d]:
                s2 = str(s).strip().upper()
                if s2 in GS_SHIFTS and (tl, d, s2) in x:
                    tl_vars.append(x[(tl, d, s2)])
        if tl_vars:
            has_tl = model.NewBoolVar(f"has_tl_{d}")
            model.AddMaxEquality(has_tl, tl_vars)
            objective_terms.append((1 - has_tl) * TL_PENALTY)

    # Requests
    for r in requests:
        if r.day not in week:
            continue
        if r.name not in people_by_name:
            continue

        day = r.day
        shift = r.shift.strip().upper()

        if r.type == "OFF":
            if shift == "ANY":
                vars_that_day = [x[(r.name, day, str(s).strip().upper())] for s in week[day] if (r.name, day, str(s).strip().upper()) in x]
                if vars_that_day:
                    model.Add(sum(vars_that_day) == 0)
            else:
                if (r.name, day, shift) in x:
                    model.Add(x[(r.name, day, shift)] == 0)

        elif r.type in ("WANT", "AVOID"):
            w = max(0, int(r.weight))
            if shift == "ANY":
                vars_that_day = [x[(r.name, day, str(s).strip().upper())] for s in week[day] if (r.name, day, str(s).strip().upper()) in x]
                if vars_that_day:
                    worked = model.NewBoolVar(f"worked_{r.name}_{day}")
                    model.AddMaxEquality(worked, vars_that_day)
                    if r.type == "WANT":
                        objective_terms.append((1 - worked) * w)
                    else:
                        objective_terms.append(worked * w)
            else:
                if (r.name, day, shift) in x:
                    var = x[(r.name, day, shift)]
                    if r.type == "WANT":
                        objective_terms.append((1 - var) * w)
                    else:
                        objective_terms.append(var * w)

    # small deterministic jitter so rerolls differ a bit
    if random_seed:
        for (pn, d, s), var in x.items():
            jitter = (hash((pn, d, s, int(random_seed))) % 3)  # 0..2
            if jitter:
                objective_terms.append(var * jitter)

    model.Minimize(sum(objective_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10.0
    try:
        solver.parameters.random_seed = int(random_seed)
    except Exception:
        pass
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)

    # With unfilled vars, we *should* always be feasible; but just in case,
    # return all holes so the UI/export still works.
    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        return {(d, str(s).strip().upper()): None for d, shifts in week.items() for s in shifts}

    assignments: Dict[Tuple[str, str], Optional[str]] = {}
    for d, shifts in week.items():
        for s in shifts:
            s2 = str(s).strip().upper()
            assignments[(d, s2)] = None
            # if unfilled is 1, keep None
            if solver.Value(unfilled[(d, s2)]) == 1:
                continue
            for pn in person_names:
                key = (pn, d, s2)
                if key in x and solver.Value(x[key]) == 1:
                    assignments[(d, s2)] = pn
                    break

    return assignments
