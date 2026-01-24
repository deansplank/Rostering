from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional

from ortools.sat.python import cp_model


# ----------------------------
# Data models
# ----------------------------

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

GONDOLA_SHIFTS = {"AM1", "AM2", "MC1_GON", "MC2", "PM1", "PM2"}
GS_SHIFTS = {"TILL1", "TILL2", "TILL3", "GATE", "FLOOR", "FLOOR2", "MC1_GS"}

PM_SHIFTS = {"PM1", "PM2"}
AM_SHIFTS = {"AM1", "AM2"}
SOFT_BLOCK_AFTER_PM = {"MC1_GON"}  # "preferably not"


@dataclass(frozen=True)
class Person:
    name: str
    can_gondola: bool
    can_gs: bool


@dataclass(frozen=True)
class PersonRules:
    allowed_shifts: Optional[set[str]] = None
    forbidden_shifts: set[str] = None


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
                    can_gondola=row["can_gondola"].strip() == "1",
                    can_gs=row["can_gs"].strip() == "1",
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

            allowed = set(s.strip() for s in allowed_raw.split("|") if s.strip()) or None
            forbidden = set(s.strip() for s in forbidden_raw.split("|") if s.strip())

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
                    shift=row["shift"].strip(),
                    weight=int(row["weight"]),
                )
            )
    return reqs


def load_week_template(path: Path) -> Dict[str, List[str]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    days = data["days"]
    weekday_shifts = data["weekday_shifts"]
    weekend_shifts = data["weekend_shifts"]
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
# Solver
# ----------------------------

def solve_week(
    people: List[Person],
    week: Dict[str, List[str]],
    rules: Dict[str, PersonRules],
    requests: List[Request],
    random_seed: int = 0,
) -> Dict[Tuple[str, str], str]:
    """
    Returns mapping (day, shift) -> person_name
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
                if is_shift_allowed(p, s, pr):
                    x[(p.name, d, s)] = model.NewBoolVar(f"x_{p.name}_{d}_{s}")

    # 1) Every shift slot filled exactly once
    for d, shifts in week.items():
        for s in shifts:
            candidates = [x[(pn, d, s)] for pn in person_names if (pn, d, s) in x]
            if not candidates:
                raise ValueError(f"No eligible staff for {d} {s} (check staff.csv / rules.csv)")
            model.Add(sum(candidates) == 1)

    # 2) One shift per person per day
    for pn in person_names:
        for d, shifts in week.items():
            vars_that_day = [x[(pn, d, s)] for s in shifts if (pn, d, s) in x]
            if vars_that_day:
                model.Add(sum(vars_that_day) <= 1)

    # 3) PM -> next day not AM
    for pn in person_names:
        for i, d in enumerate(DAYS[:-1]):
            next_d = DAYS[i + 1]
            if d not in week or next_d not in week:
                continue

            pm_vars = [x[(pn, d, s)] for s in week[d] if s in PM_SHIFTS and (pn, d, s) in x]
            am_vars = [x[(pn, next_d, s)] for s in week[next_d] if s in AM_SHIFTS and (pn, next_d, s) in x]

            if pm_vars and am_vars:
                # If any PM on day d, then no AM on next day.
                model.Add(sum(am_vars) == 0).OnlyEnforceIf(pm_vars)  # works because pm_vars are bools
                # BUT OnlyEnforceIf expects a single literal or list meaning AND.
                # We want: if (PM1 or PM2) then ...
                # So create helper var: worked_pm = OR(pm_vars)
                worked_pm = model.NewBoolVar(f"worked_pm_{pn}_{d}")
                model.AddMaxEquality(worked_pm, pm_vars)
                model.Add(sum(am_vars) == 0).OnlyEnforceIf(worked_pm)

    # Requests & soft penalties
    objective_terms = []

    # Fairness: try equal number of shifts per person
    total_slots = sum(len(shifts) for shifts in week.values())
    target = total_slots / len(people)

    # Count shifts per person
    for pn in person_names:
        vars_p = [var for (name, _, _), var in x.items() if name == pn]
        if not vars_p:
            continue
        count = model.NewIntVar(0, len(DAYS), f"count_{pn}")
        model.Add(count == sum(vars_p))

        # penalize deviation from target (scaled)
        # use abs via two vars
        dev = model.NewIntVar(0, len(DAYS), f"dev_{pn}")
        # dev >= count - floor(target), dev >= floor(target) - count
        t = int(round(target))
        model.Add(dev >= count - t)
        model.Add(dev >= t - count)
        objective_terms.append(dev * 10)

    # Soft rule: avoid MC1_GON the day after PM (penalty)
    for pn in person_names:
        for i, d in enumerate(DAYS[:-1]):
            next_d = DAYS[i + 1]
            if d not in week or next_d not in week:
                continue
            pm_vars = [x[(pn, d, s)] for s in week[d] if s in PM_SHIFTS and (pn, d, s) in x]
            next_mc1 = [x[(pn, next_d, s)] for s in week[next_d] if s in SOFT_BLOCK_AFTER_PM and (pn, next_d, s) in x]
            if pm_vars and next_mc1:
                worked_pm = model.NewBoolVar(f"worked_pm_soft_{pn}_{d}")
                model.AddMaxEquality(worked_pm, pm_vars)
                # If worked_pm and assigned MC1_GON next day -> penalty
                # Implement as: penalty_var >= worked_pm + mc1 - 1 (AND)
                pen = model.NewBoolVar(f"pen_pm_to_mc1_{pn}_{d}")
                model.Add(pen >= worked_pm + next_mc1[0] - 1)
                objective_terms.append(pen * 5)

    # Requests
    for r in requests:
        if r.day not in week:
            continue
        if r.name not in people_by_name:
            continue

        if r.type == "OFF":
            # hard constraint
            if r.shift == "ANY":
                # no assignment that day
                vars_that_day = [x[(r.name, r.day, s)] for s in week[r.day] if (r.name, r.day, s) in x]
                if vars_that_day:
                    model.Add(sum(vars_that_day) == 0)
            else:
                if (r.name, r.day, r.shift) in x:
                    model.Add(x[(r.name, r.day, r.shift)] == 0)

        elif r.type in ("WANT", "AVOID"):
            # soft penalties/rewards
            if r.shift == "ANY":
                # WANT any shift that day: reward if works; AVOID any: penalty if works
                vars_that_day = [x[(r.name, r.day, s)] for s in week[r.day] if (r.name, r.day, s) in x]
                if vars_that_day:
                    worked = model.NewBoolVar(f"worked_{r.name}_{r.day}")
                    model.AddMaxEquality(worked, vars_that_day)
                    if r.type == "WANT":
                        objective_terms.append((1 - worked) * r.weight)  # penalize NOT working
                    else:
                        objective_terms.append(worked * r.weight)        # penalize working
            else:
                if (r.name, r.day, r.shift) in x:
                    var = x[(r.name, r.day, r.shift)]
                    if r.type == "WANT":
                        objective_terms.append((1 - var) * r.weight)  # penalize not satisfying
                    else:
                        objective_terms.append(var * r.weight)        # penalize assigning

    # Add a tiny “randomness” jitter so repeated runs aren’t identical
    # (deterministic per seed if you keep seed stable)
    if random_seed != 0:
        # add small penalties based on hashing
        for (pn, d, s), var in x.items():
            jitter = (hash((pn, d, s, random_seed)) % 3)  # 0..2
            if jitter:
                objective_terms.append(var * jitter)

    model.Minimize(sum(objective_terms))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 10.0
    # solver.parameters.random_seed = random_seed  # (not always exposed in python cp-sat versions)

    status = solver.Solve(model)

    if status not in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        raise RuntimeError("No feasible roster found with current constraints/requests.")

    assignments: Dict[Tuple[str, str], str] = {}
    for d, shifts in week.items():
        for s in shifts:
            for pn in person_names:
                key = (pn, d, s)
                if key in x and solver.Value(x[key]) == 1:
                    assignments[(d, s)] = pn
                    break

    return assignments


def main():
    base = Path(".")
    people = load_staff(base / "staff.csv")
    rules = load_rules(base / "rules.csv")
    requests = load_requests(base / "requests.csv")
    week = load_week_template(base / "week_template.json")

    assignments = solve_week(people, week, rules, requests, random_seed=42)

    # Print simple output for now (Excel comes next)
    for d in DAYS:
        if d not in week:
            continue
        print(f"\n{d}")
        for s in week[d]:
            print(f"  {s:8s} -> {assignments[(d, s)]}")

if __name__ == "__main__":
    main()
