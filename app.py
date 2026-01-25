import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import shutil
from copy import copy

from openpyxl import load_workbook

from generate_roster import (
    solve_week,
    load_staff,
    load_rules,
    load_requests,
    load_week_template,
)
from export_excel import export_roster_to_template


# ============================================================
# CONFIG
# ============================================================
BASE = Path(".")
STAFF_PATH = BASE / "staff.csv"
RULES_PATH = BASE / "rules.csv"
WEEK_PATH = BASE / "week_template.json"
AVAIL_PATH = BASE / "availability.csv"
REQ_PATH = BASE / "requests.csv"
TEMPLATE_PATH = BASE / "roster_template_FINAL.xlsx"

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

SHIFT_OPTIONS = [
    "ANY",
    "AM1", "AM2", "MC1_GON", "MC2", "PM1", "PM2",
    "TILL1", "TILL2", "TILL3", "GATE", "FLOOR", "FLOOR2", "MC1_GS",
]
TYPE_OPTIONS = ["OFF", "WANT", "AVOID"]


# ============================================================
# HELPERS
# ============================================================
def output_filename():
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return BASE / f"Generated_Roster_{ts}.xlsx"


def ensure_staff_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure staff.csv supports capability + visibility.
    Adds show_gondola/show_gs if missing, defaulting to capability.
    """
    if "can_gondola" not in df.columns:
        df["can_gondola"] = 0
    if "can_gs" not in df.columns:
        df["can_gs"] = 1
    if "show_gondola" not in df.columns:
        df["show_gondola"] = df["can_gondola"]
    if "show_gs" not in df.columns:
        df["show_gs"] = df["can_gs"]

    # Normalise types
    df["can_gondola"] = df["can_gondola"].fillna(0).astype(int)
    df["can_gs"] = df["can_gs"].fillna(1).astype(int)
    df["show_gondola"] = df["show_gondola"].fillna(df["can_gondola"]).astype(int)
    df["show_gs"] = df["show_gs"].fillna(df["can_gs"]).astype(int)

    return df


def sync_availability(staff_df: pd.DataFrame) -> pd.DataFrame:
    """Keep availability.csv aligned with staff.csv."""
    names = staff_df["name"].tolist()

    if AVAIL_PATH.exists():
        avail = pd.read_csv(AVAIL_PATH)
    else:
        avail = pd.DataFrame(columns=["name"] + DAYS)

    if "name" not in avail.columns:
        avail.insert(0, "name", "")

    for d in DAYS:
        if d not in avail.columns:
            avail[d] = 1

    avail["name"] = avail["name"].astype(str).str.strip()
    avail = avail[avail["name"].isin(names)]

    existing = set(avail["name"].tolist())
    missing = [n for n in names if n not in existing]
    if missing:
        add_rows = pd.DataFrame([{"name": n, **{d: 1 for d in DAYS}} for n in missing])
        avail = pd.concat([avail, add_rows], ignore_index=True)

    # Keep same order as staff.csv
    avail["__order"] = avail["name"].apply(lambda x: names.index(x) if x in names else 9999)
    avail = avail.sort_values("__order").drop(columns="__order").reset_index(drop=True)

    avail.to_csv(AVAIL_PATH, index=False)
    return avail


def detect_sheets(template: Path):
    """
    Detects sheet names from the workbook.
    Prefers 'Gondola' / 'Guest Services' / 'GS Host' etc.
    """
    wb = load_workbook(template)
    sheets = wb.sheetnames

    gondola = next((s for s in sheets if "gondola" in s.lower()), sheets[0])
    gs = next(
        (s for s in sheets if "guest" in s.lower() or s.lower().startswith("gs")),
        sheets[1] if len(sheets) > 1 else sheets[0],
    )
    return gondola, gs


def make_runtime_template() -> Path:
    """Copy the master template to a runtime copy so Excel locks never break generation."""
    rt = BASE / "_template_runtime.xlsx"
    shutil.copyfile(TEMPLATE_PATH, rt)
    return rt


def _sheet_day_cols(ws):
    """Map 'Mon'..'Sun' -> column index from header row 4."""
    day_to_col = {}
    for c in range(1, 60):
        v = ws.cell(4, c).value
        if v is None:
            continue
        s = str(v).strip()
        if s in DAYS:
            day_to_col[s] = c
    return day_to_col


def _sheet_name_rows(ws, start_row=5, name_col=1, max_scan=800):
    """Map staff name -> row index."""
    name_to_row = {}
    for r in range(start_row, start_row + max_scan):
        v = ws.cell(r, name_col).value
        if v is None:
            continue
        name = str(v).strip()
        if name:
            name_to_row[name] = r
    return name_to_row


def _find_last_staff_row(ws, start_row=5, name_col=1, max_scan=800):
    last = start_row - 1
    for r in range(start_row, start_row + max_scan):
        v = ws.cell(r, name_col).value
        if v is not None and str(v).strip():
            last = r
    return max(last, start_row)


def _copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    """Copy cell style/formatting from src_row to dst_row (values not copied)."""
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        # openpyxl styles can be StyleProxy objects; copy() materializes a concrete style
        if src.has_style:
            dst._style = copy(src._style)

def apply_visibility_to_runtime_template(runtime_template: Path, gondola_sheet: str, gs_sheet: str, staff_df: pd.DataFrame):
    """
    On the *runtime* template copy only:
    - Remove hidden staff from each sheet (clear name + day cells)
    - Ensure visible staff exist (append to bottom, copying formatting)
    """
    staff_df = staff_df.copy()
    staff_df["name"] = staff_df["name"].astype(str).str.strip()
    staff_df = ensure_staff_columns(staff_df)

    show_gondola = [n for n in staff_df.loc[staff_df["show_gondola"] == 1, "name"].tolist() if n]
    show_gs = [n for n in staff_df.loc[staff_df["show_gs"] == 1, "name"].tolist() if n]

    wb = load_workbook(runtime_template)
    ws_g = wb[gondola_sheet]
    ws_s = wb[gs_sheet]

    def process_sheet(ws, allowed_names):
        allowed_set = set(allowed_names)
        day_cols = _sheet_day_cols(ws)
        name_rows = _sheet_name_rows(ws)

        # 1) Clear names that are not allowed on this sheet
        for name, r in list(name_rows.items()):
            if name not in allowed_set:
                ws.cell(r, 1).value = ""
                for c in day_cols.values():
                    ws.cell(r, c).value = ""

        # Re-scan after clears
        name_rows = _sheet_name_rows(ws)

        # 2) Append missing allowed names (copy formatting from last staff row)
        last_staff_row = _find_last_staff_row(ws)
        # pick a formatting template row (last existing row with a name, or row 5)
        template_row = last_staff_row

        max_col = max(day_cols.values()) if day_cols else 3 + len(DAYS) - 1

        for name in allowed_names:
            if name in name_rows:
                continue
            new_row = last_staff_row + 1
            last_staff_row = new_row

            _copy_row_style(ws, template_row, new_row, max_col)
            ws.cell(new_row, 1).value = name
            # clear day cells (keep formatting)
            for c in day_cols.values():
                ws.cell(new_row, c).value = ""

    process_sheet(ws_g, show_gondola)
    process_sheet(ws_s, show_gs)

    wb.save(runtime_template)


# ============================================================
# STREAMLIT SETUP
# ============================================================
st.set_page_config(page_title="Skyline Roster Generator", layout="wide")
st.title("🗓️ Skyline Roster Generator")

# ============================================================
# LOAD CORE DATA
# ============================================================
staff_df = pd.read_csv(STAFF_PATH)
staff_df["name"] = staff_df["name"].astype(str).str.strip()
staff_df = ensure_staff_columns(staff_df)
staff_df.to_csv(STAFF_PATH, index=False)

avail_df = sync_availability(staff_df)

# ============================================================
# MANAGE STAFF
# ============================================================
st.header("Manage Staff")

with st.expander("➕ Add staff"):
    with st.form("add_staff", clear_on_submit=True):
        name = st.text_input("Full name").strip()

        can_gondola = st.checkbox("Can work Gondola", value=False)
        can_gs = st.checkbox("Can work Guest Services", value=True)

        show_gondola = st.checkbox("Show on Gondola roster", value=can_gondola)
        show_gs = st.checkbox("Show on Guest Services roster", value=can_gs)

        submit = st.form_submit_button("Add staff")

    if submit:
        if not name:
            st.error("Name required.")
        elif name.lower() in staff_df["name"].str.lower().tolist():
            st.error("Staff member already exists.")
        else:
            staff_df.loc[len(staff_df)] = {
                "name": name,
                "can_gondola": int(can_gondola),
                "can_gs": int(can_gs),
                "show_gondola": int(show_gondola),
                "show_gs": int(show_gs),
            }
            staff_df.to_csv(STAFF_PATH, index=False)
            sync_availability(staff_df)
            st.success(f"Added {name}")
            st.rerun()

with st.expander("➖ Remove staff"):
    remove = st.selectbox("Select staff", [""] + staff_df["name"].tolist())
    confirm = st.checkbox("Confirm removal")

    if st.button("Remove staff"):
        if not remove or not confirm:
            st.error("Select a staff member and confirm.")
        else:
            staff_df = staff_df[staff_df["name"] != remove].reset_index(drop=True)
            staff_df.to_csv(STAFF_PATH, index=False)
            sync_availability(staff_df)
            st.success(f"Removed {remove}")
            st.rerun()

st.divider()

# ============================================================
# AVAILABILITY
# ============================================================
st.header("Availability (Days Off)")
st.caption("Tick = available, untick = day off (hard rule)")

avail_edit = avail_df.copy()
for d in DAYS:
    avail_edit[d] = avail_edit[d].fillna(1).astype(int).astype(bool)

edited_avail = st.data_editor(
    avail_edit, disabled=["name"], use_container_width=True
)

if st.button("💾 Save Availability"):
    save = edited_avail.copy()
    for d in DAYS:
        save[d] = save[d].fillna(True).astype(bool).astype(int)
    save.to_csv(AVAIL_PATH, index=False)
    st.success("Availability saved")

st.divider()

# ============================================================
# REQUESTS
# ============================================================
st.header("Requests")

if REQ_PATH.exists():
    req_df = pd.read_csv(REQ_PATH)
else:
    req_df = pd.DataFrame(columns=["name", "day", "type", "shift", "weight"])

req_df = st.data_editor(
    req_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "name": st.column_config.SelectboxColumn("name", options=staff_df["name"].tolist()),
        "day": st.column_config.SelectboxColumn("day", options=DAYS),
        "type": st.column_config.SelectboxColumn("type", options=TYPE_OPTIONS),
        "shift": st.column_config.SelectboxColumn("shift", options=SHIFT_OPTIONS),
        "weight": st.column_config.NumberColumn("weight", min_value=0, step=1),
    },
)

if st.button("💾 Save Requests"):
    req_df.to_csv(REQ_PATH, index=False)
    st.success("Requests saved")

st.divider()

# ============================================================
# GENERATE ROSTER
# ============================================================
st.header("Generate Roster")
seed = st.slider("Randomness", 0, 1000, 42)

if st.button("🚀 Generate"):
    # 1) Make runtime template copy
    runtime_template = make_runtime_template()

    # 2) Detect sheet names (Gondola / GS Host / Guest Services etc)
    gondola_sheet, gs_sheet = detect_sheets(runtime_template)

    # 3) Apply visibility rules to runtime template (THIS is what stops gondola staff showing on GS)
    apply_visibility_to_runtime_template(runtime_template, gondola_sheet, gs_sheet, staff_df)

    # 4) Load model inputs
    people = load_staff(STAFF_PATH)
    rules = load_rules(RULES_PATH)
    week = load_week_template(WEEK_PATH)

    # Availability → OFF requests
    off_rows = []
    for _, r in edited_avail.iterrows():
        for d in DAYS:
            if not bool(r[d]):
                off_rows.append(
                    {"name": r["name"], "day": d, "type": "OFF", "shift": "ANY", "weight": 999}
                )

    combined = pd.concat([req_df, pd.DataFrame(off_rows)], ignore_index=True)
    tmp_req = BASE / "_requests_runtime.csv"
    combined.to_csv(tmp_req, index=False)

    requests = load_requests(tmp_req)

    # 5) Solve
    assignments = solve_week(
        people=people,
        week=week,
        rules=rules,
        requests=requests,
        random_seed=int(seed),
    )

    # 6) Export
    out = output_filename()

    export_roster_to_template(
        assignments=assignments,
        template_path=runtime_template,
        output_path=out,
        gondola_sheet=gondola_sheet,
        gs_sheet=gs_sheet,
        gondola_gs_label="GS Host",
    )

    with open(out, "rb") as f:
        st.download_button("⬇️ Download roster", f, file_name=out.name)