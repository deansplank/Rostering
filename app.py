import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from generate_roster import solve_week, load_staff, load_rules, load_requests, load_week_template
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


def default_weight(row):
    t = str(row.get("type", "")).strip().upper()
    w = row.get("weight", None)
    if pd.isna(w) or str(w).strip() == "":
        return 999 if t == "OFF" else 10
    return int(w)


def sync_availability_with_staff(staff_df: pd.DataFrame) -> pd.DataFrame:
    """Ensure availability.csv includes exactly all staff (add missing, drop removed)."""
    staff_names = staff_df["name"].astype(str).str.strip().tolist()

    if AVAIL_PATH.exists():
        avail_df = pd.read_csv(AVAIL_PATH)
    else:
        avail_df = pd.DataFrame(columns=["name"] + DAYS)

    if "name" not in avail_df.columns:
        avail_df.insert(0, "name", "")

    for d in DAYS:
        if d not in avail_df.columns:
            avail_df[d] = 1

    avail_df["name"] = avail_df["name"].astype(str).str.strip()

    # Keep only staff that still exist
    avail_df = avail_df[avail_df["name"].isin(staff_names)]

    # Add missing staff (default available)
    existing = set(avail_df["name"].tolist())
    missing = [n for n in staff_names if n not in existing]
    if missing:
        add_rows = pd.DataFrame([{"name": n, **{d: 1 for d in DAYS}} for n in missing])
        avail_df = pd.concat([avail_df, add_rows], ignore_index=True)

    # Order to match staff.csv
    avail_df["__order"] = avail_df["name"].apply(lambda x: staff_names.index(x) if x in staff_names else 9999)
    avail_df = avail_df.sort_values("__order").drop(columns="__order").reset_index(drop=True)

    avail_df.to_csv(AVAIL_PATH, index=False)
    return avail_df


def detect_template_sheets(template_path: Path) -> tuple[str, str]:
    """
    Detects the Gondola and GS sheet names from the workbook so export never KeyErrors.
    - Gondola: name contains "gondola" (case-insensitive)
    - GS: name contains "guest" or "gs"
    Fallback: if there are exactly 2 sheets, pick the other as GS.
    """
    wb = load_workbook(template_path)
    sheets = wb.sheetnames

    gondola_sheet = None
    gs_sheet = None

    for s in sheets:
        ls = s.lower().strip()
        if gondola_sheet is None and "gondola" in ls:
            gondola_sheet = s
        if gs_sheet is None and ("guest" in ls or ls == "gs" or "gs " in ls or " gs" in ls or "gs_host" in ls or "gs host" in ls):
            gs_sheet = s

    # Fallbacks
    if gondola_sheet is None and sheets:
        gondola_sheet = sheets[0]

    if gs_sheet is None:
        if len(sheets) == 2:
            gs_sheet = sheets[1] if sheets[0] == gondola_sheet else sheets[0]
        elif sheets:
            # pick a non-gondola sheet if possible
            for s in sheets:
                if s != gondola_sheet:
                    gs_sheet = s
                    break
            if gs_sheet is None:
                gs_sheet = gondola_sheet

    return gondola_sheet, gs_sheet


def _build_name_to_row(ws, name_col: int = 1, start_row: int = 5, max_scan: int = 600) -> dict[str, int]:
    m: dict[str, int] = {}
    for r in range(start_row, start_row + max_scan):
        v = ws.cell(r, name_col).value
        if v is None:
            continue
        name = str(v).strip()
        if name:
            m[name] = r
    return m


def _build_day_to_col(ws, header_row: int = 4, start_col: int = 3, max_scan: int = 30) -> dict[str, int]:
    m: dict[str, int] = {}
    for c in range(start_col, start_col + max_scan):
        v = ws.cell(header_row, c).value
        if v is None:
            continue
        day = str(v).strip()
        if day:
            m[day] = c
    return m


def ensure_template_has_staff_rows(template_path: Path, staff_df: pd.DataFrame) -> None:
    """
    Keeps your existing template formatting.
    Adds missing staff names to the bottom of the correct sheet by copying the last staff row styles.
    Does NOT delete rows (removals are handled by clearing names optionally).
    """
    gondola_sheet, gs_sheet = detect_template_sheets(template_path)
    wb = load_workbook(template_path)
    ws_g = wb[gondola_sheet]
    ws_s = wb[gs_sheet]

    # Determine which names should be on each sheet
    staff_df = staff_df.copy()
    staff_df["name"] = staff_df["name"].astype(str).str.strip()

    gondola_names = staff_df.loc[staff_df["can_gondola"] == 1, "name"].tolist()
    gs_names = staff_df.loc[staff_df["can_gs"] == 1, "name"].tolist()

    def add_missing(ws, names):
        name_to_row = _build_name_to_row(ws)
        day_to_col = _build_day_to_col(ws)

        # Find last filled staff row (based on any name in col A)
        last_row = 4
        for r in sorted(name_to_row.values()):
            last_row = max(last_row, r)

        # If template has no staff rows at all, we still append starting at row 5
        template_row = last_row if last_row >= 5 else 5
        # Copy styles from this row if possible
        can_copy_style = template_row >= 5 and ws.cell(template_row, 1).value is not None

        for n in names:
            if n in name_to_row:
                continue

            new_row = last_row + 1
            last_row = new_row

            # Copy formatting from template_row into new_row
            if can_copy_style:
                # Copy style for name, notes, and day cells (C..)
                max_col = max(day_to_col.values()) if day_to_col else 3 + len(DAYS) - 1
                for c in range(1, max_col + 1):
                    src = ws.cell(template_row, c)
                    dst = ws.cell(new_row, c)
                    if src.has_style:
                        dst._style = src._style
                    dst.number_format = src.number_format
                    dst.protection = src.protection
                    dst.alignment = src.alignment

            # Write values
            ws.cell(new_row, 1).value = n
            # Clear shift cells
            for c in day_to_col.values():
                ws.cell(new_row, c).value = ""

    add_missing(ws_g, gondola_names)
    add_missing(ws_s, gs_names)

    wb.save(template_path)


# ============================================================
# PAGE SETUP
# ============================================================
st.set_page_config(page_title="Skyline Roster Generator", layout="wide")
st.title("🗓️ Skyline Roster Generator")

# Load staff fresh each run
staff_df = pd.read_csv(STAFF_PATH)
staff_df["name"] = staff_df["name"].astype(str).str.strip()

# Keep availability synced always
avail_df = sync_availability_with_staff(staff_df)


# ============================================================
# MANAGE STAFF
# ============================================================
st.header("Manage Staff")

with st.expander("➕ Add staff", expanded=False):
    with st.form("add_staff_form_unique", clear_on_submit=True):
        name = st.text_input("Full name").strip()
        can_gondola = st.checkbox("Can work Gondola", value=False)
        can_gs = st.checkbox("Can work Guest Services", value=True)
        submitted = st.form_submit_button("Add")

    if submitted:
        if not name:
            st.error("Name cannot be blank.")
        else:
            existing = set(staff_df["name"].str.lower().tolist())
            if name.lower() in existing:
                st.error("That name already exists.")
            else:
                staff_df.loc[len(staff_df)] = [name, int(can_gondola), int(can_gs)]
                staff_df.to_csv(STAFF_PATH, index=False)

                # Sync availability + ensure template has the new name
                staff_df = pd.read_csv(STAFF_PATH)
                staff_df["name"] = staff_df["name"].astype(str).str.strip()
                sync_availability_with_staff(staff_df)
                ensure_template_has_staff_rows(TEMPLATE_PATH, staff_df)

                st.success(f"Added {name}")
                st.rerun()

with st.expander("➖ Remove staff", expanded=False):
    remove_name = st.selectbox("Select staff member", [""] + staff_df["name"].tolist())
    confirm = st.checkbox("Confirm removal (this cannot be undone)")

    if st.button("Remove selected staff"):
        if not remove_name:
            st.error("Pick a staff member first.")
        elif not confirm:
            st.error("Tick the confirmation box.")
        else:
            staff_df = staff_df[staff_df["name"] != remove_name].reset_index(drop=True)
            staff_df.to_csv(STAFF_PATH, index=False)

            # Sync availability (template rows can stay; exporter ignores blank names)
            staff_df = pd.read_csv(STAFF_PATH)
            staff_df["name"] = staff_df["name"].astype(str).str.strip()
            sync_availability_with_staff(staff_df)

            st.success(f"Removed {remove_name}")
            st.rerun()

st.divider()


# ============================================================
# AVAILABILITY (Days Off)
# ============================================================
st.header("Availability (Days Off)")
st.caption("✅ ticked = available. Untick = day off (hard).")

avail_df = pd.read_csv(AVAIL_PATH)

avail_edit = avail_df.copy()
for d in DAYS:
    avail_edit[d] = avail_edit[d].fillna(1).astype(int).astype(bool)

edited_avail = st.data_editor(
    avail_edit,
    disabled=["name"],
    use_container_width=True,
    num_rows="fixed",
)

if st.button("💾 Save Availability"):
    save = edited_avail.copy()
    for d in DAYS:
        save[d] = save[d].fillna(True).astype(bool).astype(int)
    save.to_csv(AVAIL_PATH, index=False)
    st.success("Saved availability")

st.divider()


# ============================================================
# REQUESTS (dropdowns)
# ============================================================
st.header("Requests & Preferences")

people_for_dropdown = load_staff(STAFF_PATH)
name_options = [p.name for p in people_for_dropdown]

if REQ_PATH.exists():
    req_df = pd.read_csv(REQ_PATH)
else:
    req_df = pd.DataFrame(columns=["name", "day", "type", "shift", "weight"])

for col in ["name", "day", "type", "shift", "weight"]:
    if col not in req_df.columns:
        req_df[col] = ""

edited_req = st.data_editor(
    req_df,
    num_rows="dynamic",
    use_container_width=True,
    column_config={
        "name": st.column_config.SelectboxColumn("name", options=name_options, required=True),
        "day": st.column_config.SelectboxColumn("day", options=DAYS, required=True),
        "type": st.column_config.SelectboxColumn("type", options=TYPE_OPTIONS, required=True),
        "shift": st.column_config.SelectboxColumn("shift", options=SHIFT_OPTIONS, required=True),
        "weight": st.column_config.NumberColumn("weight", step=1, min_value=0),
    },
)

if len(edited_req) > 0:
    edited_req["type"] = edited_req["type"].fillna("")
    edited_req["weight"] = edited_req.apply(default_weight, axis=1)

if st.button("💾 Save Requests"):
    edited_req.to_csv(REQ_PATH, index=False)
    st.success("Saved requests")

st.divider()


# ============================================================
# GENERATE
# ============================================================
st.header("Generate Roster")
seed = st.slider("Randomness", 0, 1000, 42)

if st.button("🚀 Generate"):
    # Ensure template rows exist for current staff AND detect correct sheet names
    staff_now = pd.read_csv(STAFF_PATH)
    staff_now["name"] = staff_now["name"].astype(str).str.strip()
    ensure_template_has_staff_rows(TEMPLATE_PATH, staff_now)

    gondola_sheet, gs_sheet = detect_template_sheets(TEMPLATE_PATH)
    st.caption(f"Using template sheets: Gondola='{gondola_sheet}' | GS='{gs_sheet}'")

    # Reload runtime inputs
    people_rt = load_staff(STAFF_PATH)
    rules_rt = load_rules(RULES_PATH)
    week_rt = load_week_template(WEEK_PATH)

    # Availability -> OFF requests
    off_rows = []
    for _, r in edited_avail.iterrows():
        for d in DAYS:
            if not bool(r[d]):
                off_rows.append({"name": r["name"], "day": d, "type": "OFF", "shift": "ANY", "weight": 999})

    combined = pd.concat([edited_req, pd.DataFrame(off_rows)], ignore_index=True)
    runtime_req = BASE / "_requests_runtime.csv"
    combined.to_csv(runtime_req, index=False)

    requests = load_requests(runtime_req)

    assignments = solve_week(
        people=people_rt,
        week=week_rt,
        rules=rules_rt,
        requests=requests,
        random_seed=int(seed),
    )

    out = output_filename()

    # Export using detected sheet names (no more KeyError ever)
    export_roster_to_template(
        assignments=assignments,
        template_path=TEMPLATE_PATH,
        output_path=out,
        gondola_sheet=gondola_sheet,
        gs_sheet=gs_sheet,
        gondola_gs_label="GS Host",
    )

    with open(out, "rb") as f:
        st.download_button("⬇️ Download roster", f, file_name=out.name)
