import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime

from generate_roster import solve_week, load_staff, load_rules, load_requests, load_week_template
from export_excel import export_roster_to_template

# ---------- CONFIG ----------
BASE = Path(".")
TEMPLATE_PATH = BASE / "roster_template_FINAL.xlsx"
DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

def output_filename():
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return BASE / f"Generated_Roster_{ts}.xlsx"

st.set_page_config(page_title="Skyline Roster Generator", layout="wide")
st.title("🗓️ Skyline Roster Generator")

# ---------- LOAD CORE DATA ----------
people = load_staff(BASE / "staff.csv")
rules = load_rules(BASE / "rules.csv")
week = load_week_template(BASE / "week_template.json")

# ============================================================
# AVAILABILITY (Days Off) GRID
# ============================================================
st.header("Availability (Days Off)")

avail_path = BASE / "availability.csv"

if avail_path.exists():
    avail_df = pd.read_csv(avail_path)
else:
    # Create default availability (everyone available every day)
    names = [p.name for p in people]
    avail_df = pd.DataFrame({"name": names})
    for d in DAYS:
        avail_df[d] = 1

# Ensure all required columns exist
if "name" not in avail_df.columns:
    avail_df.insert(0, "name", [p.name for p in people])

for d in DAYS:
    if d not in avail_df.columns:
        avail_df[d] = 1

# Convert 1/0 -> True/False so Streamlit shows checkboxes
avail_edit = avail_df.copy()
for d in DAYS:
    avail_edit[d] = (
        avail_edit[d]
        .fillna(1)        # default blanks to available
        .astype(int)
        .astype(bool)
    )


st.caption("✅ ticked = available. Untick = day off (hard rule).")

avail_edited = st.data_editor(
    avail_edit,
    use_container_width=True,
    disabled=["name"],
    num_rows="fixed",
)

if st.button("💾 Save Availability"):
    save_df = avail_edited.copy()
    for d in DAYS:
        save_df[d] = save_df[d].astype(bool).astype(int)
    save_df.to_csv(avail_path, index=False)
    st.success("Availability saved")

st.divider()

# ============================================================
# REQUESTS EDITOR (WANT / AVOID / OFF for specific shifts)
# ============================================================
st.header("Staff Requests & Preferences")

req_path = BASE / "requests.csv"

if req_path.exists():
    req_df = pd.read_csv(req_path)
else:
    req_df = pd.DataFrame(columns=["name", "day", "type", "shift", "weight"])

st.caption(
    "Type: OFF (hard), WANT or AVOID (soft). "
    "Shift can be ANY or a specific shift like AM1, PM2."
)

edited_req_df = st.data_editor(
    req_df,
    num_rows="dynamic",
    use_container_width=True
)

if st.button("💾 Save Requests"):
    edited_req_df.to_csv(req_path, index=False)
    st.success("Requests saved")

st.divider()

# ============================================================
# GENERATE
# ============================================================
st.header("Generate Roster")

random_seed = st.slider(
    "Randomness (change this to get a different but fair roster)",
    min_value=0,
    max_value=1000,
    value=42
)

if st.button("🚀 Generate Roster"):
    with st.spinner("Generating roster..."):
        # Save latest edits so Generate always uses what you see
        # Availability
        runtime_avail = avail_edited.copy()
        for d in DAYS:
            runtime_avail[d] = (
                runtime_avail[d]
                .fillna(True)
                .astype(bool)
                .astype(int)
            )


        # Build OFF requests from unticked availability
        off_rows = []
        for _, row in runtime_avail.iterrows():
            name = str(row["name"]).strip()
            for day in DAYS:
                if int(row[day]) == 0:
                    off_rows.append({
                        "name": name,
                        "day": day,
                        "type": "OFF",
                        "shift": "ANY",
                        "weight": 999
                    })

        # Combine manual requests with auto-generated OFF rows
        manual_df = edited_req_df.copy()

        # Ensure required columns exist
        if "weight" not in manual_df.columns:
            manual_df["weight"] = ""

        # Default missing weights:
        # - OFF → 999
        # - WANT / AVOID → 10
        manual_df["weight"] = manual_df.apply(
            lambda r: 999 if r.get("type") == "OFF" else (r.get("weight") if pd.notna(r.get("weight")) else 10),
            axis=1
        )

        off_df = pd.DataFrame(off_rows)

        combined_df = pd.concat([manual_df, off_df], ignore_index=True)


        # Write a runtime requests file for the solver
        runtime_req_path = BASE / "_requests_runtime.csv"
        combined_df.to_csv(runtime_req_path, index=False)

        # Load requests through your existing loader
        requests = load_requests(runtime_req_path)

        assignments = solve_week(
            people=people,
            week=week,
            rules=rules,
            requests=requests,
            random_seed=int(random_seed),
        )

        out_path = output_filename()

        export_roster_to_template(
            assignments=assignments,
            template_path=TEMPLATE_PATH,
            output_path=out_path,
            gondola_gs_label="GS Host",
        )

    st.success("Roster generated!")

    with open(out_path, "rb") as f:
        st.download_button(
            "⬇️ Download Excel Roster",
            f,
            file_name=out_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
