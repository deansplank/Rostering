from __future__ import annotations

from pathlib import Path
from typing import Dict, Tuple

from openpyxl import load_workbook


# These should match the shift IDs used by your solver
GONDOLA_SHIFTS = {"AM1", "AM2", "MC1_GON", "MC2", "PM1", "PM2"}
GS_SHIFTS = {"TILL1", "TILL2", "TILL3", "GATE", "FLOOR", "FLOOR2", "MC1_GS"}


from openpyxl.cell.cell import MergedCell

def _write_value_safe(ws, row: int, col: int, value: str) -> None:
    """
    Writes to a cell, but if it's part of a merged range, writes to the top-left anchor.
    """
    cell = ws.cell(row, col)

    # If it's a merged cell, find its merged range and write to the anchor cell.
    if isinstance(cell, MergedCell):
        for merged in ws.merged_cells.ranges:
            if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
                ws.cell(merged.min_row, merged.min_col).value = value
                return
        # Fallback: do nothing if we can't find the merge range
        return

    cell.value = value






def _build_name_to_row(ws, name_col: int = 1, start_row: int = 5, max_scan: int = 300) -> dict[str, int]:
    """
    Reads names from column A starting at row 5 until blank, returns name -> row index.
    """
    m: dict[str, int] = {}
    for r in range(start_row, start_row + max_scan):
        v = ws.cell(r, name_col).value
        if v is None:
            continue
        name = str(v).strip()
        if not name:
            continue
        m[name] = r
    return m


def _build_day_to_col(ws, header_row: int = 4, start_col: int = 3, max_scan: int = 30) -> dict[str, int]:
    """
    Reads day headers from row 4 starting at column C, returns day -> column index.
    Expected headers like Mon Tue Wed Thu Fri Sat Sun.
    """
    m: dict[str, int] = {}
    for c in range(start_col, start_col + max_scan):
        v = ws.cell(header_row, c).value
        if v is None:
            continue
        day = str(v).strip()
        if day:
            m[day] = c
    return m


def export_roster_to_template(
    assignments: Dict[Tuple[str, str], str],
    template_path: str | Path,
    output_path: str | Path,
    gondola_sheet: str = "Gondola",
    gs_sheet: str = "Guest Services",
    gondola_gs_label: str = "GS Host",  # how GS shifts appear on Gondola sheet
) -> None:
    """
    assignments: (day, shift) -> person_name
    Writes shift codes into the template at (person row, day column).
    - Gondola sheet: writes gondola shifts; writes 'GS Host' if that gondola person is allocated a GS shift.
    - GS sheet: writes GS shift codes for GS staff.
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    wb = load_workbook(template_path)

    ws_g = wb[gondola_sheet]
    ws_s = wb[gs_sheet]

    # Build lookup maps from the template itself (no manual cell mapping needed)
    g_name_to_row = _build_name_to_row(ws_g)
    s_name_to_row = _build_name_to_row(ws_s)

    g_day_to_col = _build_day_to_col(ws_g)
    s_day_to_col = _build_day_to_col(ws_s)

    # Clear existing filled cells (optional, but helps when regenerating)
    def clear_sheet(ws, name_to_row, day_to_col):
        for _, r in name_to_row.items():
            for _, c in day_to_col.items():
                _write_value_safe(ws, r, c, "")


    clear_sheet(ws_g, g_name_to_row, g_day_to_col)
    clear_sheet(ws_s, s_name_to_row, s_day_to_col)

    # Convert (day, shift)->person into (person, day)->shift_code for each sheet
    for (day, shift), person in assignments.items():
        # Gondola sheet behaviour
        if person in g_name_to_row and day in g_day_to_col:
            r = g_name_to_row[person]
            c = g_day_to_col[day]
            if shift in GONDOLA_SHIFTS:
                _write_value_safe(ws_g, r, c, shift)
            elif shift in GS_SHIFTS:
                # show GS allocation on gondola roster like your screenshot
                _write_value_safe(ws_g, r, c, gondola_gs_label)
            else:
                # unknown shift - write raw
                _write_value_safe(ws_g, r, c, shift)


        # GS sheet behaviour
        if person in s_name_to_row and day in s_day_to_col:
            r = s_name_to_row[person]
            c = s_day_to_col[day]
            if shift in GS_SHIFTS:
                _write_value_safe(ws_s, r, c, shift)
            # If a GS staff member got a gondola shift, we ignore it on GS sheet

    wb.save(output_path)
