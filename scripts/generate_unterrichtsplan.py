#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_unterrichtsplan.py
============================
Generates a structured Unterrichtsplan Excel file from a Zeitplan and Modulplan.

Requirements: pip install openpyxl pandas

Usage:
    python generate_unterrichtsplan.py \
        --zeitplan "Zeitplan_HH.xlsx" \
        --modulplan "Modulplan_KBM.xlsx" \
        --kursname "KBM" \
        --output "Unterrichtsplan_KBM_2026.xlsx" \
        --ue-pro-tag 9
"""

import argparse
import math
import os
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import Rule
from openpyxl.styles.differential import DifferentialStyle

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

HEADER_BG = "1F4E79"
HEADER_FG = "FFFFFF"
ALT_ROW_BG = "DDEEFF"
META_BG = "F2F2F2"
SUMME_BG = "D9D9D9"
THIN_BORDER_COLOR = "BFBFBF"

# Column widths for Sheet 1
COL_WIDTHS_S1 = {"A": 10, "B": 46, "C": 13, "D": 13, "E": 8, "F": 8, "G": 32}

# Day-type colour map for Sheet 3 conditional formatting
DAY_TYPE_COLORS: Dict[str, Tuple[str, str]] = {
    "Fe":   ("92D050", "000000"),
    "Pr":   ("FFD966", "000000"),
    "Prüf": ("FF7070", "FFFFFF"),
    "Ft":   ("FFB6C1", "000000"),
    "M0":   ("C6EFCE", "000000"),
    "M1":   ("FFEB9C", "000000"),
    "M2":   ("FFC7CE", "000000"),
    "M3":   ("BDD7EE", "000000"),
    "M4":   ("E2EFDA", "000000"),
    "M5":   ("FCE4D6", "000000"),
    "M6":   ("DAEEF3", "000000"),
    "M7":   ("EBF1DE", "000000"),
    "M8":   ("F2DCDB", "000000"),
    "M9":   ("FFF2CC", "000000"),
    "M10":  ("DDEBF7", "000000"),
    "M11":  ("D9EAD3", "000000"),
    "M12":  ("FCE4D6", "000000"),
    "M13":  ("C9C9FF", "000000"),
    "M14":  ("D0E4F5", "000000"),
    "M15":  ("FFF8DC", "000000"),
    "M16":  ("E8D5FF", "000000"),
    "IM0":  ("D9B2FF", "000000"),
    "IM1":  ("9DC3E6", "000000"),
    "IM2":  ("F4B183", "000000"),
    "IM3":  ("FF9999", "000000"),
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _thin_border() -> Border:
    side = Side(style="thin", color=HEADER_BG)
    return Border(left=side, right=side, top=side, bottom=side)


def _header_fill(hex_color: str = HEADER_BG) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _make_font(bold: bool = False, size: int = 10, color: str = "000000",
               italic: bool = False) -> Font:
    return Font(name="Arial", bold=bold, size=size, color=color, italic=italic)


def _excel_serial_to_date(serial: float) -> Optional[date]:
    """Convert an Excel date serial number to a Python date."""
    try:
        serial = int(serial)
        # Excel incorrectly treats 1900 as a leap year; offset accordingly
        if serial >= 60:
            serial -= 1
        delta = timedelta(days=serial - 1)
        return date(1900, 1, 1) + delta
    except (ValueError, TypeError, OverflowError):
        return None


def _normalize_col_name(name: str) -> str:
    if not isinstance(name, str):
        return ""
    return name.strip().lower()


# ---------------------------------------------------------------------------
# Zeitplan parsing
# ---------------------------------------------------------------------------

def _detect_year_sheets(wb: openpyxl.Workbook) -> List[str]:
    """Return sheet names that look like annual calendar sheets, in order."""
    annual_patterns = [
        "1. jahr", "2. jahr", "3. jahr", "4. jahr",
        "1.jahr", "2.jahr", "3.jahr", "4.jahr",
    ]
    result: List[str] = []
    # First try canonical year names
    for name in wb.sheetnames:
        if name.strip().lower() in annual_patterns:
            result.append(name)
    if result:
        return result
    # Fall back: sheets whose name is a 4-digit year
    for name in wb.sheetnames:
        if name.strip().isdigit() and 2000 <= int(name.strip()) <= 2100:
            result.append(name)
    return result


def _detect_feiertage(wb: openpyxl.Workbook, bundesland: str) -> set:
    """
    Parse the 'Schulfreie Tage' sheet and return a set of holiday dates
    for the given Bundesland column (e.g. 'HH').
    """
    holidays: set = set()
    target = None
    for name in wb.sheetnames:
        if "schulfreie" in name.lower() or "feiertag" in name.lower():
            target = name
            break
    if target is None:
        return holidays

    ws = wb[target]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return holidays

    # Find the header row containing bundesland column
    header_row_idx = None
    bl_col_idx = None
    date_col_idx = None

    for row_idx, row in enumerate(rows):
        for col_idx, cell in enumerate(row):
            if isinstance(cell, str) and cell.strip().upper() == bundesland.upper():
                bl_col_idx = col_idx
                header_row_idx = row_idx
                break
        if bl_col_idx is not None:
            break

    if header_row_idx is None or bl_col_idx is None:
        return holidays

    # Find column with dates (first column with date-like values after header)
    for col_idx, cell in enumerate(rows[header_row_idx]):
        if isinstance(cell, str) and ("datum" in cell.lower() or "date" in cell.lower()):
            date_col_idx = col_idx
            break
    if date_col_idx is None:
        date_col_idx = 0

    for row in rows[header_row_idx + 1:]:
        if not row or row[date_col_idx] is None:
            continue
        raw_date = row[date_col_idx]
        if isinstance(raw_date, date):
            holiday_date = raw_date
        elif isinstance(raw_date, (int, float)):
            holiday_date = _excel_serial_to_date(raw_date)
        else:
            continue
        if holiday_date is None:
            continue
        # Check if this bundesland has an 'x' or similar marker
        if bl_col_idx < len(row) and row[bl_col_idx]:
            marker = str(row[bl_col_idx]).strip().lower()
            if marker and marker not in ("0", "nein", "no", "false"):
                holidays.add(holiday_date)

    return holidays


def _parse_annual_sheet(
    ws, holidays: set
) -> List[Tuple[date, Optional[str]]]:
    """
    Parse one annual-calendar worksheet and return (date, day_type) tuples.
    day_type: None = teaching day, 'Fe', 'Pr', 'Prüf', 'Ft', 'WE' (weekend)
    """
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # ---- Find month header row: a row where at least 3 cells are Excel date
    #      serials (numeric) in the range of valid years, or datetime objects.
    month_row_idx = None
    month_col_map: Dict[int, date] = {}  # col_index -> first-day-of-month date

    for row_idx, row in enumerate(rows[:10]):
        month_candidates: Dict[int, date] = {}
        for col_idx, cell in enumerate(row):
            if cell is None:
                continue
            if isinstance(cell, (date,)):
                month_candidates[col_idx] = date(cell.year, cell.month, 1) if not isinstance(cell, type(None)) else None
            elif isinstance(cell, float) and 10000 < cell < 60000:
                d = _excel_serial_to_date(cell)
                if d:
                    month_candidates[col_idx] = date(d.year, d.month, 1)
            elif isinstance(cell, int) and 10000 < cell < 60000:
                d = _excel_serial_to_date(cell)
                if d:
                    month_candidates[col_idx] = date(d.year, d.month, 1)
        if len(month_candidates) >= 3:
            month_row_idx = row_idx
            month_col_map = month_candidates
            break

    if month_row_idx is None:
        # Fallback: look for cells with month names
        month_names_de = {
            "januar": 1, "februar": 2, "märz": 3, "april": 4,
            "mai": 5, "juni": 6, "juli": 7, "august": 8,
            "september": 9, "oktober": 10, "november": 11, "dezember": 12,
            "jan": 1, "feb": 2, "mär": 3, "apr": 4,
            "jun": 6, "jul": 7, "aug": 8, "sep": 9,
            "okt": 10, "nov": 11, "dez": 12,
        }
        for row_idx, row in enumerate(rows[:10]):
            month_candidates: Dict[int, Tuple[int, int]] = {}
            for col_idx, cell in enumerate(row):
                if isinstance(cell, str):
                    key = cell.strip().lower()[:3]
                    if key in month_names_de:
                        # Try to infer year from adjacent cells
                        month_candidates[col_idx] = month_names_de[cell.strip().lower()[:3] if len(cell) >= 3 else cell.strip().lower()]
            if len(month_candidates) >= 3:
                month_row_idx = row_idx
                # We don't have year info; skip this sheet gracefully
                print(f"  [WARN] Sheet '{ws.title}': month names found but year cannot be determined. Skipping.")
                return []

    if month_row_idx is None or not month_col_map:
        print(f"  [WARN] Sheet '{ws.title}': no month header row found. Skipping.")
        return []

    # ---- Find day column: first column before month columns that has integers 1-31
    day_col_idx = None
    min_month_col = min(month_col_map.keys())
    for col_idx in range(min_month_col):
        col_vals = [rows[r][col_idx] for r in range(month_row_idx + 1, min(month_row_idx + 35, len(rows)))]
        int_vals = [v for v in col_vals if isinstance(v, (int, float)) and 1 <= v <= 31]
        if len(int_vals) >= 20:
            day_col_idx = col_idx
            break
    if day_col_idx is None:
        # Use the column just before the first month column
        day_col_idx = max(0, min_month_col - 1)

    # ---- Iterate data rows (day rows)
    calendar: List[Tuple[date, Optional[str]]] = []
    sorted_month_cols = sorted(month_col_map.keys())

    for row_idx in range(month_row_idx + 1, len(rows)):
        row = rows[row_idx]
        if not row:
            continue
        day_cell = row[day_col_idx] if day_col_idx < len(row) else None
        if day_cell is None:
            continue
        try:
            day_num = int(day_cell)
        except (ValueError, TypeError):
            continue
        if not (1 <= day_num <= 31):
            continue

        for col_idx in sorted_month_cols:
            month_first = month_col_map[col_idx]
            # Validate date
            try:
                current_date = date(month_first.year, month_first.month, day_num)
            except ValueError:
                continue  # day doesn't exist in this month

            cell_val = row[col_idx] if col_idx < len(row) else None
            day_type = _classify_day(current_date, cell_val, holidays)
            calendar.append((current_date, day_type))

    return calendar


def _classify_day(
    d: date, cell_val, holidays: set
) -> Optional[str]:
    """
    Return day_type string or None for a teaching day.
    Priority: weekend > feiertag > cell marker
    """
    if d.weekday() >= 5:  # Saturday=5, Sunday=6
        return "WE"
    if d in holidays:
        return "Ft"
    if cell_val is None:
        return None
    if isinstance(cell_val, (int, float)):
        # Numeric → regular day (some calendars put day numbers here)
        return None
    if isinstance(cell_val, str):
        val = cell_val.strip()
        if val == "":
            return None
        # Normalize common markers
        mapping = {
            "fe": "Fe", "ferien": "Fe",
            "pr": "Pr", "praktikum": "Pr", "prakt.": "Pr",
            "prüf": "Prüf", "prüfung": "Prüf", "pruef": "Prüf",
            "ft": "Ft", "feiertag": "Ft",
        }
        return mapping.get(val.lower(), val)
    return None


def parse_zeitplan(
    path: str, bundesland: str = "HH"
) -> List[Tuple[date, Optional[str]]]:
    """
    Parse the Zeitplan Excel file.
    Returns an ordered list of (date, day_type) for all calendar days found.
    Teaching days have day_type == None.
    """
    print(f"[INFO] Parsing Zeitplan: {path}")
    wb = openpyxl.load_workbook(path, data_only=True)

    holidays = _detect_feiertage(wb, bundesland)
    print(f"  Found {len(holidays)} Feiertage for '{bundesland}'.")

    annual_sheets = _detect_year_sheets(wb)
    if not annual_sheets:
        print("  [WARN] No annual sheets detected. Trying all non-special sheets.")
        skip = {"schulfreie", "feiertag", "legende", "hinweise", "info"}
        annual_sheets = [
            s for s in wb.sheetnames
            if not any(kw in s.lower() for kw in skip)
        ]

    all_days: List[Tuple[date, Optional[str]]] = []
    seen_dates: set = set()

    for sheet_name in annual_sheets:
        print(f"  Processing sheet: '{sheet_name}'")
        ws = wb[sheet_name]
        sheet_days = _parse_annual_sheet(ws, holidays)
        added = 0
        for entry in sheet_days:
            if entry[0] not in seen_dates:
                all_days.append(entry)
                seen_dates.add(entry[0])
                added += 1
        print(f"    → {added} calendar days added.")

    all_days.sort(key=lambda x: x[0])
    teaching_count = sum(1 for _, dt in all_days if dt is None)
    print(f"  Total calendar days: {len(all_days)}, teaching days: {teaching_count}")
    return all_days


# ---------------------------------------------------------------------------
# Modulplan parsing
# ---------------------------------------------------------------------------

def parse_modulplan(path: str) -> List[Dict]:
    """
    Parse the Modulplan Excel.
    Returns list of dicts: {id, title, ue}
    """
    print(f"[INFO] Parsing Modulplan: {path}")
    df = pd.read_excel(path, header=None, dtype=str)

    # Find header row (first row containing at least 2 recognizable column names)
    header_row_idx = None
    col_map: Dict[str, int] = {}

    id_aliases = {"modul", "nr", "modul-nr", "modul nr", "modul_nr"}
    title_aliases = {"modulbezeichnung", "bezeichnung", "name", "titel", "modul-bezeichnung"}
    ue_aliases = {"ue", "unterrichtseinheiten", "stunden", "u.e.", "ue/std"}

    for row_idx, row in df.iterrows():
        row_lower = {_normalize_col_name(str(v)): ci for ci, v in enumerate(row) if pd.notna(v)}
        found_id = next((ci for k, ci in row_lower.items() if k in id_aliases), None)
        found_title = next((ci for k, ci in row_lower.items() if k in title_aliases), None)
        found_ue = next((ci for k, ci in row_lower.items() if k in ue_aliases), None)
        if found_id is not None and found_title is not None and found_ue is not None:
            header_row_idx = row_idx
            col_map = {"id": found_id, "title": found_title, "ue": found_ue}
            break
        elif found_title is not None and found_ue is not None:
            header_row_idx = row_idx
            col_map = {"id": found_id, "title": found_title, "ue": found_ue}
            break

    if header_row_idx is None:
        raise ValueError(
            "Modulplan: Could not find a header row with recognizable columns. "
            "Expected: Modul/Nr, Modulbezeichnung/Bezeichnung, UE/Stunden."
        )

    modules: List[Dict] = []
    for row_idx in range(header_row_idx + 1, len(df)):
        row = df.iloc[row_idx]

        def _get(ci):
            return row.iloc[ci] if ci is not None and ci < len(row) else None

        raw_id = _get(col_map.get("id"))
        raw_title = _get(col_map["title"])
        raw_ue = _get(col_map["ue"])

        # Skip empty rows
        if pd.isna(raw_title) or str(raw_title).strip() == "":
            continue
        if pd.isna(raw_ue) or str(raw_ue).strip() in ("", "nan"):
            continue

        try:
            ue_val = float(str(raw_ue).strip().replace(",", "."))
        except ValueError:
            continue

        if ue_val <= 0:
            continue

        mod_id = str(raw_id).strip() if raw_id is not None and not pd.isna(raw_id) else f"M{len(modules)}"

        modules.append({
            "id": mod_id,
            "title": str(raw_title).strip(),
            "ue": ue_val,
        })

    print(f"  Found {len(modules)} modules, total UE: {sum(m['ue'] for m in modules):.0f}")
    return modules


# ---------------------------------------------------------------------------
# Scheduling
# ---------------------------------------------------------------------------

def check_deficit(
    teaching_days: List[date], modules: List[Dict], ue_pro_tag: int
) -> int:
    """Return surplus (positive) or deficit (negative) in teaching days."""
    total_days_needed = sum(math.ceil(m["ue"] / ue_pro_tag) for m in modules)
    return len(teaching_days) - total_days_needed


def interactive_adjust(modules: List[Dict], deficit: int) -> List[Dict]:
    """
    Let the user interactively reduce UE of specific modules to cover the deficit.
    Returns (possibly modified) modules list.
    """
    days_to_free = abs(deficit)
    print()
    print("=" * 60)
    print(f"WARNUNG: Modulplan benötigt {days_to_free} Unterrichtstag(e) mehr als verfügbar.")
    print("=" * 60)
    print(f"\n{'Modul-ID':<12} {'Bezeichnung':<40} {'UE':>6} {'Tage':>6}")
    print("-" * 66)
    for m in modules:
        days = math.ceil(m["ue"] / 9)  # rough display
        print(f"{m['id']:<12} {m['title'][:40]:<40} {m['ue']:>6.0f} {days:>6}")
    print()
    print("Geben Sie Anpassungen ein (Modul-ID und neue UE, z.B. 'M3 72').")
    print("Drücken Sie Enter ohne Eingabe, um zu überspringen.\n")

    while days_to_free > 0:
        try:
            line = input(f"Anpassen ({days_to_free} Tag(e) noch offen) > ").strip()
        except EOFError:
            print("[INFO] Nicht-interaktiver Modus — überspringe Anpassung.")
            break
        if not line:
            break
        parts = line.split()
        if len(parts) < 2:
            print("  Format: <Modul-ID> <neue UE>  — z.B. 'M3 72'")
            continue
        mod_id, new_ue_str = parts[0], parts[1]
        try:
            new_ue = float(new_ue_str.replace(",", "."))
        except ValueError:
            print("  Ungültige UE-Zahl.")
            continue
        found = False
        for m in modules:
            if m["id"].lower() == mod_id.lower():
                old_days = math.ceil(m["ue"] / 9)
                m["ue"] = new_ue
                new_days = math.ceil(m["ue"] / 9)
                freed = old_days - new_days
                days_to_free -= freed
                print(f"  → {m['id']} angepasst: {m['ue']:.0f} UE ({freed:+d} Tage freigegeben). Noch offen: {days_to_free}")
                found = True
                break
        if not found:
            print(f"  Modul '{mod_id}' nicht gefunden.")

    return modules


def schedule_modules(
    calendar_days: List[Tuple[date, Optional[str]]],
    modules: List[Dict],
    ue_pro_tag: int,
) -> Tuple[List[Dict], Dict[date, str]]:
    """
    Assign teaching days to modules.

    Returns:
        schedule: list of dicts with id, title, ue, start_date, end_date,
                  days_assigned, ue_assigned
        day_to_module: dict mapping each calendar date to its label
                       (module ID, 'Fe', 'Pr', 'Prüf', 'Ft', 'WE', or '')
    """
    teaching_days = [d for d, dt in calendar_days if dt is None]
    day_to_type: Dict[date, Optional[str]] = {d: dt for d, dt in calendar_days}

    teaching_idx = 0
    schedule: List[Dict] = []
    day_to_module: Dict[date, str] = {}

    # Fill non-teaching days first
    for d, dt in calendar_days:
        if dt is not None and dt != "WE":
            day_to_module[d] = dt
        elif dt == "WE":
            day_to_module[d] = ""

    for module in modules:
        days_needed = math.ceil(module["ue"] / ue_pro_tag)
        assigned_teaching_days: List[date] = []

        for _ in range(days_needed):
            if teaching_idx >= len(teaching_days):
                break
            day = teaching_days[teaching_idx]
            assigned_teaching_days.append(day)
            day_to_module[day] = module["id"]
            teaching_idx += 1

        if not assigned_teaching_days:
            print(f"  [WARN] Module '{module['id']}' has no teaching days assigned (ran out).")
            schedule.append({
                "id": module["id"],
                "title": module["title"],
                "ue": module["ue"],
                "start_date": None,
                "end_date": None,
                "days_assigned": 0,
                "ue_assigned": 0,
            })
            continue

        start_date = assigned_teaching_days[0]
        end_date = assigned_teaching_days[-1]
        days_assigned = len(assigned_teaching_days)
        ue_assigned = min(module["ue"], days_assigned * ue_pro_tag)

        schedule.append({
            "id": module["id"],
            "title": module["title"],
            "ue": module["ue"],
            "start_date": start_date,
            "end_date": end_date,
            "days_assigned": days_assigned,
            "ue_assigned": ue_assigned,
        })

    return schedule, day_to_module


# ---------------------------------------------------------------------------
# Excel building helpers
# ---------------------------------------------------------------------------

def _apply_header_row(ws, row: int, headers: List[str], col_start: int = 1):
    """Write a header row with standard blue styling."""
    for ci, header in enumerate(headers, start=col_start):
        cell = ws.cell(row=row, column=ci, value=header)
        cell.font = _make_font(bold=True, color=HEADER_FG, size=10)
        cell.fill = _header_fill()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _thin_border()


def _write_s1_unterrichtsplan(
    ws, schedule: List[Dict], kursname: str, calendar_days: List[Tuple[date, Optional[str]]]
):
    """Build Sheet 1: Unterrichtsplan."""
    # Row 1: Title
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"Unterrichtsplan – {kursname}"
    title_cell.font = Font(name="Arial", bold=True, size=16, color=HEADER_FG)
    title_cell.fill = _header_fill()
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Row 2: Metadata strip
    from datetime import date as date_cls
    total_ue = sum(e["ue"] for e in schedule)
    total_days = sum(e["days_assigned"] for e in schedule)
    start_d = min((e["start_date"] for e in schedule if e["start_date"]), default=None)
    end_d = max((e["end_date"] for e in schedule if e["end_date"]), default=None)
    duration_str = ""
    if start_d and end_d:
        duration_str = f"{start_d.strftime('%d.%m.%Y')} – {end_d.strftime('%d.%m.%Y')}"

    meta_fill = PatternFill("solid", fgColor=META_BG.replace("#", ""))
    meta_font = _make_font(italic=True, size=10)
    meta_cells = [
        f"Erstellt: {date_cls.today().strftime('%d.%m.%Y')}",
        f"Gesamtdauer: {duration_str}",
        f"Gesamt-UE: {total_ue:.0f}",
        f"Gesamt-Tage: {total_days}",
        "", "", "",
    ]
    ws.merge_cells("A2:G2")
    ws["A2"].value = "  |  ".join(c for c in meta_cells if c)
    ws["A2"].font = meta_font
    ws["A2"].fill = meta_fill
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18

    # Row 3: Spacer
    ws.row_dimensions[3].height = 6

    # Row 4: Headers
    headers = ["Modul", "Modulbezeichnung", "Beginn", "Ende", "Tage", "UE", "Bemerkungen"]
    _apply_header_row(ws, row=4, headers=headers)
    ws.row_dimensions[4].height = 18

    # Rows 5+: Data
    sum_ue = 0
    sum_days = 0
    for row_offset, entry in enumerate(schedule):
        row = 5 + row_offset
        bg = "FFFFFF" if row_offset % 2 == 0 else ALT_ROW_BG
        fill = PatternFill("solid", fgColor=bg)
        font = _make_font(size=10)

        def _write(col, val, num_format=None, align="left"):
            c = ws.cell(row=row, column=col, value=val)
            c.font = font
            c.fill = fill
            c.alignment = Alignment(horizontal=align, vertical="center")
            c.border = _thin_border()
            if num_format:
                c.number_format = num_format
            return c

        _write(1, entry["id"], align="center")
        _write(2, entry["title"])
        if entry["start_date"]:
            _write(3, entry["start_date"], num_format="DD.MM.YYYY", align="center")
            _write(4, entry["end_date"], num_format="DD.MM.YYYY", align="center")
        else:
            _write(3, "—", align="center")
            _write(4, "—", align="center")
        _write(5, entry["days_assigned"], align="center")
        _write(6, int(entry["ue"]), align="center")
        _write(7, "")
        ws.row_dimensions[row].height = 15

        sum_ue += entry["ue"]
        sum_days += entry["days_assigned"]

    # SUMME row
    summe_row = 5 + len(schedule)
    summe_fill = PatternFill("solid", fgColor=SUMME_BG)
    summe_font = _make_font(bold=True, size=10)
    for col in range(1, 8):
        c = ws.cell(row=summe_row, column=col)
        c.fill = summe_fill
        c.font = summe_font
        c.border = _thin_border()
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(row=summe_row, column=1).value = "SUMME"
    ws.cell(row=summe_row, column=5).value = sum_days
    ws.cell(row=summe_row, column=6).value = int(sum_ue)
    ws.row_dimensions[summe_row].height = 16

    # Column widths
    for col_letter, width in COL_WIDTHS_S1.items():
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A5"


def _write_s2_kalendar_uebersicht(
    ws, calendar_days: List[Tuple[date, Optional[str]]], day_to_module: Dict[date, str], kursname: str
):
    """Build Sheet 2: Kalenderübersicht."""
    # Title
    ws.merge_cells("A1:E1")
    ws["A1"].value = f"Kalenderübersicht – {kursname}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color=HEADER_FG)
    ws["A1"].fill = _header_fill()
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # Headers
    headers = ["Jahr", "Monat", "Unterrichtstage", "UE", "Kumuliert UE"]
    _apply_header_row(ws, row=2, headers=headers)
    ws.row_dimensions[2].height = 16

    # Group teaching days by year/month
    from collections import defaultdict
    monthly: Dict[Tuple[int, int], int] = defaultdict(int)
    ue_per_day_approx = 9  # will be overridden if we can access ue_pro_tag

    for d, dt in calendar_days:
        if dt is None:
            monthly[(d.year, d.month)] += 1

    cumulative_ue = 0
    month_names_de = [
        "", "Januar", "Februar", "März", "April", "Mai", "Juni",
        "Juli", "August", "September", "Oktober", "November", "Dezember"
    ]
    row = 3
    for (year, month), count in sorted(monthly.items()):
        ue = count * ue_per_day_approx
        cumulative_ue += ue
        bg = "FFFFFF" if (row % 2 == 0) else ALT_ROW_BG
        fill = PatternFill("solid", fgColor=bg)
        font = _make_font(size=10)

        values = [year, month_names_de[month], count, ue, cumulative_ue]
        for ci, val in enumerate(values, start=1):
            c = ws.cell(row=row, column=ci, value=val)
            c.font = font
            c.fill = fill
            c.border = _thin_border()
            c.alignment = Alignment(horizontal="center" if ci != 2 else "left", vertical="center")
        ws.row_dimensions[row].height = 14
        row += 1

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 14
    ws.freeze_panes = "A3"


def _write_s3_jahreskalender(
    ws,
    calendar_days: List[Tuple[date, Optional[str]]],
    day_to_module: Dict[date, str],
    schedule: List[Dict],
    kursname: str,
):
    """Build Sheet 3: Jahreskalender with conditional formatting."""
    from collections import defaultdict

    # Organise calendar_days by (year, month) → {day: label}
    by_month: Dict[Tuple[int, int], Dict[int, str]] = defaultdict(dict)
    for d, dt in calendar_days:
        label = day_to_module.get(d, "")
        if label is None:
            label = ""
        by_month[(d.year, d.month)][d.day] = label

    # Sort months
    sorted_months: List[Tuple[int, int]] = sorted(by_month.keys())
    if not sorted_months:
        ws["A1"].value = "Keine Daten."
        return

    # Group months by year for header
    years: Dict[int, List[int]] = defaultdict(list)
    for year, month in sorted_months:
        years[year].append(month)

    # ---- Layout ----
    # Col A = day numbers, cols B onwards = months
    month_col_start = 2  # 1-indexed

    # Row 1: Title
    total_cols = 1 + len(sorted_months)
    last_col_letter = get_column_letter(total_cols)
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"].value = f"Jahreskalender – {kursname}"
    ws["A1"].font = Font(name="Arial", bold=True, size=14, color=HEADER_FG)
    ws["A1"].fill = _header_fill()
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Row 2: Year group headers (merged over their months)
    col_cursor = month_col_start
    for year, months_in_year in sorted(years.items()):
        span = len(months_in_year)
        start_letter = get_column_letter(col_cursor)
        end_letter = get_column_letter(col_cursor + span - 1)
        if span > 1:
            ws.merge_cells(f"{start_letter}2:{end_letter}2")
        c = ws.cell(row=2, column=col_cursor, value=str(year))
        c.font = _make_font(bold=True, size=11, color=HEADER_FG)
        c.fill = _header_fill()
        c.alignment = Alignment(horizontal="center", vertical="center")
        col_cursor += span
    ws["A2"].value = "Tag"
    ws["A2"].font = _make_font(bold=True, color=HEADER_FG)
    ws["A2"].fill = _header_fill()
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 18

    # Row 3: Month headers
    month_names_short = ["", "Jan", "Feb", "Mär", "Apr", "Mai", "Jun",
                         "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
    ws["A3"].value = ""
    ws["A3"].fill = _header_fill()
    ws["A3"].font = _make_font(bold=True, color=HEADER_FG)
    ws.row_dimensions[3].height = 16

    for col_idx, (year, month) in enumerate(sorted_months, start=month_col_start):
        label = f"{month_names_short[month]} {str(year)[2:]}"
        c = ws.cell(row=3, column=col_idx, value=label)
        c.font = _make_font(bold=True, color=HEADER_FG, size=9)
        c.fill = _header_fill()
        c.alignment = Alignment(horizontal="center", vertical="center")

    # Rows 4-34: Day rows
    for day_num in range(1, 32):
        row = 3 + day_num  # row 4=day1 ... row 34=day31
        ws.row_dimensions[row].height = 16

        # Col A: day number
        a_cell = ws.cell(row=row, column=1, value=day_num)
        a_cell.font = _make_font(bold=True, size=9)
        a_cell.fill = PatternFill("solid", fgColor="E8E8E8")
        a_cell.alignment = Alignment(horizontal="center", vertical="center")

        for col_idx, (year, month) in enumerate(sorted_months, start=month_col_start):
            cell_label = by_month[(year, month)].get(day_num, None)
            # Determine if this day is valid in this month
            try:
                d = date(year, month, day_num)
                if cell_label is None:
                    # check weekday
                    cell_label = "" if d.weekday() < 5 else ""
            except ValueError:
                cell_label = None  # invalid date for this month

            c = ws.cell(row=row, column=col_idx)
            c.value = cell_label if cell_label is not None else None
            c.font = _make_font(size=8)
            c.alignment = Alignment(horizontal="center", vertical="center")

    # Column widths
    ws.column_dimensions["A"].width = 5
    for col_idx in range(month_col_start, month_col_start + len(sorted_months)):
        ws.column_dimensions[get_column_letter(col_idx)].width = 7

    # Freeze panes
    ws.freeze_panes = "B4"

    # ---- Conditional formatting ----
    data_range = f"B4:{last_col_letter}34"
    for label, (bg_hex, fg_hex) in DAY_TYPE_COLORS.items():
        diff_style = DifferentialStyle(
            fill=PatternFill(bgColor=bg_hex),
            font=Font(color=fg_hex),
        )
        rule = Rule(
            type="containsText",
            operator="containsText",
            text=label,
            dxf=diff_style,
        )
        rule.formula = [f'NOT(ISERROR(SEARCH("{label}",B4)))']
        ws.conditional_formatting.add(data_range, rule)

    # ---- Legend (row 38+) ----
    legend_start_row = 37
    ws.cell(row=legend_start_row, column=1).value = "Legende"
    ws.cell(row=legend_start_row, column=1).font = _make_font(bold=True, size=10)

    # Build id→title lookup from schedule
    id_to_title = {e["id"]: e["title"] for e in schedule}
    # Add static entries
    static_legend = {
        "Fe": "Ferien",
        "Pr": "Praktikum",
        "Prüf": "Prüfung",
        "Ft": "Feiertag",
    }

    legend_row = legend_start_row + 1
    for label, (bg_hex, fg_hex) in DAY_TYPE_COLORS.items():
        if label.startswith("M") or label.startswith("IM"):
            title = id_to_title.get(label, "")
        else:
            title = static_legend.get(label, "")

        # Swatch
        swatch = ws.cell(row=legend_row, column=1, value=label)
        swatch.fill = PatternFill("solid", fgColor=bg_hex)
        swatch.font = Font(name="Arial", size=9, color=fg_hex, bold=True)
        swatch.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[legend_row].height = 14

        # Name
        name_cell = ws.cell(row=legend_row, column=2, value=title)
        name_cell.font = _make_font(size=9)
        name_cell.alignment = Alignment(horizontal="left", vertical="center")

        legend_row += 1


# ---------------------------------------------------------------------------
# Main build function
# ---------------------------------------------------------------------------

def build_excel(
    schedule: List[Dict],
    calendar_days: List[Tuple[date, Optional[str]]],
    day_to_module: Dict[date, str],
    kursname: str,
    output_path: str,
    ue_pro_tag: int,
):
    """Write the three-sheet output Excel file."""
    print(f"[INFO] Building Excel: {output_path}")
    wb = Workbook()

    # Sheet 1
    ws1 = wb.active
    ws1.title = "Unterrichtsplan"
    _write_s1_unterrichtsplan(ws1, schedule, kursname, calendar_days)

    # Sheet 2
    ws2 = wb.create_sheet("Kalenderübersicht")
    _write_s2_kalendar_uebersicht(ws2, calendar_days, day_to_module, kursname)

    # Sheet 3
    ws3 = wb.create_sheet("Jahreskalender")
    _write_s3_jahreskalender(ws3, calendar_days, day_to_module, schedule, kursname)

    wb.save(output_path)
    print(f"[INFO] Saved: {output_path}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generates a structured Unterrichtsplan Excel from Zeitplan and Modulplan.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--zeitplan", required=True, help="Path to Zeitplan Excel file")
    parser.add_argument("--modulplan", required=True, help="Path to Modulplan Excel file")
    parser.add_argument("--kursname", default="Kurs", help="Course name (default: Kurs)")
    parser.add_argument("--output", default=None, help="Output file path (auto-generated if omitted)")
    parser.add_argument("--ue-pro-tag", type=int, default=9, dest="ue_pro_tag",
                        help="UE per teaching day (default: 9)")
    parser.add_argument("--bundesland", default="HH",
                        help="Bundesland abbreviation for Feiertag detection (default: HH)")
    args = parser.parse_args()

    # Validate inputs
    if not os.path.isfile(args.zeitplan):
        print(f"[ERROR] Zeitplan not found: {args.zeitplan}", file=sys.stderr)
        sys.exit(1)
    if not os.path.isfile(args.modulplan):
        print(f"[ERROR] Modulplan not found: {args.modulplan}", file=sys.stderr)
        sys.exit(1)

    # Determine output path
    if args.output:
        output_path = args.output
    else:
        zeitplan_dir = os.path.dirname(os.path.abspath(args.zeitplan))
        # Try to detect start year
        try:
            wb_tmp = openpyxl.load_workbook(args.zeitplan, data_only=True, read_only=True)
            sheets = _detect_year_sheets(wb_tmp)
            wb_tmp.close()
            # If sheets are named like "2026", use that; else use current year
            start_year = sheets[0] if sheets and sheets[0].isdigit() else str(date.today().year)
        except Exception:
            start_year = str(date.today().year)
        filename = f"Unterrichtsplan_{args.kursname}_{start_year}.xlsx"
        output_path = os.path.join(zeitplan_dir, filename)

    # Parse inputs
    calendar_days = parse_zeitplan(args.zeitplan, args.bundesland)
    modules = parse_modulplan(args.modulplan)

    if not calendar_days:
        print("[ERROR] No calendar days parsed from Zeitplan.", file=sys.stderr)
        sys.exit(1)
    if not modules:
        print("[ERROR] No modules parsed from Modulplan.", file=sys.stderr)
        sys.exit(1)

    # Check deficit
    teaching_days = [d for d, dt in calendar_days if dt is None]
    deficit = check_deficit(teaching_days, modules, args.ue_pro_tag)

    if deficit < 0:
        is_interactive = sys.stdin.isatty()
        if is_interactive:
            modules = interactive_adjust(modules, deficit)
        else:
            total_needed = sum(math.ceil(m["ue"] / args.ue_pro_tag) for m in modules)
            print(
                f"\n[WARN] Deficit: {abs(deficit)} Unterrichtstag(e). "
                f"Needed: {total_needed}, Available: {len(teaching_days)}. "
                "Running in non-interactive mode — proceeding without adjustment.",
                file=sys.stderr,
            )
    else:
        print(f"[INFO] Capacity surplus: {deficit} teaching day(s) available after scheduling all modules.")

    # Schedule
    schedule, day_to_module = schedule_modules(calendar_days, modules, args.ue_pro_tag)

    # Build Excel
    build_excel(schedule, calendar_days, day_to_module, args.kursname, output_path, args.ue_pro_tag)

    print("\n[DONE]")
    print(f"  Output:        {output_path}")
    print(f"  Teaching days: {len(teaching_days)}")
    print(f"  Modules:       {len(schedule)}")
    print(f"  Total UE:      {sum(m['ue'] for m in modules):.0f}")


if __name__ == "__main__":
    main()
