"""
review_project.py — Business logic for the Review Project module.

All data-processing and file-scanning functions used by the Review Project
feature live here.  The UI layer (new_ui_window.py) imports these functions
and delegates to them; no PySide6 / UI code belongs in this module.
"""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Sheet name constant (shared by load_body_map_sheet and the UI dialog)
# ---------------------------------------------------------------------------

BODY_MAP_SHEET = "body_map_IDH"


# ---------------------------------------------------------------------------
# File-scanning helpers
# ---------------------------------------------------------------------------

def pick_latest_review_file(files: list) -> "Path | None":
    """Return the Path with the latest date encoded in its stem.

    Recognises two date patterns:
      - ``YYYYMMDD``  (e.g. ``review_20250101.xlsx``)
      - ``YYYY-MM-DD`` / ``YYYY_MM_DD``

    Falls back to the lexicographically last entry when no date is found.
    """
    def _extract_date(p: Path) -> datetime:
        stem = p.stem
        m = re.search(r'(\d{4})(\d{2})(\d{2})', stem)
        if m:
            try:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
        m = re.search(r'(\d{4})[-_](\d{2})[-_](\d{2})', stem)
        if m:
            try:
                return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                pass
        return datetime.min

    return max(files, key=_extract_date, default=None)


def scan_sdc_files(project_folder: str) -> list[Path]:
    """Return sorted list of Excel files inside *project_folder*/SAP Data Compare.

    Used by the New Review mode to autofill the Source BMA field.
    """
    sdc_folder = Path(project_folder) / "SAP Data Compare"
    if not sdc_folder.is_dir():
        return []
    return sorted(
        f for f in sdc_folder.iterdir()
        if f.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
    )


def scan_review_files(project_folder: str) -> list[Path]:
    """Return sorted list of review Excel files inside *project_folder*/Project Review.

    A file qualifies when its stem contains the word ``review`` (case-insensitive).
    Returns an empty list when the sub-folder does not exist.
    """
    pr_folder = Path(project_folder) / "Project Review"
    if not pr_folder.is_dir():
        return []
    return sorted(
        f for f in pr_folder.iterdir()
        if f.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
        and "review" in f.stem.lower()
    )


def scan_tracker_files(project_folder: str) -> list[Path]:
    """Return sorted list of Briefing Tracker Excel files inside *project_folder*/Project Review.

    A file qualifies when its stem matches the pattern ``briefing?tracker``
    (case-insensitive, ``?`` = optional separator character).
    Returns an empty list when the sub-folder does not exist.
    """
    pr_folder = Path(project_folder) / "Project Review"
    if not pr_folder.is_dir():
        return []
    return sorted(
        f for f in pr_folder.iterdir()
        if f.suffix.lower() in {".xlsx", ".xls", ".xlsm"}
        and re.search(r'briefing.?tracker', f.stem, re.IGNORECASE)
    )


# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_body_map_sheet(file_path: str) -> tuple[list[str], list[list[str]]]:
    """Read the ``body_map_IDH`` sheet from *file_path*.

    Returns a ``(headers, data_rows)`` tuple where every value is a plain
    string.  Returns ``([], [])`` on any error so the caller can handle the
    empty-table case gracefully.
    """
    try:
        import pandas as pd
        df = pd.read_excel(
            file_path,
            sheet_name=BODY_MAP_SHEET,
            engine="calamine",
            dtype=str,
        )
        df = df.fillna("")
        return list(df.columns), [list(r) for r in df.values.tolist()]
    except Exception:
        return [], []


# ---------------------------------------------------------------------------
# Missing IDH detection
# ---------------------------------------------------------------------------

def find_missing_idhs(tracker_file: str, existing_hbm: "set[str]") -> "list[str]":
    """Return IDH values from *tracker_file* not yet present in *existing_hbm*.

    Searches sheets whose name contains 'tracker' (case-insensitive) for a
    column named 'IDH number', 'IDH num', or 'IDH' (case-insensitive), scanning
    the first 20 rows to locate the header row.  Returns a deduplicated,
    order-preserving list of normalised values that are missing from
    *existing_hbm*.
    """
    IDH_ALIASES: set[str] = {"idh number", "idh num", "idh"}

    def _norm(v: str) -> str:
        """Strip whitespace; also strip trailing .0 from numeric-looking values."""
        v = str(v).strip()
        if v.endswith(".0") and v[:-2].isdigit():
            v = v[:-2]
        return v

    existing_norm: set[str] = {_norm(v) for v in existing_hbm if str(v).strip()}

    try:
        import pandas as pd
        xl = pd.ExcelFile(tracker_file, engine="calamine")
        sheets_to_check = [s for s in xl.sheet_names if "tracker" in s.lower()]
        # Fall back to all sheets if none have "tracker" in the name
        if not sheets_to_check:
            sheets_to_check = list(xl.sheet_names)

        for sheet in sheets_to_check:
            # Scan the first 20 rows to find the header row containing an IDH alias
            raw = xl.parse(sheet, header=None, dtype=str).fillna("")
            header_row: "int | None" = None
            for row_idx in range(min(20, len(raw))):
                row_vals = [str(v).strip().lower() for v in raw.iloc[row_idx]]
                if any(v in IDH_ALIASES for v in row_vals):
                    header_row = row_idx
                    break
            if header_row is None:
                continue

            df = xl.parse(sheet, header=header_row, dtype=str).fillna("")
            idh_col = next(
                (c for c in df.columns if str(c).strip().lower() in IDH_ALIASES),
                None,
            )
            if idh_col is None:
                continue

            seen: set[str] = set()
            result: list[str] = []
            for raw_val in df[idh_col]:
                norm = _norm(str(raw_val))
                if not norm or norm in existing_norm or norm in seen:
                    continue
                seen.add(norm)
                result.append(norm)
            return result

    except Exception:
        return []

    return []
