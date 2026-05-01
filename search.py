"""
search.py — Search page functions for HAT Dashboard.

All functions used by the SEARCH page live here.
"""

from __future__ import annotations

import os
from pathlib import Path

from hat_config import HatConfig

# ---------------------------------------------------------------------------
# Source-file resolution
# ---------------------------------------------------------------------------

def _find_excel(folder: str, name_fragment: str | None = None) -> str | None:
    """Return the first .xlsx/.xls file in *folder* whose name contains
    *name_fragment* (case-insensitive).  If *name_fragment* is None, any
    Excel file is accepted.  Returns None when nothing is found."""
    folder_p = Path(folder)
    if not folder_p.is_dir():
        return None
    for f in sorted(folder_p.iterdir()):
        if f.suffix.lower() not in {".xlsx", ".xls"}:
            continue
        if name_fragment is None or name_fragment.lower() in f.name.lower():
            return str(f)
    return None


def resolve_source_files(source: str, cfg: HatConfig | None = None) -> dict[str, str | None]:
    """Return a mapping of source-key → absolute file path (or None if not found).

    Parameters
    ----------
    source:
        Value selected in the "Source" dropdown.  One of:
        ``"All"``, ``"Library"``, ``"TSC"``, ``"RSD (master)"``.
    cfg:
        A :class:`HatConfig` instance.  A fresh one is created when omitted.

    Returns
    -------
    dict with keys ``"library"``, ``"tsc"``, ``"rsd"``; each value is either
    an absolute path string or ``None`` when the file could not be located.
    """
    if cfg is None:
        cfg = HatConfig()

    result: dict[str, str | None] = {"library": None, "tsc": None, "rsd": None}

    use_library = source in {"All", "Library"}
    use_tsc     = source in {"All", "TSC"}
    use_rsd     = source in {"All", "RSD (master)"}

    if use_library:
        folder = cfg.search_library()
        result["library"] = _find_excel(folder) if folder else None

    if use_tsc:
        folder = cfg.search_tsc_data_folder()
        result["tsc"] = _find_excel(folder, name_fragment="tsc") if folder else None

    if use_rsd:
        folder = cfg.search_rsd_master_folder()
        result["rsd"] = _find_excel(folder, name_fragment="rsd") if folder else None

    return result


# ---------------------------------------------------------------------------
# Asset-type filtering
# ---------------------------------------------------------------------------

# Canonical build-type buckets for TSC rows
# Each set contains the canonical value plus recognised misspellings/variants.
# Matching is always case-insensitive.
_2D_VARIANTS: frozenset[str] = frozenset({
    # resizing family
    "resizing", "resize", "resized", "reszing", "resizng", "resizeing",
    # upload family
    "upload", "uplod", "uploaded", "uploded", "uploading", "uplaod", "uplaoded",
})

_3D_VARIANTS: frozenset[str] = frozenset({
    "clone", "master", "na",
})


def _classify_build_type(raw: str) -> str:
    """Return ``"2D"``, ``"3D"``, or ``"other"`` for a raw Build Type cell value."""
    normed = str(raw).strip().lower()
    if normed in _2D_VARIANTS:
        return "2D"
    if normed in _3D_VARIANTS:
        return "3D"
    return "other"


def _find_column(df_columns, target: str) -> str | None:
    """Case-insensitive column lookup; returns the actual column name or None."""
    target_lower = target.strip().lower()
    for col in df_columns:
        if str(col).strip().lower() == target_lower:
            return col
    return None


def apply_asset_type_filter(
    df,                  # pandas DataFrame
    asset_type: str,     # "All", "2D", or "3D"
    source_key: str,     # "tsc", "library", or "rsd"
):
    """Filter *df* rows according to *asset_type* and *source_key*.

    Rules
    -----
    * ``asset_type = "All"``  → no rows removed.
    * ``asset_type = "2D"``   → TSC only; keep rows whose Build Type resolves
                                to "2D"; rows with other values are dropped.
                                Library / RSD return an empty DataFrame for
                                this asset type (2D is TSC-exclusive).
    * ``asset_type = "3D"``   → TSC: keep rows whose Build Type resolves to
                                "3D".  Library: keep all rows (library is
                                entirely 3D).  RSD: keep all rows.
    """
    import pandas as pd  # local import — pandas may be heavy

    if asset_type == "All":
        return df

    if asset_type == "2D":
        if source_key != "tsc":
            # 2D is only applicable to TSC data
            return df.iloc[0:0].copy()  # empty frame with same columns
        build_col = _find_column(df.columns, "Build Type")
        if build_col is None:
            return df.iloc[0:0].copy()
        mask = df[build_col].fillna("").apply(
            lambda v: _classify_build_type(v) == "2D"
        )
        return df[mask].copy()

    if asset_type == "3D":
        if source_key in {"library", "rsd"}:
            # All library / RSD rows are treated as 3D
            return df
        # TSC: filter to clone / master rows
        build_col = _find_column(df.columns, "Build Type")
        if build_col is None:
            return df.iloc[0:0].copy()
        mask = df[build_col].fillna("").apply(
            lambda v: _classify_build_type(v) == "3D"
        )
        return df[mask].copy()

    # Unknown asset_type — return unchanged
    return df


# ---------------------------------------------------------------------------
# Column mappings  (field → {source_key: column_name | None})
# None means "do not search this source for this field"
# ---------------------------------------------------------------------------

FIELD_COLUMNS: dict[str, dict[str, str | None]] = {
    "idh": {
        "tsc":     "IDH Number",
        "rsd":     "Head Bom Mat",
        "library": "IDHs",
    },
    "pack_name": {
        "tsc":     "Product Name",
        "rsd":     "Component Desc",
        "library": "Sample Product Name",
    },
    "basic": {
        "tsc":     "Basic Number",
        "rsd":     "Basic Number",
        "library": "Basic Number",
    },
    "pack_type": {
        "tsc":     "Packaging Type",
        "rsd":     "Basic Name",
        "library": "IHA Container Configurator",
    },
    "pack_size": {
        "tsc":     "Packaging Size",
        "rsd":     "Basic Name",
        "library": "Capacity",
    },
    "label_size": {
        "tsc":     "Label Size",
        "rsd":     None,
        "library": None,
    },
    "project_name": {
        "tsc":     "Project Name",
        "rsd":     None,
        "library": "IHA Project Name",
    },
    "color": {
        "tsc":     None,
        "rsd":     None,
        "library": "Color/Material",
    },
    "sbu": {
        "tsc":     "SBU",
        "rsd":     None,
        "library": "SBU",
    },
    "build_type": {
        "tsc":     "Build Type",
        "rsd":     None,
        "library": None,
    },
    # "custom" searches all columns — handled separately
}


def build_search_filters(field_values: dict[str, str]) -> dict[str, str]:
    """Normalise raw field values from the UI into a clean filter dict.

    Parameters
    ----------
    field_values:
        Mapping of field key → raw string from the edit-text widget.
        Expected keys: ``"idh"``, ``"pack_name"``, ``"basic"``,
        ``"pack_type"``, ``"pack_size"``, ``"label_size"``,
        ``"project_name"``, ``"color"``, ``"sbu"``, ``"custom"``.

    Returns
    -------
    A dict containing only fields with non-empty values.
    If ``"custom"`` has a value, all other fields are omitted —
    the caller should search every column for the custom term.
    """
    custom = field_values.get("custom", "").strip()
    if custom:
        return {"custom": custom}

    filters: dict[str, str] = {}
    for key, raw in field_values.items():
        if key == "custom":
            continue
        val = raw.strip()
        if val:
            filters[key] = val
    return filters


def filter_df_by_fields(
    df,
    filters: dict[str, str],
    source_key: str,
) -> "pandas.DataFrame":
    """Apply *filters* to *df* and return matching rows.

    Parameters
    ----------
    df:
        A pandas DataFrame loaded from one source file.
    filters:
        Output of :func:`build_search_filters`.
    source_key:
        ``"tsc"``, ``"rsd"``, or ``"library"``.

    Notes
    -----
    * All comparisons are case-insensitive substring matches.
    * If the filter key is ``"custom"``, every column is searched and a row
      is kept if *any* cell contains the search term.
    * For regular fields, a row must satisfy *all* active filters
      (AND logic across fields).
    * If a field has no mapped column for this source (``None``), that
      filter is skipped for this source.
    """
    import pandas as pd

    if df.empty:
        return df

    if "custom" in filters:
        term = _norm(filters["custom"])
        # Vectorised column-wise OR — much faster than row-by-row apply()
        norm_df = df.astype(str).apply(lambda col: col.apply(_norm))
        mask = norm_df.apply(
            lambda col: col.str.contains(term, na=False, regex=False)
        ).any(axis=1)
        return df[mask].copy()

    mask = pd.Series([True] * len(df), index=df.index)
    for field_key, term in filters.items():
        col_name = FIELD_COLUMNS.get(field_key, {}).get(source_key)
        if col_name is None:
            # Field not applicable to this source — skip
            continue
        actual_col = _find_column(df.columns, col_name)
        if actual_col is None:
            # Column doesn't exist in this file — skip
            continue
        norm_term = _norm(term)
        norm_col  = df[actual_col].fillna("").astype(str).apply(_norm)
        # Try both orientations for dimension-like terms (302x600 ↔ 600x302)
        cell_match = pd.Series([False] * len(df), index=df.index)
        for variant in _dim_variants(term):
            cell_match |= norm_col.str.contains(variant, na=False, regex=False)
        mask &= cell_match

    return df[mask].copy()


# ---------------------------------------------------------------------------
# Display column definitions
# ---------------------------------------------------------------------------

DISPLAY_COLUMNS_MINIMAL: list[str] = [
    "IDH", "Pack Name", "Pack Type", "Pack Size", "Basic", "Project Name",
]

DISPLAY_COLUMNS_ALL: list[str] = [
    "IDH", "Pack Name", "Pack Type", "Pack Size", "Basic", "Project Name",
    "SBU", "Label Size", "Color",
]

# Maps display column name → field key (used for header colouring & extraction)
DISPLAY_COLUMN_FIELD: dict[str, str] = {
    "IDH":          "idh",
    "Pack Name":    "pack_name",
    "Pack Type":    "pack_type",
    "Pack Size":    "pack_size",
    "Basic":        "basic",
    "Project Name": "project_name",
    "SBU":          "sbu",
    "Label Size":   "label_size",
    "Color":        "color",
    "Build Type":   "build_type",
}

# ---------------------------------------------------------------------------
# Image matching
# ---------------------------------------------------------------------------

import re as _re
_IDH_PATTERN = _re.compile(r'(?<!\d)(\d{5,8})(?!\d)')


def _norm(s: str) -> str:
    """Lowercase and remove all whitespace — used for loose matching.

    Allows "302x600" to match "302 x 600" (and vice-versa) by collapsing
    spaces on both the search term and the cell value before comparing.
    """
    return _re.sub(r'\s+', '', s.lower())


_DIM_RE = _re.compile(r'^(\d+(?:\.\d+)?)x(\d+(?:\.\d+)?)(.*)$')


def _dim_variants(term: str) -> list[str]:
    """Return normalised search variants for *term*.

    If the term looks like a dimension (e.g. ``"302x600"`` after normalisation)
    both orientations are returned so that ``"302x600"`` matches ``"600x302"``
    and vice-versa.  Otherwise a single-element list is returned.
    """
    normed = _norm(term)
    m = _DIM_RE.match(normed)
    if not m:
        return [normed]
    a, b, suffix = m.group(1), m.group(2), m.group(3)
    if a == b:
        return [normed]
    return [f"{a}x{b}{suffix}", f"{b}x{a}{suffix}"]


def extract_idh_from_filename(name: str) -> str | None:
    """Return the first 5–8 digit number found in *name*, or None."""
    m = _IDH_PATTERN.search(name)
    return m.group(1) if m else None


def find_matching_images(
    idh_set: set[str],
    thumbnails_folder: str,
    fallback_folders: list[str] | None = None,
) -> list[dict]:
    """Scan *thumbnails_folder* (and any *fallback_folders*) recursively for
    images whose filename contains one of the IDH numbers in *idh_set*.

    For each IDH in *idh_set*, exactly one image is returned (the most-recently
    modified match). When no image is found for an IDH, a placeholder entry is
    returned with ``"path": None``.

    Returns a list of dicts:
    ``{"idh": str, "path": str | None, "name": str, "folder": str}``.
    """
    if not idh_set:
        return []

    search_roots: list[Path] = []
    for folder in ([thumbnails_folder] + (fallback_folders or [])):
        if folder:
            p = Path(folder)
            if p.is_dir():
                search_roots.append(p)

    IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}

    # idh → (best_path, best_mtime)
    best: dict[str, tuple[str, float]] = {}

    for root in search_roots:
        for f in root.rglob("*"):
            if f.suffix.lower() not in IMAGE_EXTS:
                continue
            # Check ALL digit sequences in the filename (not just the first).
            # Filenames often have multiple numeric parts; we want the one that
            # matches an IDH in idh_set rather than the first 5-8 digit run.
            idh: str | None = None
            for candidate in _IDH_PATTERN.findall(f.stem):
                if candidate in idh_set:
                    idh = candidate
                    break
            if idh and idh in idh_set:
                try:
                    mtime = f.stat().st_mtime
                except OSError:
                    mtime = 0.0
                if idh not in best or mtime > best[idh][1]:
                    best[idh] = (str(f), mtime)

    results: list[dict] = []
    for idh in sorted(idh_set):
        if idh in best:
            path = best[idh][0]
            f = Path(path)
            results.append({
                "idh":    idh,
                "path":   path,
                "name":   f.name,
                "folder": f.parent.name,
            })
        else:
            results.append({
                "idh":    idh,
                "path":   None,
                "name":   "",
                "folder": "",
            })
    return results


# ---------------------------------------------------------------------------
# Main search execution
# ---------------------------------------------------------------------------

def run_search(
    field_values: dict[str, str],
    multi_values: dict[str, list[str]],
    source: str,
    asset_type: str,
    status: str,
    sample_limit: int | None,
    cfg: HatConfig | None = None,
) -> dict:
    """Execute a search and return structured results.

    Parameters
    ----------
    field_values:
        Raw text from each search field.
        Keys: idh, pack_name, basic, pack_type, pack_size, label_size,
              project_name, color, sbu, custom.
    multi_values:
        Multi-value lists (OR logic within a field, AND across fields).
        Keys are field keys; these override field_values for the same key.
    source:
        "TSC", "RSD (master)", or "Library".
    asset_type:
        "All", "2D", or "3D".
    status:
        "Completed" or "All".
    sample_limit:
        Maximum rows to return.  None means no limit.
    cfg:
        HatConfig instance.

    Returns
    -------
    dict with keys:
        ``"rows"``           — list[dict] of normalised display-column rows
        ``"active_filters"`` — set[str] of field keys that had search values
        ``"source_key"``     — "tsc", "rsd", or "library"
        ``"idh_list"``       — list[str] of IDH values from matched rows
        ``"error"``          — str error message or None
    """
    import pandas as pd

    if cfg is None:
        cfg = HatConfig()

    source_map = {"TSC": "tsc", "RSD (master)": "rsd", "Library": "library"}
    source_key = source_map.get(source, "tsc")

    files = resolve_source_files(source, cfg)
    file_path = files.get(source_key)
    if not file_path:
        return {
            "rows": [], "active_filters": set(),
            "source_key": source_key, "idh_list": [],
            "error": f"Source file not found for '{source}'. Please check your settings.",
        }

    try:
        df = pd.read_excel(file_path, dtype=str)
    except Exception as exc:
        return {
            "rows": [], "active_filters": set(),
            "source_key": source_key, "idh_list": [],
            "error": f"Failed to load file:\n{exc}",
        }

    # Asset-type filter
    df = apply_asset_type_filter(df, asset_type, source_key)

    # Status filter (TSC only)
    if status != "All" and source_key == "tsc":
        status_col = _find_column(df.columns, "Status")
        if status_col is not None:
            df = df[df[status_col].fillna("").str.lower() == status.lower()].copy()

    # Build combined filter dict (multi_values override single field_values)
    base = build_search_filters(field_values)
    list_filters: dict[str, list[str]] = {k: [v] for k, v in base.items()}
    for k, vals in multi_values.items():
        cleaned = [v.strip() for v in vals if v.strip()]
        if cleaned:
            list_filters[k] = cleaned
    active_filters: set[str] = set(list_filters.keys())

    # Apply filters
    if df.empty:
        matched = df
    elif "custom" in list_filters:
        terms = list_filters["custom"]
        # Vectorised column-wise OR for each term (AND across terms)
        norm_df = df.astype(str).apply(lambda col: col.apply(_norm))
        mask = pd.Series([True] * len(df), index=df.index)
        for term in terms:
            t = _norm(term)
            mask &= norm_df.apply(
                lambda col, _t=t: col.str.contains(_t, na=False, regex=False)
            ).any(axis=1)
        matched = df[mask].copy()
    else:
        mask = pd.Series([True] * len(df), index=df.index)
        for field_key, terms in list_filters.items():
            col_name = FIELD_COLUMNS.get(field_key, {}).get(source_key)
            if col_name is None:
                continue
            actual_col = _find_column(df.columns, col_name)
            if actual_col is None:
                continue
            # OR within field: row matches if cell contains ANY of the terms.
            # Whitespace is collapsed; dimension terms also try both orientations
            # so "302x600" matches "600x302" (and vice-versa).
            norm_col = df[actual_col].fillna("").astype(str).apply(_norm)
            field_mask = pd.Series([False] * len(df), index=df.index)
            for term in terms:
                for variant in _dim_variants(term):
                    field_mask |= norm_col.str.contains(
                        variant, na=False, regex=False
                    )
            mask &= field_mask
        matched = df[mask].copy()

    # Apply sample limit
    if sample_limit is not None:
        matched = matched.head(sample_limit)

    # Normalise to display rows
    def _get_val(row: "pd.Series", field_key: str) -> str:
        col_name = FIELD_COLUMNS.get(field_key, {}).get(source_key)
        if not col_name:
            return ""
        actual_col = _find_column(row.index, col_name)
        if actual_col is None:
            return ""
        val = row.get(actual_col, "")
        return "" if pd.isna(val) else str(val)

    rows: list[dict[str, str]] = []
    for _, row in matched.iterrows():
        record = {
            display_col: _get_val(row, fk)
            for display_col, fk in DISPLAY_COLUMN_FIELD.items()
        }
        # Resolve Build Type based on source
        if source_key == "library":
            record["Build Type"] = "3D"
        elif source_key == "rsd":
            record["Build Type"] = "NA"
        # TSC: value already extracted via FIELD_COLUMNS above
        rows.append(record)

    # Collect IDH values for image matching
    idh_col_name = FIELD_COLUMNS.get("idh", {}).get(source_key)
    idh_list: list[str] = []
    if idh_col_name:
        actual_idh_col = _find_column(matched.columns, idh_col_name)
        if actual_idh_col:
            idh_list = [
                v for v in
                matched[actual_idh_col].fillna("").astype(str).str.strip().tolist()
                if v
            ]

    return {
        "rows": rows,
        "active_filters": active_filters,
        "source_key": source_key,
        "idh_list": idh_list,
        "error": None,
    }
