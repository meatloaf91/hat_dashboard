r"""
hat_config.py — Configuration helper for HAT Dashboard.

Config file location: %APPDATA%\HAT Dashboard\hat_dashboard_config.ini
(Created automatically on first save.)

INI structure
─────────────
[general]
root_folder = <path>         ; root location for MASTER hierarchy

[tsc]
input_folder  =          ; empty → <root>/HAT DASHBOARD ROOT/MASTER/Tracker Status Collector
output_folder =

[thumbnail]
input_folder  =          ; empty → <root>/HAT DASHBOARD ROOT/MASTER/Thumbnail Generator
output_folder =

[sap_reformat]
input_folder  =          ; empty → <root>/HAT DASHBOARD ROOT/MASTER/SAP Data Reformat
output_folder =

Derivation rules (when a per-tool key is empty):
  tsc          input/output → <root>/HAT DASHBOARD ROOT/MASTER/Tracker Status Collector
  thumbnail    input/output → <root>/HAT DASHBOARD ROOT/MASTER/Thumbnail Generator
  sap_reformat input/output → <root>/HAT DASHBOARD ROOT/MASTER/SAP Data Reformat
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from configparser import ConfigParser

# ---------------------------------------------------------------------------
# Folder hierarchy created under the chosen root (Setup Root mode)
# ---------------------------------------------------------------------------
FOLDER_HIERARCHY: list[str] = [
    "HAT DASHBOARD ROOT/MASTER/Tracker Status Collector",
    "HAT DASHBOARD ROOT/MASTER/Thumbnail Generator",
    "HAT DASHBOARD ROOT/MASTER/SAP Data Reformat",
    "HAT DASHBOARD ROOT/MASTER/Excel Library",
]

# Subfolders created inside a new project folder (Setup a Project mode)
PROJECT_SUBFOLDERS: list[str] = [
    "Packshot Naming Generator",
    "SAP Data Reformat",
    "SAP Data Compare",
    "Project Viewer",
]

_DEFAULT_SECTIONS: dict[str, dict[str, str]] = {
    "general":      {"root_folder": ""},
    "tsc":          {"input_folder": "", "output_folder": ""},
    "thumbnail":    {"input_folder": "", "output_folder": ""},
    "sap_reformat": {"input_folder": "", "output_folder": ""},
    "excel_library":{"folder": ""},
    "sap_compare":  {"rsd_target_folder": "", "rsd_master_folder": "", "tsc_data_folder": ""},
    "search":       {"rsd_master_folder": "", "tsc_data_folder": "", "library": "", "thumbnails_folder": ""},
}


def _exe_dir() -> Path:
    """Return the directory next to the running script / packaged exe."""
    if getattr(sys, "frozen", False):       # PyInstaller bundle
        return Path(sys.executable).parent
    return Path(__file__).parent


def _config_path() -> Path:
    appdata = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    return appdata / "HAT Dashboard" / "hat_dashboard_config.ini"


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

class HatConfig:
    """Thin wrapper around ConfigParser for HAT Dashboard settings."""

    def __init__(self) -> None:
        self._path = _config_path()
        self._cfg = ConfigParser()
        self._load()

    # ── persistence ────────────────────────────────────────────────────────

    def _load(self) -> None:
        # seed defaults so every section/key always exists
        for section, keys in _DEFAULT_SECTIONS.items():
            if not self._cfg.has_section(section):
                self._cfg.add_section(section)
            for key, value in keys.items():
                if not self._cfg.has_option(section, key):
                    self._cfg.set(section, key, value)
        # overwrite with whatever is on disk (if present)
        if self._path.exists():
            self._cfg.read(self._path, encoding="utf-8")

    def save(self) -> None:
        # Remove any sections/keys not in the current schema before writing
        for section in self._cfg.sections():
            if section not in _DEFAULT_SECTIONS:
                self._cfg.remove_section(section)
            else:
                for key in self._cfg.options(section):
                    if key not in _DEFAULT_SECTIONS[section]:
                        self._cfg.remove_option(section, key)
        self._path.parent.mkdir(parents=True, exist_ok=True)
        with open(self._path, "w", encoding="utf-8") as fh:
            self._cfg.write(fh)

    def config_path(self) -> Path:
        return self._path

    # ── raw get / set ───────────────────────────────────────────────────────

    def get(self, section: str, key: str, fallback: str = "") -> str:
        return self._cfg.get(section, key, fallback=fallback).strip()

    def set(self, section: str, key: str, value: str) -> None:
        if not self._cfg.has_section(section):
            self._cfg.add_section(section)
        self._cfg.set(section, key, value.strip())

    # ── root folder helpers ─────────────────────────────────────────────────

    def root_folder(self) -> str:
        return self.get("general", "root_folder")

    def set_root_folder(self, path: str) -> None:
        self.set("general", "root_folder", path)

    def create_folder_hierarchy(self, root: str) -> list[str]:
        """Create MASTER subfolders under *root*/HAT DASHBOARD ROOT. Returns list of created paths."""
        created: list[str] = []
        root_p = Path(root)
        for rel in FOLDER_HIERARCHY:
            p = root_p / rel
            if not p.exists():
                p.mkdir(parents=True, exist_ok=True)
                created.append(str(p))
        return created

    def create_project_folder_hierarchy(self, location: str, project_name: str) -> tuple[list[str], str]:
        """Create a project folder and its subfolders. Returns (created_paths, project_path)."""
        created: list[str] = []
        project_p = Path(location) / project_name
        for sub in PROJECT_SUBFOLDERS:
            p = project_p / sub
            if not p.exists():
                p.mkdir(parents=True, exist_ok=True)
                created.append(str(p))
        return created, str(project_p)

    def project_parent_folder(self) -> str:
        return self.get("general", "project_parent_folder", fallback="")

    # ── resolved paths (per-tool override → root derivation) ───────────────

    def _resolve(self, section: str, key: str, root_rel: str) -> str:
        """Return the stored override, or derive from root, or ''."""
        override = self.get(section, key)
        if override:
            return override
        root = self.root_folder()
        if root:
            return str(Path(root) / root_rel)
        return ""

    def tsc_input(self) -> str:
        return self._resolve("tsc", "input_folder",
                             "HAT DASHBOARD ROOT/MASTER/Tracker Status Collector")

    def tsc_output(self) -> str:
        return self._resolve("tsc", "output_folder",
                             "HAT DASHBOARD ROOT/MASTER/Tracker Status Collector")

    def thumbnail_input(self) -> str:
        return self._resolve("thumbnail", "input_folder",
                             "HAT DASHBOARD ROOT/MASTER/Thumbnail Generator")

    def thumbnail_output(self) -> str:
        return self._resolve("thumbnail", "output_folder",
                             "HAT DASHBOARD ROOT/MASTER/Thumbnail Generator")

    def sap_reformat_input(self) -> str:
        return self._resolve("sap_reformat", "input_folder",
                             "HAT DASHBOARD ROOT/MASTER/SAP Data Reformat")

    def sap_reformat_output(self) -> str:
        return self._resolve("sap_reformat", "output_folder",
                             "HAT DASHBOARD ROOT/MASTER/SAP Data Reformat")

    def excel_library_folder(self) -> str:
        return self._resolve("excel_library", "folder",
                             "HAT DASHBOARD ROOT/MASTER/Excel Library")

    def rsd_master_folder(self) -> str:
        """Return folder to scan for the RSD master file.
        Priority: [sap_compare] rsd_master_folder → sap_reformat_input()."""
        override = self.get("sap_compare", "rsd_master_folder")
        if override:
            return override
        return self.sap_reformat_input()

    # ── search page folder helpers ──────────────────────────────────────────

    def search_rsd_master_folder(self) -> str:
        return self._resolve("search", "rsd_master_folder",
                             "HAT DASHBOARD ROOT/MASTER/SAP Data Reformat")

    def search_tsc_data_folder(self) -> str:
        return self._resolve("search", "tsc_data_folder",
                             "HAT DASHBOARD ROOT/MASTER/Tracker Status Collector")

    def search_library(self) -> str:
        return self._resolve("search", "library",
                             "HAT DASHBOARD ROOT/MASTER/Excel Library")

    def search_thumbnails_folder(self) -> str:
        return self._resolve("search", "thumbnails_folder",
                             "HAT DASHBOARD ROOT/MASTER/Thumbnail Generator")

    # ── validation ──────────────────────────────────────────────────────────

    def is_root_configured(self) -> bool:
        """True when root_folder is set AND that directory exists on disk."""
        root = self.root_folder()
        return bool(root) and Path(root).is_dir()

    def validate_tool_folders(self, tool: str) -> bool:
        """
        Return True if every folder the *tool* needs actually exists on disk.
        tool ∈ {'tsc', 'thumbnail', 'sap_reformat'}
        """
        checks: dict[str, list[str]] = {
            "tsc":         [self.tsc_input(), self.tsc_output()],
            "thumbnail":   [self.thumbnail_input(), self.thumbnail_output()],
            "sap_reformat":[self.sap_reformat_input(), self.sap_reformat_output()],
        }
        paths = checks.get(tool, [])
        return bool(paths) and all(Path(p).is_dir() for p in paths if p)
