from __future__ import annotations

import csv
import os
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
from PySide6.QtWidgets import QFileDialog


class ReformatSAPData:
    def __init__(self, ui):
        self.sap_files = []
        self.sap_file_names = []
        self.output_location = ""
        self.row_threshold = 15
        self.timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        self.output_name = "SAP_FD_"

        self.btn_select_files_to_format = ui.btn_format_files_to_reformat
        self.btn_select_output_location = ui.btn_format_output_location
        self.btn_run_process = ui.btn_format_run_process
        self.textEdit_selected_files = ui.textEdit_format_files_to_format
        self.textEdit_output_loc = ui.textEdit_format_output_location
        self.cb_remove_nonsmu = ui.checkBox_format_remove_non_SMU
        self.cb_remove_irrelevantsmu = ui.checkBox_format_irrelevant_SMU

    def get_output_location(self):
        path = QFileDialog.getExistingDirectory(parent=None, caption="", dir="", options=QFileDialog.Options())
        if path:
            self.output_location = path
            self.textEdit_output_loc.setPlainText(f"{path}")
        else:
            self.output_location = None

    def get_files_to_reformat(self):
        dialog = QFileDialog()
        dialog.setWindowTitle("Select Excel File(s)")
        dialog.setNameFilter("Excel Files (*.xlsx *.xlsm *.xltx *.xltm *.xls *.csv)")
        dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)

        if dialog.exec():
            file_paths = dialog.selectedFiles()
            self.sap_files = file_paths or []
            self.sap_file_names = [os.path.basename(excel_file) for excel_file in self.sap_files] or []

        if self.sap_file_names:
            self.textEdit_selected_files.setPlainText(",\n".join(self.sap_file_names))

    def remove_unneeded_columns_in_a_file(self, sap_files, output_dir):
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        for path in sap_files:
            path = Path(path)

            df = pd.read_csv(
                path,
                encoding="utf-16",
                sep=r"\s{2,}",
                engine="python",
                skiprows=3,
            )

            df = df.dropna(how="all").reset_index(drop=True)

            out_name = output_dir / f"{path.stem}_formatted.xlsx"
            df.to_excel(out_name, index=False)

    def csv_to_xlsx(self, sap_file, output_location):
        out_dir = Path(output_location)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_file = out_dir / "sap_output.xlsx"

        df = pd.read_csv(
            sap_file,
            encoding="utf-16",
            sep="\t",
            skiprows=3,
            engine="python",
            quotechar='"',
            quoting=csv.QUOTE_MINIMAL,
            on_bad_lines="skip",
        )

        df = df.iloc[:, 2:]
        df = df.dropna(how="all").reset_index(drop=True)

        df.columns = df.columns.str.strip()
        s = df["Sort String"].astype(str).str.strip()
        df = df[s.ne("") & s.ne("nan")].reset_index(drop=True)

        df = df[df["Sort String"].astype(str).str.strip().eq("SMU")].reset_index(drop=True)

        df.to_excel(out_file, index=False)

    def copy_csv(self, sap_file, output_location):
        src = Path(sap_file)
        if not src.exists() or src.suffix.lower() != ".csv":
            raise FileNotFoundError(f"Valid CSV not found: {src}")

        dst_dir = Path(output_location)
        dst_dir.mkdir(parents=True, exist_ok=True)

        new_name = "_copy"
        dst_name = new_name if new_name else src.name
        if not dst_name.lower().endswith(".csv"):
            dst_name += ".csv"

        dst = dst_dir / dst_name
        shutil.copy2(src, dst)
        return str(dst)

    def process_steps(self):
        self.csv_to_xlsx(self.sap_files[0], self.output_location)

    def run_process(self) -> None:
        self.btn_select_files_to_format.clicked.connect(self.get_files_to_reformat)
        self.btn_select_output_location.clicked.connect(self.get_output_location)
        self.btn_run_process.clicked.connect(self.process_steps)
        self.btn_run_process.clicked.connect(self.process_steps)


class SapTableReformatError(ValueError):
    pass


class SapTableReformatter:
    HEADER_COUNT = 6

    INDEX_HEAD_BOM_MAT = 0
    INDEX_BOM_COMPONENT = 1
    INDEX_SORT_STRING = 2
    INDEX_COMPONENT_DESC = 3
    INDEX_BASIC_NUMBER = 4
    INDEX_BASIC_NAME = 5

    CLEANUP_PREFIXES_1 = (
        "sal",
        "flex",
        "pall",
        "accl",
        "core",
        "film",
        "ship",
        "wgl",
        "sheet",
        "shee",
        "pl",
        "t-secur",
        "saco",
        "tear",
        "bulk",
    )

    CLEANUP_PREFIXES_2 = (
        "sal",
        "flex",
        "pall",
        "accl",
        "core",
        "film",
        "ship",
        "wgl",
        "sheet",
        "shee",
        "pl",
        "t-secur",
        "saco",
        "rbosac",
        "bag",
        "inst",
        "tear",
        "bulk",
    )

    def reformat_from_table(self, table, cleanup_mode: int) -> tuple[list[list[str]], int]:
        rows = self._extract_rows(table)
        self._validate_row_completeness(rows)

        cleaned_rows = self._remove_required_blank_rows(rows)
        aggregated_rows = self._aggregate_rows(cleaned_rows, cleanup_mode)

        if not aggregated_rows:
            raise SapTableReformatError("No rows remain after cleanup. Please review SAP table input.")

        return self._add_basic_combinations(aggregated_rows)

    def reformat_from_rows(self, rows: list[list[str]], cleanup_mode: int) -> tuple[list[list[str]], int]:
        """Same logic as reformat_from_table but accepts raw row lists directly."""
        self._validate_row_completeness(rows)

        cleaned_rows = self._remove_required_blank_rows(rows)
        aggregated_rows = self._aggregate_rows(cleaned_rows, cleanup_mode)

        if not aggregated_rows:
            raise SapTableReformatError("No rows remain after cleanup. Please review SAP data.")

        return self._add_basic_combinations(aggregated_rows)

    def _extract_rows(self, table) -> list[list[str]]:
        rows: list[list[str]] = []
        for row_idx in range(table.rowCount()):
            row_values: list[str] = []
            for col_idx in range(min(table.columnCount(), self.HEADER_COUNT)):
                item = table.item(row_idx, col_idx)
                row_values.append(item.text().strip() if item is not None else "")
            rows.append(row_values)
        return rows

    def _validate_row_completeness(self, rows: list[list[str]]) -> None:
        non_empty_rows = [row for row in rows if any(cell != "" for cell in row)]
        if not non_empty_rows:
            raise SapTableReformatError("SAP table has no input data.")

        complete_rows = [row for row in non_empty_rows if all(cell != "" for cell in row[: self.HEADER_COUNT])]
        ratio = len(complete_rows) / len(non_empty_rows)
        if ratio < 0.5:
            raise SapTableReformatError(
                "At least 50% of SAP table rows with data must have complete cell inputs before reformatting."
            )

    def _remove_required_blank_rows(self, rows: list[list[str]]) -> list[list[str]]:
        filtered: list[list[str]] = []
        for row in rows:
            if not any(cell != "" for cell in row):
                continue
            head = row[self.INDEX_HEAD_BOM_MAT]
            bom = row[self.INDEX_BOM_COMPONENT]
            sort_string = row[self.INDEX_SORT_STRING]
            if head == "" or bom == "" or sort_string == "":
                continue
            filtered.append(row)
        return filtered

    def _apply_cleanup_mode(self, rows: list[list[str]], cleanup_mode: int) -> list[list[str]]:
        if cleanup_mode == 3:
            return rows

        if cleanup_mode == 1:
            prefixes = self.CLEANUP_PREFIXES_1
        elif cleanup_mode == 2:
            prefixes = self.CLEANUP_PREFIXES_2
        else:
            raise SapTableReformatError("Unknown cleanup mode selected.")

        filtered: list[list[str]] = []
        for row in rows:
            basic_name = row[self.INDEX_BASIC_NAME]
            basic_name_text = basic_name.strip().lower()
            # Cleanup 2: remove rows whose basic name is purely numeric (un-identifiable)
            if cleanup_mode == 2 and basic_name.strip().isdigit():
                continue
            if "acco" in basic_name_text or "paco" in basic_name_text:
                continue
            if self._starts_with_any_prefix(basic_name, prefixes):
                continue

            filtered.append(row)

        return filtered

    def _starts_with_any_prefix(self, value: str, prefixes: tuple[str, ...]) -> bool:
        text = value.strip().lower()
        return any(text.startswith(prefix) for prefix in prefixes)

    def _aggregate_rows(self, rows: list[list[str]], cleanup_mode: int) -> list[list[str]]:
        grouped: dict[str, list[list[str]]] = {}
        for row in rows:
            head_bom_mat = row[self.INDEX_HEAD_BOM_MAT]
            grouped.setdefault(head_bom_mat, []).append(row)

        out_rows: list[list[str]] = []

        for head_bom_mat, head_rows in grouped.items():
            smu_rows = [row for row in head_rows if row[self.INDEX_SORT_STRING].strip().lower() == "smu"]

            if not smu_rows:
                out_rows.append(
                    [
                        head_bom_mat,
                        "NO",
                        "",
                        self._join_unique([row[self.INDEX_COMPONENT_DESC] for row in head_rows], separator=", "),
                        "",
                        "",
                        "comb 0",
                    ]
                )
                continue

            eligible_smu_rows = self._apply_cleanup_mode(smu_rows, cleanup_mode)
            if not eligible_smu_rows:
                # All SMU rows were in the b2r group.  As a fallback, retain the
                # first row whose Basic Name starts with "flex" (if one exists).
                flex_fallback = next(
                    (row for row in smu_rows if row[self.INDEX_BASIC_NAME].strip().lower().startswith("flex")),
                    None,
                )
                if flex_fallback is not None:
                    eligible_smu_rows = [flex_fallback]
                else:
                    out_rows.append(
                        [
                            head_bom_mat,
                            "NO",
                            "",
                            self._join_unique([row[self.INDEX_COMPONENT_DESC] for row in smu_rows], separator=", "),
                            "",
                            "",
                            "comb 0",
                        ]
                    )
                    continue

            eligible_with_basic_number = [
                row
                for row in eligible_smu_rows
                if self._strip_leading_zeros(row[self.INDEX_BASIC_NUMBER]) != ""
            ]

            if not eligible_with_basic_number:
                out_rows.append(
                    [
                        head_bom_mat,
                        "NO",
                        self._join_unique([row[self.INDEX_BOM_COMPONENT] for row in eligible_smu_rows], separator=", "),
                        self._join_unique([row[self.INDEX_COMPONENT_DESC] for row in eligible_smu_rows], separator=", "),
                        "",
                        "",
                        "comb 0",
                    ]
                )
                continue

            basic_number_text, basic_name_text = self._build_basic_number_and_name_strings(
                eligible_with_basic_number,
            )

            out_rows.append(
                [
                    head_bom_mat,
                    "YES",
                    self._join_unique([row[self.INDEX_BOM_COMPONENT] for row in eligible_with_basic_number], separator=", "),
                    self._join_unique([row[self.INDEX_COMPONENT_DESC] for row in eligible_with_basic_number], separator=", "),
                    basic_number_text,
                    basic_name_text,
                    "",
                ]
            )

        return out_rows

    def _build_basic_number_and_name_strings(
        self,
        rows: list[list[str]],
        name_source_rows: list[list[str]] | None = None,
    ) -> tuple[str, str]:
        ordered_basic_numbers: list[str] = []
        basic_number_to_name: dict[str, str] = {}

        source_rows = name_source_rows if name_source_rows is not None else rows

        for row in source_rows:
            basic_number = self._strip_leading_zeros(row[self.INDEX_BASIC_NUMBER])
            if basic_number == "":
                continue

            basic_name = (row[self.INDEX_BASIC_NAME] or "").strip()
            if basic_name == "":
                continue

            if basic_number not in basic_number_to_name:
                basic_number_to_name[basic_number] = basic_name

        for row in rows:
            basic_number = self._strip_leading_zeros(row[self.INDEX_BASIC_NUMBER])
            if basic_number == "":
                continue

            if basic_number not in basic_number_to_name:
                ordered_basic_numbers.append(basic_number)
                basic_number_to_name[basic_number] = ""
            elif basic_number not in ordered_basic_numbers:
                ordered_basic_numbers.append(basic_number)

        basic_number_text = ", ".join(ordered_basic_numbers)
        ordered_basic_names = [basic_number_to_name[number] for number in ordered_basic_numbers if basic_number_to_name[number] != ""]
        basic_name_text = " | ".join(ordered_basic_names)
        return basic_number_text, basic_name_text

    def _add_basic_combinations(self, rows: list[list[str]]) -> tuple[list[list[str]], int]:
        combination_map: dict[str, str] = {}  # basic_number_value -> final BC label
        bc_code_counts: dict[str, int] = {}   # base BC code -> how many times assigned
        out_rows: list[list[str]] = []

        for row in rows:
            basic_number_value = row[4].strip()
            basic_name_value = row[5].strip()
            preassigned_comb = row[6].strip()

            if preassigned_comb == "comb 0" or basic_number_value == "":
                combination_label = "comb 0"
            else:
                if basic_number_value not in combination_map:
                    base_code = self._generate_bc_code(
                        basic_number_value, basic_name_value
                    )
                    # Deduplicate: if the same base code was already used for
                    # a different basic_number_value, append _2, _3, etc.
                    occurrence = bc_code_counts.get(base_code, 0) + 1
                    bc_code_counts[base_code] = occurrence
                    if occurrence == 1:
                        combination_map[basic_number_value] = base_code
                    else:
                        combination_map[basic_number_value] = f"{base_code}_{occurrence}"
                combination_label = combination_map[basic_number_value]

            # Preserve Basic Name (row[5]) and add the combination label as the
            # Basic Combination column (after Basic Number and before Basic Name).
            out_rows.append(
                [
                    row[0],
                    row[1],
                    row[2],
                    row[3],
                    row[4],
                    combination_label,
                    row[5],
                ]
            )

        return out_rows, len(combination_map)

    def _generate_bc_code(self, basic_number_text: str, basic_name_text: str) -> str:
        """Build a descriptive BC code from Basic Number and Basic Name values.

        Format: ``{first_digit+last2digits}{sep}{name_abbrevs}``
        where *sep* is ``-`` for 1-2 components and ``_`` for 3+.
        """
        numbers = [n.strip() for n in basic_number_text.split(",") if n.strip()]
        names = [n.strip() for n in basic_name_text.split(" | ") if n.strip()]

        # First digit + last 2 digits of each Basic Number concatenated.
        number_parts = []
        for n in numbers:
            if len(n) >= 3:
                number_parts.append(n[0] + n[-2:])
            else:
                number_parts.append(n)
        number_str = "".join(number_parts)

        # Abbreviation for each Basic Name.
        name_abbrevs = [self._name_abbreviation(n) for n in names]
        # Pad with "na" when fewer names than numbers.
        while len(name_abbrevs) < len(numbers):
            name_abbrevs.append("na")

        separator = "_" if len(numbers) >= 3 else "-"
        return f"{number_str}{separator}{'-'.join(name_abbrevs)}"

    @staticmethod
    def _name_abbreviation(name: str) -> str:
        """Return a short abbreviation for a Basic Name value."""
        # Strip leading special characters / punctuation.
        cleaned = re.sub(r'^[^a-zA-Z0-9]+', '', name.strip())
        if not cleaned:
            return "na"
        # Purely numeric names → "na"
        if cleaned.replace(" ", "").isdigit():
            return "na"
        upper = cleaned.upper()
        if upper.startswith("CARTG"):
            return "cr"
        if upper.startswith("CAN"):
            return "cn"
        if upper.startswith("CAP"):
            return "cp"
        return cleaned[:2].lower()

    def _strip_leading_zeros(self, value: str) -> str:
        text = value.strip()
        if not text:
            return ""

        match = re.fullmatch(r"0*(\d+)", text)
        if match:
            digits = match.group(1)
            return digits if digits else "0"
        return text

    def _join_unique(self, values: list[str], separator: str) -> str:
        seen: set[str] = set()
        ordered: list[str] = []
        for raw in values:
            value = (raw or "").strip()
            if not value or value in seen:
                continue
            seen.add(value)
            ordered.append(value)
        return separator.join(ordered)


__all__ = ["ReformatSAPData", "SapTableReformatError", "SapTableReformatter"]
