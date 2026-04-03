import os

from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QMessageBox,
)

import stylesheet as ss
from general_functions import THUMBNAIL_OUTPUT_HINT, clear_thumbnail_inputs

# Subfolder names that indicate a potential final-output project structure (case-insensitive).
_PROTECTED_SUBFOLDER_NAMES: frozenset[str] = frozenset({"completed", "source", "draft"})
# Subfolder names whose images should be silently excluded (case-insensitive).
_HIRES_SUBFOLDER_NAMES: frozenset[str] = frozenset({"hires", "highres", "high resolution", "high res"})

try:
    from PIL import Image
except Exception:  # pragma: no cover - dependency check at runtime
    Image = None


class ThumbnailGenerator:
    def __init__(self, ui):
        self.btn_images_folder = ui.btn_pg7_images_folder
        self.btn_output_location = ui.btn_pg7_output
        self.btn_run_process = ui.btn_run_process_generate_thumbnails
        self.input_images_folder = ui.input_pg7_images_folder
        self.input_output_location = ui.input_pg7_output
        self.progress_bar = getattr(ui, "thumbnail_progress_bar", None)

        self.images_folder = ""
        self.output_location = ""
        self.selected_folders: list[str] = []
        self.supported_formats = (".png", ".tif", ".tiff", ".jpg", ".jpeg")

        # Output is derived from source folder names by design.
        self.btn_output_location.setEnabled(False)
        self.input_output_location.setReadOnly(True)
        self.input_output_location.setText(THUMBNAIL_OUTPUT_HINT)

    def run_process(self) -> None:
        self.btn_images_folder.clicked.connect(self.browse_images_folder)
        self.btn_run_process.clicked.connect(self.process_images)

    def browse_images_folder(self, start_dir: str = "") -> None:
        default = start_dir if start_dir else (self.selected_folders[-1] if self.selected_folders else "")
        path = QFileDialog.getExistingDirectory(
            self.btn_run_process,
            "Select Images Folder",
            default,
            QFileDialog.Option.ShowDirsOnly,
        )
        if not path:
            return

        normalized = os.path.normpath(path)
        if os.path.isdir(normalized) and normalized not in self.selected_folders:
            self.selected_folders.append(normalized)

        self.images_folder = "; ".join(self.selected_folders)
        self.input_images_folder.setText(self.images_folder)

    def _sync_inputs(self) -> None:
        typed = self.input_images_folder.text().strip()
        if typed:
            parsed = [os.path.normpath(p.strip()) for p in typed.split(";") if p.strip()]
            if parsed:
                self.selected_folders = list(dict.fromkeys(parsed))
        self.images_folder = "; ".join(self.selected_folders)

    def _get_source_folders(self) -> list[str]:
        self._sync_inputs()
        valid = [folder for folder in self.selected_folders if os.path.isdir(folder)]
        return list(dict.fromkeys(valid))

    def _has_subfolders(self, folder: str) -> bool:
        try:
            for entry in os.scandir(folder):
                if entry.is_dir():
                    return True
        except Exception:
            return False
        return False

    def _folder_has_direct_images(self, folder: str) -> bool:
        try:
            for name in os.listdir(folder):
                if name.lower().endswith(self.supported_formats):
                    return True
        except Exception:
            return False
        return False

    def _scan_subfolder_rules(self, folder: str) -> tuple[list[str], list[str]]:
        """Scan direct subfolders of *folder* and return (protected_names, hires_names).

        protected_names — subfolders whose names match Source/Draft/Completed;
                          these trigger a warning and are excluded from processing.
        hires_names     — subfolders whose names indicate hi-res sources;
                          these are silently excluded from processing.
        """
        protected: list[str] = []
        hires: list[str] = []
        try:
            for entry in os.scandir(folder):
                if not entry.is_dir():
                    continue
                name_lower = entry.name.lower()
                if name_lower in _PROTECTED_SUBFOLDER_NAMES:
                    protected.append(entry.name)
                elif name_lower in _HIRES_SUBFOLDER_NAMES:
                    hires.append(entry.name)
        except Exception:
            pass
        return protected, hires

    def _iter_image_files(self, folder: str, recursive: bool, excluded_dirs: frozenset[str] = frozenset()):
        if recursive:
            for root, dirs, files in os.walk(folder):
                # Prune excluded directories in-place so os.walk skips them entirely.
                dirs[:] = [d for d in dirs if d.lower() not in excluded_dirs]
                for name in files:
                    if name.lower().endswith(self.supported_formats):
                        yield os.path.join(root, name), name
            return

        for name in os.listdir(folder):
            if name.lower().endswith(self.supported_formats):
                yield os.path.join(folder, name), name

    def _build_jobs(self, folders: list[str], excluded_dirs: frozenset[str] = frozenset()) -> list[dict]:
        jobs: list[dict] = []
        if len(folders) == 1:
            folder = folders[0]
            # Parent mode: one selected folder with children and no direct images.
            if self._has_subfolders(folder) and not self._folder_has_direct_images(folder):
                jobs.append(
                    {
                        "source": folder,
                        "output": folder,  # overwrite in-place
                        "recursive": True,
                        "excluded_dirs": excluded_dirs,
                    }
                )
                return jobs

        # Single regular folder OR multiple selected folders.
        for folder in folders:
            jobs.append(
                {
                    "source": folder,
                    "output": folder,  # overwrite in-place
                    "recursive": False,
                    "excluded_dirs": frozenset(),
                }
            )
        return jobs

    def _unique_output_path(self, folder: str, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        target = os.path.join(folder, filename)
        if not os.path.exists(target):
            return target

        counter = 2
        while True:
            alt_name = f"{base}_{counter}{ext}"
            alt_path = os.path.join(folder, alt_name)
            if not os.path.exists(alt_path):
                return alt_path
            counter += 1

    def _set_processing_state(self, running: bool, total: int = 0) -> None:
        self.btn_run_process.setEnabled(not running)
        if self.progress_bar is not None:
            self.progress_bar.setVisible(running)
            if running:
                if total > 0:
                    self.progress_bar.setRange(0, total)
                    self.progress_bar.setValue(0)
                else:
                    self.progress_bar.setRange(0, 0)
            else:
                self.progress_bar.setRange(0, 100)
                self.progress_bar.setValue(0)

        app = QApplication.instance()
        if app is not None:
            app.processEvents()

    def _show_message(self, title: str, text: str, icon: QMessageBox.Icon) -> None:
        msg = QMessageBox(self.btn_run_process)
        msg.setWindowTitle(title)
        msg.setIcon(icon)
        msg.setText(text)
        msg.setStyleSheet(ss.msg_stylesheet)
        msg.exec()

    def clear_all_fields(self) -> None:
        clear_thumbnail_inputs(self)

    def process_images(self) -> None:
        if Image is None:
            self._show_message(
                "Error",
                "Pillow (PIL) is not installed. Please install it to generate thumbnails.",
                QMessageBox.Icon.Critical,
            )
            return

        source_folders = self._get_source_folders()
        if not source_folders:
            self._show_message(
                "Error",
                "Please select at least one folder first.",
                QMessageBox.Icon.Warning,
            )
            return

        resample_filter = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS

        # ── Subfolder safety rules (parent-mode only) ─────────────────────────
        excluded_dirs: frozenset[str] = frozenset()
        if len(source_folders) == 1:
            folder = source_folders[0]
            if self._has_subfolders(folder) and not self._folder_has_direct_images(folder):
                protected, hires = self._scan_subfolder_rules(folder)
                if protected:
                    folder_name = os.path.basename(folder)
                    protected_list = ", ".join(sorted(protected))
                    self._show_message(
                        "Warning",
                        f'"{folder_name}" could be an actual project folder with subfolder: '
                        f"[{protected_list}]. "
                        f"Make sure you are not processing final packshot outputs.",
                        QMessageBox.Icon.Warning,
                    )
                excluded_dirs = frozenset(
                    name.lower() for name in (protected + hires)
                )

        jobs = self._build_jobs(source_folders, excluded_dirs)

        all_files: list[tuple[str, str, str]] = []
        output_folders = set()
        for job in jobs:
            source = job["source"]
            output = job["output"]
            recursive = job["recursive"]
            os.makedirs(output, exist_ok=True)
            output_folders.add(output)
            excluded = job.get("excluded_dirs", frozenset())
            for filepath, filename in self._iter_image_files(source, recursive, excluded):
                all_files.append((filepath, filename, output))

        if not all_files:
            self._show_message(
                "Information",
                "No supported image files (.tif/.tiff/.png/.jpg/.jpeg) were found.",
                QMessageBox.Icon.Information,
            )
            return

        errors = []
        processed_count = 0
        skipped_count = 0
        total = len(all_files)
        self._set_processing_state(True, total=total)

        try:
            for index, (filepath, filename, output_folder) in enumerate(all_files, start=1):
                try:
                    with Image.open(filepath) as img:
                        # Skip images that are already thumbnail-sized (≤300 on both sides)
                        if img.width <= 300 and img.height <= 300:
                            skipped_count += 1
                            if self.progress_bar is not None:
                                self.progress_bar.setValue(index)
                                app = QApplication.instance()
                                if app is not None:
                                    app.processEvents()
                            continue

                        ext = os.path.splitext(filename)[1].lower()
                        original_format = img.format

                        if ext in (".png", ".tif", ".tiff") and img.mode in ("RGBA", "LA"):
                            alpha = img.getchannel("A")
                            bbox = alpha.getbbox()
                            if bbox:
                                img = img.crop(bbox)

                        img.thumbnail((300, 300), resample_filter)

                        if ext in (".jpg", ".jpeg") and img.mode != "RGB":
                            img = img.convert("RGB")

                        # Overwrite the original file in-place
                        output_path = filepath
                        save_kwargs = {"dpi": (50, 50)}
                        if ext in (".jpg", ".jpeg"):
                            save_kwargs["quality"] = 85

                        fallback_format = {
                            ".jpg": "JPEG",
                            ".jpeg": "JPEG",
                            ".png": "PNG",
                            ".tif": "TIFF",
                            ".tiff": "TIFF",
                        }.get(ext, "PNG")
                        format_to_save = original_format or fallback_format
                        img.save(output_path, format=format_to_save, **save_kwargs)
                        processed_count += 1
                except Exception as exc:
                    errors.append(f"{filepath}: {exc}")

                if self.progress_bar is not None:
                    self.progress_bar.setValue(index)
                    app = QApplication.instance()
                    if app is not None:
                        app.processEvents()
        finally:
            self._set_processing_state(False)

        if errors:
            error_preview = "\n".join(errors[:10])
            suffix = "\n..." if len(errors) > 10 else ""
            self._show_message(
                "Completed With Errors",
                (
                    f"Processed {processed_count}/{total} images "
                    f"({skipped_count} already thumbnail-sized, skipped).\n\n"
                    f"Issues:\n{error_preview}{suffix}"
                ),
                QMessageBox.Icon.Warning,
            )
            self.clear_all_fields()
            return

        self._show_message(
            "Done",
            (
                f"Thumbnail generation completed.\n"
                f"Resized: {processed_count}  |  Already small (skipped): {skipped_count}"
            ),
            QMessageBox.Icon.Information,
        )
        self.clear_all_fields()
