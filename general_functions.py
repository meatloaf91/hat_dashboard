from __future__ import annotations

from typing import Callable

from PySide6.QtWidgets import (
    QCheckBox,
    QLineEdit,
    QProgressBar,
    QRadioButton,
    QTextEdit,
    QWidget,
)


THUMBNAIL_OUTPUT_HINT = "Auto-generated: <source>_thumbnails"


def clear_widget_inputs(widget: QWidget) -> None:
    """Generic clear for common input widgets inside a container widget."""
    if widget is None:
        return

    for field in widget.findChildren(QLineEdit):
        field.clear()
    for field in widget.findChildren(QTextEdit):
        field.clear()
    for field in widget.findChildren(QCheckBox):
        field.setChecked(False)
    for field in widget.findChildren(QRadioButton):
        field.setChecked(False)
    for bar in widget.findChildren(QProgressBar):
        bar.setValue(0)
        bar.setVisible(False)


def clear_thumbnail_inputs(generator) -> None:
    """Reset Thumbnail Generator module state + UI fields."""
    if generator is None:
        return

    generator.images_folder = ""
    generator.output_location = ""
    generator.selected_folders = []
    generator.input_images_folder.clear()
    generator.input_output_location.setText(THUMBNAIL_OUTPUT_HINT)

    if generator.progress_bar is not None:
        generator.progress_bar.setRange(0, 100)
        generator.progress_bar.setValue(0)
        generator.progress_bar.setVisible(False)


def clear_other_panel_inputs(ui, active_panel: str) -> None:
    """
    Clear all non-active panels.
    `active_panel` should be one of: status, thumbnail, packshot
    """

    def _safe(callable_obj: Callable[[], None] | None) -> None:
        if callable_obj is None:
            return
        try:
            callable_obj()
        except Exception:
            return

    clearers = {
        "status": lambda: _safe(getattr(getattr(ui, "status_collector", None), "clear_all_fields", None)),
        "thumbnail": lambda: _safe(getattr(getattr(ui, "thumbnail_generator", None), "clear_all_fields", None)),
        "packshot": lambda: _safe(getattr(getattr(ui, "packshot_naming_generator", None), "clear_all_fields", None)),
    }

    for panel_name, clear_fn in clearers.items():
        if panel_name != active_panel:
            clear_fn()

    # Optional fallback: clear panel widgets if module clearers do not exist.
    panel_roots = {
        "status": getattr(ui, "right_panel_tracker", None),
        "thumbnail": getattr(ui, "right_panel_thumbnail", None),
        "packshot": getattr(ui, "right_panel_packshot", None),
    }
    for panel_name, root in panel_roots.items():
        if panel_name != active_panel:
            clear_widget_inputs(root)
