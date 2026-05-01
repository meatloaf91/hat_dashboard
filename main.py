from __future__ import annotations
import sys
from PySide6.QtWidgets import QApplication

from new_ui_window import NewUIWindow


def main() -> int:
    app = QApplication(sys.argv)
    window = NewUIWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

#To Do:
# Include a feature that shows all library inputs for better searching.
# Work on the Review panel.