"""Точка входа оконной версии — из неё собирается Konvektor.exe.

PyInstaller: pyinstaller --onefile --noconsole --name Konvektor app_gui.py
"""

import sys

from konvektor.gui import main

if __name__ == "__main__":
    sys.exit(main())
