"""Точка входа консольной версии — из неё собирается Konvektor-cli.exe.

PyInstaller: pyinstaller --onefile --console --name Konvektor-cli app_cli.py
"""

import sys

from konvektor.cli import main

if __name__ == "__main__":
    sys.exit(main())
