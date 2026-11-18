"""Точка входа консольной версии — из неё собирается Konvertor-cli.exe.

PyInstaller: pyinstaller --onefile --console --name Konvertor-cli app_cli.py
"""

import sys

from konvektor.cli import main

if __name__ == "__main__":
    sys.exit(main())
