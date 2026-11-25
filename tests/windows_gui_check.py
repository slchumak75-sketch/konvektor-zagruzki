"""Проверка реальных виджетов Windows, включая стили кнопок."""
from konvektor.gui import KonvertorApp
from tkinter import ttk

app = KonvertorApp()
try:
    app.update()
    assert app.winfo_exists()
    assert len(app._button_images) == 3, "Скруглённые кнопки не созданы"
    assert "Pill.background" in str(ttk.Style(app).layout("Header.TButton"))
    assert "Pill.background" in str(ttk.Style(app).layout("Header.TMenubutton"))
    print("Windows Tk GUI: OK")
finally:
    app.destroy()
