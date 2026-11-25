"""Проверка настоящих заголовков и цветов виджетов в Windows CI."""
from konvektor.gui import KonvertorApp, PREVIEW_COLUMNS, NAVY, BG_SOFT
from tkinter import ttk

app = KonvertorApp()
try:
    app.update()
    assert app.winfo_exists()
    assert "headings" in tuple(map(str, app.tk.splitlist(app.preview.cget("show"))))
    for title in PREVIEW_COLUMNS:
        assert app.preview.heading(title, "text") == title
    style = ttk.Style(app)
    assert style.lookup("Grid.Treeview.Heading", "foreground").lower() == NAVY.lower()
    assert style.lookup("Grid.Treeview.Heading", "background").lower() == BG_SOFT.lower()
    assert NAVY != BG_SOFT
    assert "Pill.background" not in str(style.layout("Header.TButton"))
    print("Windows Tk GUI и заголовки: OK")
finally:
    app.destroy()
