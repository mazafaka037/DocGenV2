import json
import os
import re
import zipfile
import tempfile
import traceback
from pathlib import Path
from datetime import datetime
import tkinter as tk
import customtkinter as ctk
import platform

import sys
from pathlib import Path

def get_base_path() -> Path:
    """Return folder with static resources in dev and in bundled apps.
       Handles: direct script run, PyInstaller (--onefile and --onedir) and py2app."""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            return Path(meipass)
        exe = Path(sys.argv[0]).resolve()
        maybe_resources = exe.parent.parent / "Resources"
        if maybe_resources.exists():
            return maybe_resources
        return exe.parent
    return Path(__file__).resolve().parent

base = get_base_path()

DEFAULTS_JSON = base / "defaults.json"
TEMPLATES_JSON = base / "templates.json"
WORKERS_JSON = base / "workers.json"
COUNTERS_JSON = base / "counters.json"
SETTINGS_JSON = base / "settings.json"
BRIGADES_JSON = base / "brigades.json"
TEMPLATE_PERMIT = base / "template_permit.docx"
TEMPLATE_SPISOK = base / "template_spisok.docx"
TEMPLATE_ORDER = base / "template_order.docx"
TEMPLATE_PB_ORDER = base / "template_pb_order.docx"
F_1606 = base / "наряд-допуск (1606-А).docx"
HUMAN_SAFE_NUMB = base / "{human} ({safe_numb}).docx"
TOGGLE_APPEARANCE_MODE_BETWEEN_LIGHT_AND_DARK_AND_PERSIST_IN_SETTINGS_JSON = base / "Toggle appearance mode between Light and Dark and persist in settings.json"

ctk.set_appearance_mode('Light')  # System, Dark, Light
ctk.set_default_color_theme('blue')



# --- Автоматические биндинги для полей ввода ---
# Оборачиваем конструкторы ctk.CTkEntry, tk.Text и ctk.CTkEntry так, чтобы у всех полей были стандартные горячие клавиши и поведение клика.
# Реализуем copy/paste/cut через clipboard, и не возвращаем "break" для клика, чтобы сохранить стандартное поведение фокуса.
try:
    _orig_tk_Entry = ctk.CTkEntry
    _orig_tk_Text = tk.Text
    _orig_ttk_Entry = ctk.CTkEntry
except Exception:
    _orig_tk_Entry = None

def _add_edit_bindings(w):
    try:
        # Copy
        def copy_ev(e=None):
            try:
                # for Text widgets, selection_get works; for Entry too.
                try:
                    sel = w.selection_get()
                except Exception:
                    return
                w.clipboard_clear()
                w.clipboard_append(sel)
            except Exception:
                pass
            # let default handlers (if any) run
            return

        # Paste
        def paste_ev(e=None):
            try:
                txt = w.clipboard_get()
            except Exception:
                return
            try:
                # replace selection if any
                try:
                    sel_first = w.index("sel.first")
                    sel_last = w.index("sel.last")
                    w.delete(sel_first, sel_last)
                except Exception:
                    pass
                try:
                    w.insert("insert", txt)
                except Exception:
                    # for Entry widgets, use icursor and insert at position
                    try:
                        pos = w.index("insert")
                        w.insert(pos, txt)
                    except Exception:
                        pass
            except Exception:
                pass
            return

        # Cut
        def cut_ev(e=None):
            try:
                try:
                    sel = w.selection_get()
                except Exception:
                    sel = ""
                if sel:
                    w.clipboard_clear()
                    w.clipboard_append(sel)
                    try:
                        w.delete("sel.first", "sel.last")
                    except Exception:
                        # Entry may require numeric indices
                        try:
                            start = w.index("sel.first")
                            end = w.index("sel.last")
                            w.delete(start, end)
                        except Exception:
                            pass
            except Exception:
                pass
            return

        # Click: set cursor and focus, but do not block event propagation
        def on_click(e):
            try:
                # for Text: index like "line.column", for Entry: numeric index
                if hasattr(w, "index"):
                    try:
                        idx = w.index("@%d,%d" % (e.x, e.y))
                        try:
                            w.mark_set("insert", idx)
                        except Exception:
                            try:
                                w.icursor(idx)
                            except Exception:
                                pass
                    except Exception:
                        pass
                w.focus_set()
            except Exception:
                try:
                    w.focus_set()
                except Exception:
                    pass
            # do not return "break" so default behavior (selection, click) still works
            return

        # Bindings for common accelerators (Ctrl/Command)
        w.bind("<Control-c>", lambda e: copy_ev(e), add=True)
        w.bind("<Control-x>", lambda e: cut_ev(e), add=True)
        w.bind("<Control-v>", lambda e: paste_ev(e), add=True)
        w.bind("<Control-C>", lambda e: copy_ev(e), add=True)
        w.bind("<Control-X>", lambda e: cut_ev(e), add=True)
        w.bind("<Control-V>", lambda e: paste_ev(e), add=True)
        # macOS Command key
        w.bind("<Command-c>", lambda e: copy_ev(e), add=True)
        w.bind("<Command-x>", lambda e: cut_ev(e), add=True)
        w.bind("<Command-v>", lambda e: paste_ev(e), add=True)

        # Click-to-set-insert (works for Entry and Text). Use add=True to not override existing binds.
        w.bind("<Button-1>", on_click, add=True)
    except Exception:
        pass

def _Entry(*args, **kwargs):
    w = _orig_tk_Entry(*args, **kwargs)
    _add_edit_bindings(w)
    return w

def _Text(*args, **kwargs):
    w = _orig_tk_Text(*args, **kwargs)
    _add_edit_bindings(w)
    return w

def _ttk_Entry(*args, **kwargs):
    w = _orig_ttk_Entry(*args, **kwargs)
    _add_edit_bindings(w)
    return w

# Apply replacements if available
try:
    if _orig_tk_Entry is not None:
        ctk.CTkEntry = _Entry
        # tk.Text = _Text  # disabled to avoid scrolledtext import error
        ctk.CTkEntry = _ttk_Entry
except Exception:
    pass

# --- конец биндингов ---


from tkinter import ttk, Toplevel, messagebox, scrolledtext, filedialog
import ctypes
from docxtpl import DocxTemplate

# --- helper: create a CTk-styled toplevel, fallback to Toplevel with CTk frame bg ---
def make_ctk_toplevel(root, title="", geometry=None):
    """
    Try to create ctk.CTkToplevel (so window chrome follows CTk theme).
    If not available, create plain Toplevel and set its bg to CTk frame fg_color (if possible).
    Ensure the window is deiconified, lifted and focused to avoid immediate minimization on some systems.
    """
    try:
        win = ctk.CTkToplevel(root)
        if geometry:
            try: win.geometry(geometry)
            except Exception: pass
        if title:
            try: win.title(title)
            except Exception: pass
        try:
            win.update_idletasks()
            win.deiconify(); win.lift(); win.focus_force()
            try:
                win.attributes("-topmost", True)
                win.after(100, lambda: win.attributes("-topmost", False))
            except Exception:
                pass
        except Exception:
            pass
        return win
    except Exception:
        win = Toplevel(root)
        if geometry:
            try: win.geometry(geometry)
            except Exception: pass
        if title:
            try: win.title(title)
            except Exception: pass
        try:
            tmp = ctk.CTkFrame(win)
            bg = tmp.cget("fg_color")
            tmp.destroy()
            try:
                win.configure(bg=bg)
            except Exception:
                pass
        except Exception:
            pass
        try:
            win.update_idletasks()
            win.deiconify(); win.lift(); win.focus_force()
            try:
                win.attributes("-topmost", True)
                win.after(100, lambda: win.attributes("-topmost", False))
            except Exception:
                pass
        except Exception:
            pass
        return win



def _do_add_selected_inner(sel_window, names, vars_list):
    try:
        picks = [name for i,name in enumerate(names) if vars_list[i].get()]
        if not picks:
            messagebox.showwarning("Внимание","Выберите хотя бы одного работника")
            return
        s = widgets["spisok_workers"]["widget"]
        try:
            existing = [ln.strip() for ln in s.get("1.0","end").splitlines() if ln.strip()]
        except Exception:
            existing = []
        to_add = [n for n in picks if n not in existing]
        if to_add:
            add_text = "\n".join(to_add)
            try:
                if existing:
                    s.insert("end", "\n" + add_text)
                else:
                    s.insert("1.0", add_text)
            except Exception:
                try:
                    s.insert("1.0", add_text)
                except Exception:
                    pass
            schedule_autosave(); refresh_permit_workers_display()
        try:
            sel_window.destroy()
        except Exception:
            pass
    except Exception:
        try:
            sel_window.destroy()
        except Exception:
            pass

def _do_add_all_inner(sel_window, all_names):
    try:
        if not all_names:
            messagebox.showwarning("Внимание","Список работников пуст")
            return
        s = widgets["spisok_workers"]["widget"]
        try:
            existing = [ln.strip() for ln in s.get("1.0","end").splitlines() if ln.strip()]
        except Exception:
            existing = []
        to_add = [n for n in all_names if n not in existing]
        if not to_add:
            messagebox.showinfo("Внимание", "Нечего добавлять — все работники уже в списке.")
            try:
                sel_window.destroy()
            except Exception:
                pass
            return
        add_text = "\n".join(to_add)
        try:
            if existing:
                s.insert("end", "\n" + add_text)
            else:
                s.insert("1.0", add_text)
        except Exception:
            try:
                s.insert("1.0", add_text)
            except Exception:
                pass
        schedule_autosave(); refresh_permit_workers_display()
        try:
            sel_window.destroy()
        except Exception:
            pass
    except Exception:
        try:
            sel_window.destroy()
        except Exception:
            pass


# -------------------------
# Настройки / mapping
# -------------------------
mapping = {
    "shared": {
        "fio": {"label": "Ответственный руководитель (И.п/Р.п/Д.п)"},

        "fio1": {"label": "Ответственный исполнитель (ФИО)"},
        "fio3": {"label": "Лицо, выдавшее наряд (ФИО)"},
        "a": {"label": "Выдан"},
        "aa": {"label": "Выдан — месяц.год (формируется из a)"},
        "b": {"label": "Действителен до"},
        "bb": {"label": "Действителен до — месяц.год (формируется из b)"},
        "d": {"label": "Начало работ"},
        "dd": {"label": "Начало работ — месяц.год (формируется из d)"},
        "e": {"label": "Окончание работ"},
        "ee": {"label": "Окончание работ — месяц.год (формируется из e)"},
        "location_address": {"label": "Место выполнения работ"},
        "numb": {"label": "Глобальный номер документа"}
    },
    "permit": {
        "work_scope": {"label": "На выполнение работ"},
        "content": {"label": "Содержание работ"},
        "terms": {"label": "Условия проведения работ"},
        "w0": {"label": "Работник 0 (Фамилия И.О. для наряда)"},
        "w1": {"label": "Работник 1 (Фамилия И.О.)"},
        "w2": {"label": "Работник 2 (Фамилия И.О.)"},
        "w3": {"label": "Работник 3 (Фамилия И.О.)"},
        "w4": {"label": "Работник 4 (Фамилия И.О.)"},
        "w5": {"label": "Работник 5 (Фамилия И.О.)"},
        "w6": {"label": "Работник 6 (Фамилия И.О.)"},
        "w7": {"label": "Работник 7 (Фамилия И.О.)"},
        "w8": {"label": "Работник 8 (Фамилия И.О.)"},
        "w9": {"label": "Работник 9 (Фамилия И.О.)"},
        "w10": {"label": "Работник 10 (Фамилия И.О.)"},
        "w11": {"label": "Работник 11 (Фамилия И.О.)"},
        "materials": {"label": "Материалы"},
        "tools": {"label": "Инструменты"},
        "devices": {"label": "Приспособления"},
        "time": {"label": "Время"},
        "hazards": {"label": "Опасные и вредные факторы"}
    },
    "spisok": {
        "predmet": {"label": "На выполнение работ и содержание (в Р.п.)"},
        "workers": {"label": "Список работников"},
        # Поля position и place убраны из UI по запросу, но могут присутствовать в worker-карточках.
        "position": {"label": "Должность (общая)"},
        "birth": {"label": "Дата рождения (общая)"},
        "pass": {"label": "Серия и номер (общая)"},
        "place": {"label": "Кем выдан (общая)"}
    }
}

# --- файлы и хранилище ---
if platform.system() == 'Darwin':
    APPDIR = Path.home() / 'Library' / 'Application Support' / 'DocGenApp'
else:
    APPDIR = Path(os.getenv('APPDATA') or Path.home()) / 'DocGenApp'
APPDIR.mkdir(parents=True, exist_ok=True)

DEFAULTS_FILE = APPDIR / "defaults.json"
TEMPLATES_FILE = APPDIR / "templates.json"
WORKERS_FILE = APPDIR / "workers.json"
COUNTERS_FILE = APPDIR / "counters.json"
SETTINGS_FILE = APPDIR / "settings.json"
BRIGADES_FILE = APPDIR / "brigades.json"

TEMPLATE_PERMIT = TEMPLATE_PERMIT
TEMPLATE_SPISOK = TEMPLATE_SPISOK
TEMPLATE_ORDER = TEMPLATE_ORDER
TEMPLATE_PB_ORDER = TEMPLATE_PB_ORDER

def ensure_storage():
    if not DEFAULTS_FILE.exists(): DEFAULTS_FILE.write_text("{}", encoding="utf-8")
    if not TEMPLATES_FILE.exists(): TEMPLATES_FILE.write_text(json.dumps({"fields": {}}, ensure_ascii=False, indent=2), encoding="utf-8")
    if not WORKERS_FILE.exists(): WORKERS_FILE.write_text("[]", encoding="utf-8")
    if not COUNTERS_FILE.exists(): COUNTERS_FILE.write_text(json.dumps({"numb":F_1606}, ensure_ascii=False, indent=2), encoding="utf-8")
    if not SETTINGS_FILE.exists(): SETTINGS_FILE.write_text(json.dumps({"autosave": True}, ensure_ascii=False, indent=2), encoding="utf-8")
    if not BRIGADES_FILE.exists(): BRIGADES_FILE.write_text("[]", encoding="utf-8")
ensure_storage()

with open(DEFAULTS_FILE, encoding="utf-8") as f:
    defaults = json.load(f)
with open(TEMPLATES_FILE, encoding="utf-8") as f:
    templates = json.load(f)
with open(WORKERS_FILE, encoding="utf-8") as f:
    workers_db = json.load(f)
with open(COUNTERS_FILE, encoding="utf-8") as f:
    counters = json.load(f)
with open(SETTINGS_FILE, encoding="utf-8") as f:
    settings = json.load(f)

# Apply saved appearance mode (Light/Dark) from settings
try:
    appearance = settings.get("appearance_mode", "Light")
    ctk.set_appearance_mode(appearance)
except Exception:
    pass
with open(BRIGADES_FILE, encoding="utf-8") as f:
    try:
        brigades_db = json.load(f)
        if not isinstance(brigades_db, list): brigades_db = []
    except Exception:
        brigades_db = []

def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    try:
        # Если сохраняются шаблоны, обновим все UI элементы, связанные с шаблонами
        try:
            from pathlib import Path as _Path
            pth = _Path(path)
            if "templates.json" in str(pth) or pth == TEMPLATES_FILE:
                if "refresh_all_template_ui" in globals():
                    refresh_all_template_ui()
        except Exception:
            # fallback: если path задан как строка
            try:
                if isinstance(path, str) and "templates.json" in path:
                    if "refresh_all_template_ui" in globals():
                        refresh_all_template_ui()
            except Exception:
                pass
    except Exception:
        pass

def refresh_all_template_ui():
    """
    Вызывает все функции с префиксом refresh_ чтобы обновить UI, связанный с шаблонами/combobox'ами.
    Это гарантирует, что после изменения (удаления/добавления) шаблонов все dropdown-ы обновятся.
    """
    for name, obj in list(globals().items()):
        try:
            if callable(obj) and name.startswith("refresh_"):
                obj()
        except Exception:
            pass



def tpl_name(item):
    '''Return a display name for a template entry (handles dicts and legacy strings).'''
    try:
        if isinstance(item, dict):
            name = item.get("name", "") or (item.get("content", "").splitlines()[0] if item.get("content","") else "")
            return str(name)
        s = str(item)
        lines = s.splitlines()
        return lines[0] if lines else s
    except Exception:
        try:
            return str(item)
        except Exception:
            return ""

def tpl_content(item):
    '''Return the full content for a template entry (dict or string).'''
    try:
        if isinstance(item, dict):
            return item.get("content", "") or ""
        return str(item)
    except Exception:
        return ""


def save_brigades_db():
    global brigades_db
    save_json(BRIGADES_FILE, brigades_db)

# UI constants (подогнаны чтобы кнопки внизу помещали текст)
# --- UI: единые шрифт/размеры/размеры полей ---
BUTTON_WIDTH_DEFAULT = 20
BUTTON_WIDTH_LARGE = 36

#  Настройте здесь желаемый шрифт и размеры:
if platform.system() == 'Darwin':
    UI_FONT_FAMILY = ''  # let Tk use system default on macOS (San Francisco)
else:
    UI_FONT_FAMILY = 'Segoe UI'   # можно поставить "Segoe UI", "Tahoma" или др.
UI_FONT_SIZE   = 12        # базовый размер шрифта для интерфейса
UI_BUTTON_SIZE = 11        # размер шрифта для кнопок
UI_BUTTON_LARGE_SIZE = 12  # для больших кнопок (если нужно)

# Константы для единого размера полей и текстовых блоков:
ENTRY_WIDTH_DEFAULT = 80   # используйте вместо литералов width=80/100
TEXT_HEIGHT_DEFAULT = 4    # используйте вместо литералов height=4/18 и т.д.

# Готовые к использованию кортежи (как раньше)
BUTTON_FONT = (UI_FONT_FAMILY, UI_BUTTON_SIZE)
BUTTON_FONT_LARGE = (UI_FONT_FAMILY, UI_BUTTON_LARGE_SIZE)
DEFAULT_FONT = (UI_FONT_FAMILY, UI_FONT_SIZE)

# Helper to create a combobox using CTk if available, falling back to ttk.Combobox.

# Registry for template comboboxes so they can be refreshed when templates change
template_combos = []  # list of tuples (key, combo_instance)






class TemplateCombo:
    """
    Simplified TemplateCombo: NO display text field (only a toggle button).
    The button opens a popup list of templates; selecting an item sets the variable.
    Clicking the button when popup is open will close it (arrow toggles ▲/▼).
    API: get(), set(val), set_values(list), bind(event, callback), pack/grid/place proxies.
    """
    def __init__(self, parent, var=None, values=None, width=58, show_display=False):
        self.parent = parent
        self.var = var if var is not None else tk.StringVar()
        self.values = list(values) if values else []
        self.container = ctk.CTkFrame(parent) if hasattr(ctk, 'CTkFrame') else ttk.Frame(parent)
        # We intentionally do NOT create a display widget; only the button remains.
        self.show_display = False

        # Create only the button with arrow that toggles the popup
        try:
            self.button = ctk.CTkButton(self.container, text="▾", width=32, command=self._toggle_popup)
            self.button.grid(row=0, column=0, padx=(0,0))
        except Exception:
            self.button = ttk.Button(self.container, text="▾", command=self._toggle_popup)
            try:
                self.button.grid(row=0, column=0, padx=(0,0))
            except Exception:
                pass

        # make container column config consistent
        try:
            self.container.grid_columnconfigure(0, weight=1)
        except Exception:
            pass

        self.popup = None
        self._callbacks = []

    def _toggle_popup(self):
        # Explicitly toggle: if popup exists -> close, else -> open
        if self.popup and getattr(self.popup, "winfo_exists", lambda: False)():
            self._close_popup()
        else:
            self._open_popup()

    def _open_popup(self):
        # create popup to show values; if already exists, do nothing
        try:
            if self.popup and getattr(self.popup, "winfo_exists", lambda: False)():
                return
            self.popup = tk.Toplevel(self.container)
            self.popup.wm_overrideredirect(True)
            # place below the container/button
            x = self.container.winfo_rootx()
            y = self.container.winfo_rooty() + self.container.winfo_height()
            self.popup.wm_geometry(f"+{x}+{y}")
            # build listbox
            lb = tk.Listbox(self.popup, exportselection=False)
            for v in self.values:
                lb.insert("end", v)
            lb.pack(side="left", fill="both", expand=True)
            # scroll bar
            sb = ttk.Scrollbar(self.popup, command=lb.yview)
            lb.config(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            # bind selection
            def on_select(evt):
                sel = None
                try:
                    idx = lb.curselection()
                    if idx:
                        sel = lb.get(idx)
                except Exception:
                    pass
                if sel is not None:
                    try:
                        self.set(sel)
                    except Exception:
                        try:
                            self.var.set(sel)
                        except Exception:
                            pass
                    # call callbacks bound to <<ComboboxSelected>>
                    for cb in list(self._callbacks):
                        try:
                            cb(None)
                        except Exception:
                            pass
                    self._close_popup()
            lb.bind("<<ListboxSelect>>", on_select)

            # change arrow to up immediately
            try:
                if hasattr(self.button, "configure"):
                    try:
                        self.button.configure(text="▴")
                    except Exception:
                        pass
            except Exception:
                pass

            # clicking outside should close popup: bind focus out on the popup and global click
            self.popup.bind("<FocusOut>", lambda e: self._close_popup())
            # Also bind a global click on root to close when clicking outside
            try:
                root = self.container.winfo_toplevel()
                self._global_bind_after_id = None
                try:
                    # delay binding so the current click that opened the popup is not captured
                    self._global_bind_after_id = root.after(100, lambda: root.bind_all("<Button-1>", self._global_click_handler, add="+"))
                except Exception:
                    self._global_bind_after_id = None
            except Exception:
                pass

            # try to focus the popup so FocusOut works
            try:
                self.popup.focus_force()
            except Exception:
                pass

        except Exception:
            self.popup = None

    def _global_click_handler(self, event):
        # close popup if click occurred outside the popup window
        try:
            if not (self.popup and getattr(self.popup, "winfo_exists", lambda: False)()):
                return
            w = event.widget
            # If clicked widget is inside the popup, ignore
            if str(w).startswith(str(self.popup)):
                return
            # If clicked widget is the button, allow toggle to handle it (do not close here)
            if w is self.button or str(w).startswith(str(self.button)):
                return
            self._close_popup()
        except Exception:
            pass

    def _close_popup(self):
        try:
            # unbind global click handler to avoid leaks
            try:
                root = self.container.winfo_toplevel()
                try:
                    # cancel scheduled global bind if it wasn't executed yet
                    if getattr(self, '_global_bind_after_id', None):
                        try:
                            root.after_cancel(self._global_bind_after_id)
                        except Exception:
                            pass
                        self._global_bind_after_id = None
                except Exception:
                    pass
                try:
                    root.unbind_all("<Button-1>")
                except Exception:
                    pass
            except Exception:
                pass

            if self.popup and getattr(self.popup, "winfo_exists", lambda: False)():
                try:
                    self.popup.destroy()
                except Exception:
                    try:
                        self.popup.withdraw()
                    except Exception:
                        pass
            self.popup = None
            # restore arrow down
            try:
                if hasattr(self.button, "configure"):
                    try:
                        self.button.configure(text="▾")
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            self.popup = None

    def set_values(self, values):
        # rebuild values list; if popup open, recreate it to reflect changes
        try:
            self.values = list(values)
        except Exception:
            self.values = []
        if self.popup and getattr(self.popup, "winfo_exists", lambda: False)():
            try:
                self._close_popup()
                self._open_popup()
            except Exception:
                pass

    def bind(self, ev, callback):
        # accept '<<ComboboxSelected>>' and store callbacks
        if ev == '<<ComboboxSelected>>':
            try:
                if callable(callback):
                    self._callbacks.append(callback)
            except Exception:
                pass
        else:
            try:
                self.container.bind(ev, callback)
            except Exception:
                try:
                    self.button.bind(ev, callback)
                except Exception:
                    pass

    def grid(self, *a, **k): return self.container.grid(*a, **k)
    def pack(self, *a, **k): return self.container.pack(*a, **k)
    def place(self, *a, **k): return self.container.place(*a, **k)
    def widget(self): return self.button

    def get(self):
        try:
            return self.var.get()
        except Exception:
            return ""

    def set(self, val):
        try:
            self.var.set(val)
        except Exception:
            pass


def make_combo(parent, textvariable, values, width=58, key=None, show_display=True):
    """
    Factory: returns a TemplateCombo instance. If key provided, registers combo in template_combos.
    """
    combo = TemplateCombo(parent, textvariable, values=values, width=width, show_display=show_display)
    if key is not None:
        try:
            template_combos.append((key, combo))
        except Exception:
            pass
    return combo

def refresh_all_template_combos():
    # Update values for all registered template combos based on templates dict
    try:
        for key, combo in template_combos:
            vals_raw = templates.get("fields", {}).get(key, [])
            vals = [tpl_name(it) for it in vals_raw]
            combo.set_values(vals)
            # clear selection if current value not in vals
            try:
                cur = combo.get()
                if cur not in vals:
                    combo.set("")
            except Exception:
                pass
    except Exception:
        pass

# Provide compatibility name used elsewhere
refresh_all_template_ui = refresh_all_template_combos
def make_button(parent, text, command=None, width=None, font=None):
    w = BUTTON_WIDTH_DEFAULT if width is None else width
    btn = ctk.CTkButton(parent, text=text, width=w, command=command)
    try:
        if font is not None:
            btn.configure(font=font)
        else:
            btn.configure(font=BUTTON_FONT)
    except Exception:
        pass
    return btn

# Регексы для эскейпинга одиночных фигурных скобок
_single_open_re = re.compile(r'(?<!\{)\{(?!\{)')
_single_close_re = re.compile(r'(?<!\})\}(?!\})')
_placeholder_re = re.compile(r'(?<!\{)\{([\w\.\-]+)\}(?!\})', flags=re.UNICODE)

def create_escaped_docx_copy(src_path: Path) -> Path:
    tmp_fd, tmp_name = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(src_path, 'r') as zin, zipfile.ZipFile(tmp_path, 'w') as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    try:
                        text = data.decode('utf-8')
                    except Exception:
                        text = data.decode('utf-8', errors='replace')
                    text = _placeholder_re.sub(r'{{ \1 }}', text)
                    text = _single_open_re.sub('{{', text)
                    text = _single_close_re.sub('}}', text)
                    data = text.encode('utf-8')
                zout.writestr(item, data)
    except Exception:
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise
    return tmp_path

def analyze_template_for_jinja_issues(path: Path, target_name_prefix="diag") -> Path:
    diag_path = APPDIR / f"{target_name_prefix}_{path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    try:
        with zipfile.ZipFile(path, 'r') as zin, open(diag_path, "w", encoding="utf-8") as out:
            out.write(f"Diagnostic dump for: {path}\nGenerated: {datetime.now().isoformat()}\n\n")
            for item in zin.infolist():
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    raw = zin.read(item.filename)
                    try:
                        text = raw.decode("utf-8")
                    except Exception:
                        text = raw.decode("utf-8", errors='replace')
                    hits = []
                    for m in re.finditer(r'(\{\{|\}\}|\{\%|\%\}|\{|\})', text):
                        s = max(0, m.start()-80)
                        e = min(len(text), m.end()+80)
                        ctx = text[s:e].replace("\n", " ")
                        hits.append((m.group(0), m.start(), ctx))
                    out.write(f"--- {item.filename} ---\n")
                    if not hits:
                        out.write("No brace/jinja tokens found.\n\n")
                    else:
                        for token, pos, ctx in hits:
                            out.write(f"Token: {token} at pos {pos}\nContext: {ctx}\n\n")
            out.write("\n\nHints:\n- Look for broken Jinja tags split by Word (parts of {{ ... }} or {% ... %} separated by formatting).\n")
    except Exception as e:
        try:
            with open(diag_path, "w", encoding="utf-8") as out:
                out.write(f"Failed to analyze: {e}\n")
        except Exception:
            pass
    return diag_path

def render_docx_safely(template_path: Path, ctx: dict, out_path: str):
    last_exc = None
    try:
        tpl = DocxTemplate(str(template_path))
        tpl.render(ctx)
        tpl.save(out_path)
        return
    except Exception as e:
        last_exc = e
        try:
            diag_orig = analyze_template_for_jinja_issues(template_path, "orig_diag")
        except Exception:
            diag_orig = None

    try:
        tmp = create_escaped_docx_copy(template_path)
        try:
            tpl = DocxTemplate(str(tmp))
            tpl.render(ctx)
            tpl.save(out_path)
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass
            return
        finally:
            if tmp.exists():
                try:
                    tmp.unlink(missing_ok=True)
                except Exception:
                    pass
    except Exception as e2:
        try:
            if 'tmp' in locals() and tmp and tmp.exists():
                diag_esc = analyze_template_for_jinja_issues(tmp, "escaped_diag")
            else:
                diag_esc = None
        except Exception:
            diag_esc = None

        tb1 = "".join(traceback.format_exception_only(type(last_exc), last_exc)) if last_exc else ""
        tb2 = "".join(traceback.format_exception_only(type(e2), e2))
        msg = f"Render failed (original): {tb1}\nAttempt with escaped/converted copy failed: {tb2}\n"
        if diag_orig:
            msg += f"\nDiagnostic log for original template: {diag_orig}\n"
        if diag_esc:
            msg += f"\nDiagnostic log for escaped copy: {diag_esc}\n"
        msg += ("\nПодсказки:\n- Откройте указанный файл с диагностикой и найдите проблемный фрагмент.\n")
        raise RuntimeError(msg)

# -------------------------
# GUI init
# -------------------------
root = ctk.CTk()

# Global mouse-wheel handler: scroll the widget under cursor if it supports yview
# Наладим именованные шрифты — это гарантирует, что ttk и стандартные виджеты
# унаследуют единый шрифт/размер.
import tkinter.font as tkfont
try:
    default = tkfont.nametofont("TkDefaultFont")
    default.configure(family=UI_FONT_FAMILY, size=UI_FONT_SIZE)
    # применим к основным именованным шрифтам (если они есть в текущей теме)
    for name in ("TkTextFont", "TkHeadingFont", "TkMenuFont", "TkFixedFont"):
        try:
            tkfont.nametofont(name).configure(family=UI_FONT_FAMILY, size=UI_FONT_SIZE)
        except Exception:
            pass
except Exception:
    pass

if platform.system() == 'Windows':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
try:
    root.tk.call('tk', 'scaling', 1.25)
except Exception:
    pass
root.title("DocGen v2")
root.geometry("1180x780")

# NOTE: статус-бар "Готово" удалён по требованию — автосохранение теперь тихое,
# уведомление о сохранении показывается ТОЛЬКО при явном нажатии "Сохранить профиль".

autosave_var = tk.BooleanVar(value=settings.get("autosave", True))
_autosave_id = None

def schedule_autosave(delay=800):
    if not autosave_var.get(): return
    global _autosave_id
    if _autosave_id: root.after_cancel(_autosave_id)
    _autosave_id = root.after(delay, autosave_now)

def autosave_now():
    global _autosave_id
    _autosave_id = None
    # тихо сохраняем профиль без визуальной индикации
    save_profile(False)

# -------------------------
# Утилиты
# -------------------------
def parse_ddmmyyyy(text):
    t=(text or "").strip()
    if not t: return "", ""
    p=t.split(".")
    return (p[0].zfill(2), f"{p[1].zfill(2)}.{p[2]}") if len(p)==3 else ("","")

def get_next_numb():
    n=counters.get("numb",1606)
    counters["numb"]=n+1
    save_json(COUNTERS_FILE,counters)
    return n

def short_name(full):
    s=(full or "").strip()
    if not s: return ""
    parts = s.split()
    if len(parts)==1: return parts[0]
    surname = parts[0]
    initials = ""
    for p in parts[1:3]:
        if p: initials += p[0].upper() + "."
    return f"{surname} {initials}"

# -------------------------
# Mousewheel support helper
# -------------------------
def add_mousewheel_support(widget):
    """Legacy placeholder: no global mousewheel handling. Left empty to keep compatibility."""
    try:
        return
    except Exception:
        pass
    """
    Adds mousewheel scrolling to the widget for multiple platforms.
    Works for widgets exposing yview_scroll (Text, Listbox, ScrolledText, etc.).
    """
    def _on_mousewheel(event):
        # determine delta in "units" for yview_scroll
        if hasattr(event, "delta") and event.delta:
            # Windows/Mac
            # On Windows, delta usually is multiple of 120.
            delta = -1 * int(event.delta / 120)
        else:
            # X11: Button-4 = up, Button-5 = down
            if event.num == 4:
                delta = -1
            else:
                delta = 1
        try:
            widget.yview_scroll(delta, "units")
        except Exception:
            try:
                # some widgets expose .scroll
                widget.scroll(delta, "units")
            except Exception:
                pass

    # bind directly to widget (safer than bind_all)
    try:
        widget.bind("<MouseWheel>", _on_mousewheel)   # Windows / Mac
        widget.bind("<Button-4>", _on_mousewheel)     # Linux scroll up
        widget.bind("<Button-5>", _on_mousewheel)     # Linux scroll down
    except Exception:
        pass

# -------------------------
# GUI widgets container
# -------------------------
widgets = {}

# Универсальный сборщик значений из всех полей
def collect_context():
    ctx = {}
    for group in (mapping["shared"], mapping["permit"], mapping["spisok"]):
        for key in group.keys():
            if key in widgets:
                w = widgets[key]
                if w["type"] == "text":
                    ctx[key] = w["widget"].get("1.0", "end").strip()
                else:
                    ctx[key] = w["widget"].get().strip()
    return ctx

tabview = ctk.CTkTabview(root)
tabview.pack(fill='both', expand=True)
# Try to center the tabbar (works if CTkTabview internal structure exposes it)
try:
    if hasattr(tabview, '_tabbar'):
        try:
            tabview._tabbar.pack_configure(anchor='center')
        except Exception:
            pass
    if hasattr(tabview, '_segmented_button'):
        try:
            tabview._segmented_button.pack_configure(anchor='center')
        except Exception:
            pass
except Exception:
    pass
# Remove unexpected top frames that may have appeared before the notebook
def _remove_top_rects():
    """Aggressively remove stray widgets placed above the notebook (visual artefacts).
    Scans both root children and root children and destroys/forgets
    widgets that are located above the notebook's Y coordinate or are tiny/empty.
    """
    try:
        root.update_idletasks()
    except Exception:
        pass
        try:
            nb_y = nb.winfo_rooty()
        except Exception:
            nb_y = None

        # helper to test and remove a widget if it's above the notebook or empty/small
        def _maybe_remove(w):
            try:
                # ignore known important containers
                if w in (frame_bot,) or getattr(w, "_preserve_widget", False):
                    return False
            except Exception:
                pass
            try:
                cy = None
                try:
                    cy = w.winfo_rooty()
                except Exception:
                    cy = None
                # if widget is placed above notebook -> remove
                if nb_y is not None and cy is not None and cy < nb_y + 10:
                    try:
                        w.destroy()
                        return True
                    except Exception:
                        try:
                            w.pack_forget()
                            return True
                        except Exception:
                            try:
                                w.place_forget()
                                return True
                            except Exception:
                                return False
                # else, if very small and empty -> remove
                try:
                    ch_h = w.winfo_height()
                    if ch_h is not None and ch_h < 30:
                        kids = w.winfo_children()
                        if not kids:
                            try:
                                w.destroy()
                                return True
                            except Exception:
                                try:
                                    w.pack_forget()
                                    return True
                                except Exception:
                                    return False
                        else:
                            empty = True
                            for k in kids:
                                try:
                                    if hasattr(k, 'cget') and 'text' in k.keys():
                                        if k.cget('text').strip():
                                            empty = False
                                            break
                                except Exception:
                                    empty = False
                                    break
                            if empty:
                                try:
                                    w.destroy()
                                    return True
                                except Exception:
                                    return False
                except Exception:
                    pass
            except Exception:
                pass
            return False

        # scan direct children of root
        try:
            for child in list(root.winfo_children()):
                try:
                    _maybe_remove(child)
                except Exception:
                    pass
        except Exception:
            pass

        # scan children of root (they may be stacked above nb)
        try:
            pass
        except Exception:
            pass

    except Exception:
        pass


def load_fio_templates_normalized():
    """
    Normalize templates['fields']['fio_combined'] into a list of dicts:
    [{"name": <first line>, "content": <full content>}, ...]
    Also migrates old string entries to dict form if found.
    """
    tpl_field = templates.setdefault("fields", {}).setdefault("fio_combined", [])
    norm = []
    changed = False
    for item in tpl_field:
        if isinstance(item, dict):
            name = item.get("name","") or (item.get("content","").splitlines()[0] if item.get("content","") else "")
            content = item.get("content","")
            norm.append({"name": name, "content": content})
        else:
            content = str(item)
            first = content.splitlines()[0] if content.splitlines() else content
            norm.append({"name": first, "content": content})
            changed = True
    if changed:
        templates["fields"]["fio_combined"] = norm
        save_json(TEMPLATES_FILE, templates)
    return norm

fio_templates_norm = load_fio_templates_normalized()
fio_combobox_var = tk.StringVar()
tabview.add("Общее")
tab_shared = tabview.tab("Общее")
ctk.CTkLabel(tab_shared, text="Ответственный руководитель").pack(anchor="w", padx=6, pady=(8,2))
tpl_frame = ctk.CTkFrame(tab_shared); tpl_frame.pack(fill="x", padx=6, pady=(2,4))
fio_combobox = make_combo(tpl_frame, fio_combobox_var, [t["name"] for t in fio_templates_norm], width=58)
fio_combobox.grid(row=0, column=0, padx=(0,6))
fio_combobox.bind('<<ComboboxSelected>>', lambda e: fio_load_template())



def refresh_fio_combobox_values():
    """
    Обновляет значения в выпадашке 'Ответственный руководитель'.
    Использует публичный метод set_values(...) у кастомного TemplateCombo, если он доступен,
    и корректно очищает текущее значение, если оно больше не присутствует в списке.
    """
    global fio_templates_norm
    try:
        fio_templates_norm = load_fio_templates_normalized()
    except Exception:
        # Если функция загрузки отсутствует или падает — оставим старый список пустым
        fio_templates_norm = []

    vals = [t.get("name", "") for t in fio_templates_norm]

    # Сначала попробуем аккуратно установить значения через API виджета
    try:
        # Если это кастомный TemplateCombo с методом set_values
        fio_combobox.set_values(vals)
    except Exception:
        # Фоллбек: если это обычный ttk.Combobox
        try:
            fio_combobox["values"] = vals
        except Exception:
            # Всё равно игнорируем — пусть останутся предыдущие значения
            pass

    # Проверяем текущее значение — если оно отсутствует, очищаем выбор
    try:
        cur = ""
        # Попробуем получить через get()
        try:
            cur = fio_combobox.get().strip()
        except Exception:
            # Или через переменную, если используется StringVar
            try:
                cur = fio_combobox_var.get().strip()
            except Exception:
                cur = ""
        if cur and cur not in vals:
            try:
                # Попытка очистить выбор через API виджета
                fio_combobox.set("")
            except Exception:
                try:
                    fio_combobox_var.set("")
                except Exception:
                    pass
    except Exception:
        # На всякий случай — сбросим переменную, если она есть
        try:
            fio_combobox_var.set("")
        except Exception:
            pass

def fio_load_template():
    name = fio_combobox_var.get().strip()
    if not name:
        return
    t = next((x for x in fio_templates_norm if x.get("name","") == name), None)
    if not t:
        t = next((x for x in fio_templates_norm if x.get("content","") == name), None)
    if t:
        fio_txt.delete("1.0","end")
        fio_txt.insert("1.0", t.get("content",""))
        schedule_autosave()

def fio_add_template():
    txt = fio_txt.get("1.0","end").rstrip()
    if not txt:
        return
    first = txt.splitlines()[0] if txt.splitlines() else txt
    templates.setdefault("fields",{}).setdefault("fio_combined", [])
    lst = templates["fields"]["fio_combined"]
    idx = next((i for i,item in enumerate(lst) if (isinstance(item, dict) and item.get("name","")==first) or (isinstance(item, str) and item.splitlines()[0]==first)), None)
    if idx is not None:
        if not messagebox.askyesno("Подтвердите", f'Шаблон с именем "{first}" существует. Перезаписать?'):
            return
        lst[idx] = {"name": first, "content": txt}
    else:
        lst.append({"name": first, "content": txt})
    save_json(TEMPLATES_FILE, templates)
    refresh_fio_combobox_values()
    schedule_autosave()

def fio_del_template():
    name = fio_combobox_var.get().strip()
    if not name:
        messagebox.showwarning("Внимание", "Выберите шаблон для удаления")
        return
    lst = templates.setdefault("fields",{}).setdefault("fio_combined", [])
    idx = next((i for i, item in enumerate(lst) if (isinstance(item, dict) and item.get("name","")==name) or (isinstance(item, str) and item.splitlines()[0]==name)), None)
    if idx is None:
        messagebox.showinfo("Внимание", "Шаблон не найден")
        return
    if not messagebox.askyesno("Подтвердите", f"Удалить шаблон '{name}'?"):
        return
    lst.pop(idx)
    save_json(TEMPLATES_FILE, templates)
    refresh_fio_combobox_values()

def fio_clear():
    fio_txt.delete("1.0","end"); schedule_autosave()

make_button(tpl_frame, text="Добавить шаблон", command=fio_add_template).grid(row=0, column=2, padx=4)
make_button(tpl_frame, text="Удалить шаблон", command=fio_del_template).grid(row=0, column=3, padx=4)
make_button(tpl_frame, text="Очистить", command=fio_clear).grid(row=0, column=4, padx=6)



# Ответственный руководитель — три однострочных поля (Именительный, Родительный, Дательный падеж)
# Create entries with placeholders for the three grammatical cases
fio_entries = []
placeholders = ["Именительный падеж", "Родительный падеж", "Дательный падеж"]
for i, ph in enumerate(placeholders):
    ent = ctk.CTkEntry(tab_shared, width=300, font=DEFAULT_FONT, placeholder_text=ph)
    ent.pack(fill="x", padx=6, pady=(0 if i>0 else 2, 4))
    ent.bind("<KeyRelease>", lambda e: schedule_autosave())
    fio_entries.append(ent)

# Wrapper to emulate tk.Text API for existing code that uses fio_txt
class FioWidget:
    def __init__(self, entries):
        self.entries = entries
    def get(self, *args, **kwargs):
        # simulate tk.Text.get("1.0","end") => join with newlines
        lines = [e.get().strip() for e in self.entries]
        return "\n".join(lines)
    def delete(self, *args, **kwargs):
        for e in self.entries:
            try:
                e.delete(0, "end")
            except Exception:
                try:
                    e.delete(0, tk.END)
                except Exception:
                    pass
    def insert(self, *args, **kwargs):
        # handle insert(index, text) or insert(text)
        if len(args) >= 2:
            text = args[1]
        elif len(args) == 1:
            text = args[0]
        else:
            text = kwargs.get('text', '')
        parts = str(text).splitlines()
        for i in range(len(self.entries)):
            try:
                self.entries[i].delete(0, "end")
            except Exception:
                pass
            if i < len(parts):
                try:
                    self.entries[i].insert(0, parts[i])
                except Exception:
                    pass

# Prefill fio entries from defaults (fio, fio2, fio4)
try:
    # Use separate keys fio, fio2, fio4 if present
    v0 = defaults.get('fio', '')
    v1 = defaults.get('fio2', '')
    v2 = defaults.get('fio4', '')
    if v0 or v1 or v2:
        try:
            if isinstance(v0, str) and v0:
                fio_entries[0].delete(0, 'end'); fio_entries[0].insert(0, v0)
        except Exception:
            pass
        try:
            if isinstance(v1, str) and v1:
                fio_entries[1].delete(0, 'end'); fio_entries[1].insert(0, v1)
        except Exception:
            pass
        try:
            if isinstance(v2, str) and v2:
                fio_entries[2].delete(0, 'end'); fio_entries[2].insert(0, v2)
        except Exception:
            pass
except Exception:
    pass

fio_txt = FioWidget(fio_entries)
widgets["fio_combined"] = {"type":"text", "widget": fio_txt}


# helper: template-enabled single-line field

def make_template_entry(parent, key, label):
    '''
    Factory: returns a TemplateCombo instance for single-line template-enabled fields.
    Stores templates consistently as dicts: {"name": <first line>, "content": <full text>}
    '''
    ctk.CTkLabel(parent, text=label).pack(anchor="w", padx=6, pady=(6,0))
    frame = ctk.CTkFrame(parent); frame.pack(fill="x", padx=6, pady=(0,2))
    combo_var = tk.StringVar()
    # Prepare name-only values from stored templates (supports dicts and legacy strings)
    vals_raw = templates.get("fields", {}).get(key, [])
    vals = [tpl_name(it) for it in vals_raw]
    combo = make_combo(frame, textvariable=combo_var, values=vals, width=60, key=key)
    combo.grid(row=0, column=0, padx=(0,4))
    combo.bind('<<ComboboxSelected>>', lambda e: load())
    def load(): 
        entry.delete(0,"end"); 
        # determine selection from combo or variable
        try:
            sel = combo.get().strip()
        except Exception:
            sel = combo_var.get().strip()
        # find template by display name
        try:
            lst = templates.get("fields", {}).get(key, [])
            t = next((it for it in lst if tpl_name(it) == sel), None)
            content = tpl_content(t) if t is not None else sel
        except Exception:
            content = sel
        entry.insert(0, content); schedule_autosave()

    def add_template():
        # save current widget content as a template (store as dict {name,content})
        w = widgets.get(key, {}).get("widget")
        if not w:
            return
        if widgets.get(key, {}).get("type") == "text":
            try:
                txt = w.get("1.0","end").strip()
            except Exception:
                txt = ""
        else:
            try:
                txt = w.get().strip()
            except Exception:
                try:
                    txt = w.get("1.0","end").strip()
                except Exception:
                    txt = ""
        if not txt:
            return
        first = txt.splitlines()[0] if txt.splitlines() else txt
        templates.setdefault("fields",{}).setdefault(key, [])
        lst = templates["fields"][key]
        # normalize legacy string entries to dicts
        for i,item in enumerate(list(lst)):
            if not isinstance(item, dict):
                lst[i] = {"name": tpl_name(item), "content": tpl_content(item)}
        # check for existing by name
        idx = next((i for i,item in enumerate(lst) if (isinstance(item, dict) and item.get("name","")==first)), None)
        if idx is not None:
            if not messagebox.askyesno("Подтвердите", f'Шаблон с именем "{first}" существует. Перезаписать?'):
                return
            lst[idx] = {"name": first, "content": txt}
        else:
            lst.append({"name": first, "content": txt})
        save_json(TEMPLATES_FILE, templates)
        try:
            refresh_all_template_ui()
        except Exception:
            pass
        # update combo values
        try:
            vals = [tpl_name(it) for it in templates.get("fields", {}).get(key, [])]
            combo.set_values(vals)
        except Exception:
            pass

    def del_template():
        v = ""
        try:
            v = combo.get().strip()
        except Exception:
            try:
                v = combo_var.get().strip()
            except Exception:
                v = ""
        if not v:
            return
        lst = templates.get("fields",{}).get(key,[])
        # find by display name
        idx = next((i for i,item in enumerate(lst) if tpl_name(item) == v), None)
        if idx is None:
            return
        removed = lst.pop(idx)
        save_json(TEMPLATES_FILE, templates)
        try:
            refresh_all_template_ui()
        except Exception:
            pass
        try:
            vals = [tpl_name(it) for it in templates.get("fields",{}).get(key,[])]
            combo.set_values(vals)
        except Exception:
            pass
        try:
            combo_var.set("")
        except Exception:
            try:
                combo.set("")
            except Exception:
                pass
        try:
            w = widgets.get(key, {}).get("widget")
            if w:
                try:
                    cur_text = w.get("1.0", "end").strip() if widgets.get(key,{},).get("type")=="text" else w.get().strip()
                except Exception:
                    cur_text = ""
                rem_content = tpl_content(removed)
                if cur_text == rem_content:
                    try:
                        if widgets.get(key, {}).get("type") == "text":
                            w.delete("1.0","end")
                        else:
                            w.delete(0,"end")
                    except Exception:
                        pass
        except Exception:
            pass

    def clear_entry():
        entry.delete(0,"end"); schedule_autosave()
    make_button(frame, text="Добавить шаблон", command=add_template).grid(row=0, column=2, padx=4)
    make_button(frame, text="Удалить шаблон", command=del_template).grid(row=0, column=3, padx=4)
    make_button(frame, text="Очистить", command=clear_entry).grid(row=0, column=4, padx=6)
    entry = ctk.CTkEntry(parent, width=100); entry.insert(0, defaults.get(key,""))
    entry.pack(fill="x", padx=6, pady=(0,6)); entry.bind("<KeyRelease>", lambda e: schedule_autosave())
    widgets[key] = {"type":"entry","widget":entry}
    #add_mousewheel_support(entry)  # harmless if no yview
    return combo, combo_var, entry

# fio1 and fio3 are template-enabled
make_template_entry(tab_shared, "fio1", mapping["shared"]["fio1"]["label"])
make_template_entry(tab_shared, "fio3", mapping["shared"]["fio3"]["label"])

# other shared fields (dates and others)
for k, meta in mapping["shared"].items():
    if k in ("fio","fio2","fio4","fio1","fio3","aa","bb","dd","ee","numb"): continue
    if k in ("a","b","d","e"):
        ctk.CTkLabel(tab_shared, text=meta["label"]).pack(anchor="w", padx=6, pady=2)
        ent = ctk.CTkEntry(tab_shared, width=100, font=DEFAULT_FONT, placeholder_text="дд.мм.гг")
        _val = defaults.get(k,"")
        if _val:
            ent.insert(0, _val)
        ent.pack(fill="x", padx=6, pady=2)
        ent.bind("<KeyRelease>", lambda e: schedule_autosave())
        widgets[k] = {"type":"entry","widget":ent}
        #add_mousewheel_support(ent)

    else:
        ctk.CTkLabel(tab_shared, text=meta["label"]).pack(anchor="w", padx=6, pady=2)
        ent = ctk.CTkEntry(tab_shared, width=80); ent.insert(0, defaults.get(k,"")); ent.pack(fill="x", padx=6, pady=2)
        try:
            ent.bind("<KeyRelease>", lambda e: schedule_autosave())
        except Exception:
            pass
        try:
            ent.bind("<FocusOut>", lambda e: schedule_autosave())
        except Exception:
            pass
        widgets[k] = {"type":"entry","widget":ent}
        #add_mousewheel_support(ent)

# numb display
ctk.CTkLabel(tab_shared, text=mapping["shared"]["numb"]["label"]).pack(anchor="w", padx=6, pady=2)
ent_num = ctk.CTkEntry(tab_shared, width=160, font=DEFAULT_FONT)
ent_num.insert(0, str(counters.get("numb", 1606)) + (
    ("-" + str(counters.get("numb_suffix", ""))) if counters.get("numb_suffix", "") else ""))
ent_num.configure(state="readonly")
ent_num.pack(anchor="w", padx=6, pady=2)
widgets["numb"] = {"type": "entry", "widget": ent_num}
make_button(tab_shared, text="Редактировать", command=lambda: edit_numb_dialog(), width=18).pack(anchor="w", padx=6,
                                                                                                 pady=(0, 6))
# TAB: Наряд-допуск

# Create a scrollable area for the "Наряд-допуск" tab so the user can scroll with mouse wheel.
tabview.add("Наряд-допуск")
tab_permit_outer = tabview.tab("Наряд-допуск")
# Canvas and vertical scrollbar (replaced with CTkScrollableFrame for consistent CTk look)
tab_permit_outer = tabview.tab("Наряд-допуск")
tab_permit = ctk.CTkScrollableFrame(tab_permit_outer)
tab_permit.pack(fill="both", expand=True, padx=6, pady=6)

permit_text_keys_with_templates = {"work_scope","content","terms","materials","tools","devices"}




def create_template_block(parent, key, label_text):
    """
    Creates a template-enabled block for the given key.
    The combo shows only template names (first line). Selecting a template inserts
    the full content into the target widget. Templates are stored as dicts:
    {"name": <first line>, "content": <full text>}.
    """
    single_line_permit = {"content", "terms", "materials", "tools", "devices",'work_scope'}

    ctk.CTkLabel(parent, text=label_text).pack(anchor="w", padx=6, pady=(8,2))
    frame_tpl = ctk.CTkFrame(parent); frame_tpl.pack(anchor="w", padx=6, pady=(0,2))

    combo_var = tk.StringVar()
    # prepare name-only values
    vals_raw = templates.get("fields", {}).get(key, [])
    vals = [ (it.get("name") if isinstance(it, dict) else (str(it).splitlines()[0] if str(it).splitlines() else str(it))) for it in vals_raw ]
    combo = make_combo(frame_tpl, combo_var, values=vals, width=58, key=key)
    combo.grid(row=0, column=0, padx=(0,4))
    combo.bind('<<ComboboxSelected>>', lambda e: load_to_target())

    def load_to_target():
        v = ""
        try:
            v = combo.get().strip()
        except Exception:
            try:
                v = combo_var.get().strip()
            except Exception:
                v = ""
        if v and widgets.get(key):
            w = widgets[key]["widget"]
            # find template by name
            try:
                lst = templates.get("fields", {}).get(key, [])
                t = next((it for it in lst if (isinstance(it, dict) and it.get("name","")==v) or (not isinstance(it, dict) and (str(it).splitlines()[0] if str(it).splitlines() else str(it))==v)), None)
                content = t.get("content","") if isinstance(t, dict) else (str(t) if t is not None else "")
                if t is None:
                    content = v
            except Exception:
                content = v
            try:
                if widgets[key].get("type") == "text":
                    try:
                        w.delete("1.0","end"); w.insert("1.0", content)
                    except Exception:
                        pass
                else:
                    try:
                        w.delete(0,"end"); w.insert(0, content)
                    except Exception:
                        pass
            except Exception:
                pass
        schedule_autosave()

    def add_template():
        # save current widget content as a template (name = first line)
        w = widgets.get(key, {}).get("widget")
        if not w:
            return
        if widgets.get(key, {}).get("type") == "text":
            try:
                txt = w.get("1.0","end").strip()
            except Exception:
                txt = ""
        else:
            try:
                txt = w.get().strip()
            except Exception:
                try:
                    txt = w.get("1.0","end").strip()
                except Exception:
                    txt = ""
        if not txt:
            return
        first = txt.splitlines()[0] if txt.splitlines() else txt
        templates.setdefault("fields",{}).setdefault(key, [])
        lst = templates["fields"][key]
        # check for existing by name
        idx = next((i for i,item in enumerate(lst) if (isinstance(item, dict) and item.get("name","")==first) or (not isinstance(item, dict) and (str(item).splitlines()[0] if str(item).splitlines() else str(item))==first)), None)
        if idx is not None:
            if not messagebox.askyesno("Подтвердите", f'Шаблон с именем "{first}" существует. Перезаписать?'):
                return
            lst[idx] = {"name": first, "content": txt}
        else:
            lst.append({"name": first, "content": txt})
        save_json(TEMPLATES_FILE, templates)
        try:
            refresh_all_template_ui()
        except Exception:
            pass
        # update combo values
        try:
            vals = [it.get("name") if isinstance(it, dict) else (str(it).splitlines()[0] if str(it).splitlines() else str(it)) for it in templates.get("fields",{}).get(key,[])]
            combo.set_values(vals)
        except Exception:
            pass

    def del_template():
        v = ""
        try:
            v = combo.get().strip()
        except Exception:
            try:
                v = combo_var.get().strip()
            except Exception:
                v = ""
        if not v:
            return
        lst = templates.get("fields",{}).get(key,[])
        idx = next((i for i,item in enumerate(lst) if (isinstance(item, dict) and item.get("name","")==v) or (not isinstance(item, dict) and (str(item).splitlines()[0] if str(item).splitlines() else str(item))==v)), None)
        if idx is None:
            return
        removed = lst.pop(idx)
        save_json(TEMPLATES_FILE, templates)
        try:
            refresh_all_template_ui()
        except Exception:
            pass
        try:
            vals = [it.get("name") if isinstance(it, dict) else (str(it).splitlines()[0] if str(it).splitlines() else str(it)) for it in templates.get("fields",{}).get(key,[])]
            combo.set_values(vals)
        except Exception:
            pass
        try:
            combo_var.set("")
        except Exception:
            try:
                combo.set("")
            except Exception:
                pass
        try:
            w = widgets.get(key, {}).get("widget")
            if w:
                try:
                    cur_text = w.get("1.0", "end").strip() if widgets.get(key,{}).get("type")=="text" else w.get().strip()
                except Exception:
                    cur_text = ""
                rem_content = removed.get("content") if isinstance(removed, dict) else str(removed)
                if cur_text == rem_content:
                    try:
                        if widgets.get(key, {}).get("type") == "text":
                            w.delete("1.0","end")
                        else:
                            w.delete(0,"end")
                    except Exception:
                        pass
        except Exception:
            pass

    make_button(frame_tpl, text="Добавить шаблон", command=add_template).grid(row=0, column=2, padx=4)
    make_button(frame_tpl, text="Удалить шаблон", command=del_template).grid(row=0, column=3, padx=4)
    
    def clear_target():
        """
        Clear the target widget associated with this template block (key).
        Works for both text widgets and single-line entries.
        """
        try:
            w = widgets.get(key, {}).get("widget")
            if not w:
                return
            if widgets.get(key, {}).get("type") == "text":
                try:
                    w.delete("1.0", "end")
                except Exception:
                    pass
            else:
                try:
                    w.delete(0, "end")
                except Exception:
                    try:
                        w.delete("1.0", "end")
                    except Exception:
                        pass
        except Exception:
            pass
        try:
            schedule_autosave()
        except Exception:
            pass
    make_button(frame_tpl, text="Очистить", command=clear_target).grid(row=0, column=4, padx=6)

    # Choose widget type based on whether the field is single-line or multi-line
    if key in single_line_permit:
        txt = ctk.CTkEntry(parent, width=100, font=DEFAULT_FONT)
        txt.insert(0, defaults.get(key,""))
        txt.pack(fill="x", padx=6, pady=2)
        try:
            txt.bind("<KeyRelease>", lambda e: schedule_autosave())
        except Exception:
            pass
        try:
            txt.bind("<FocusOut>", lambda e: schedule_autosave())
        except Exception:
            pass
        widgets[key] = {"type":"entry","widget":txt}
    else:
        txt = tk.Text(parent, height=4, font=DEFAULT_FONT)
        txt.insert("1.0", defaults.get(key,""))
        txt.pack(fill="x", padx=6, pady=2)
        txt.bind("<<Modified>>", lambda e:(e.widget.edit_modified(False), schedule_autosave()))
        widgets[key] = {"type":"text","widget":txt}


# Ensure "На выполнение работ" (work_scope) is at the top of the permit tab
try:
    if "work_scope" in mapping.get("permit", {}):
        create_template_block(tab_permit, "work_scope", mapping["permit"]["work_scope"].get("label", "На выполнение работ"))
    else:
        create_template_block(tab_permit, "work_scope", "На выполнение работ")
except Exception:
    try:
        create_template_block(tab_permit, "work_scope", "На выполнение работ")
    except Exception:
        pass
for k, meta in mapping["permit"].items():
    if k.startswith("w"):
        continue
    # Пропускаем однострочную графу "hazards" — используем только специализированный многострочный блок ниже
    if k == "hazards":
        continue
    if k in permit_text_keys_with_templates:
        create_template_block(tab_permit, k, meta["label"])
    else:
        ctk.CTkLabel(tab_permit, text=meta["label"]).pack(anchor="w", padx=6, pady=(8,2))
        ent = ctk.CTkEntry(tab_permit, width=80); ent.insert(0, defaults.get(k,"")); ent.pack(fill="x", padx=6, pady=2)
        try:
            ent.bind("<KeyRelease>", lambda e: schedule_autosave())
        except Exception:
            pass
        try:
            ent.bind("<FocusOut>", lambda e: schedule_autosave())
        except Exception:
            pass
        widgets[k] = {"type":"entry","widget":ent}
        #add_mousewheel_support(ent)

# Hazards block: Опасные и вредные факторы — 4 строки с плейсхолдерами и поддержкой шаблонов
ctk.CTkLabel(tab_permit, text="Опасные и вредные факторы").pack(anchor="w", padx=6, pady=(12,2))
frame_tpl = ctk.CTkFrame(tab_permit); frame_tpl.pack(anchor="w", padx=6, pady=(0,2))

hazards_combo_var = tk.StringVar()

def _get_hazard_templates():
    try:
        return templates.get("fields", {}).get("hazards", [])
    except Exception:
        return []

def _hazards_template_names():
    lst = _get_hazard_templates()
    names = []
    for it in lst:
        if isinstance(it, dict):
            name = (it.get("name") or "").strip()
            if not name:
                content = it.get("content","")
                name = content.splitlines()[0] if content else ""
        else:
            s = str(it) if it is not None else ""
            name = s.splitlines()[0] if s.splitlines() else s
        names.append(name)
    return names

hazards_combo = make_combo(frame_tpl, hazards_combo_var, values=_hazards_template_names(), width=58, key="hazards")
hazards_combo.grid(row=0, column=0, padx=(0,4))
hazards_combo.bind('<<ComboboxSelected>>', lambda e: hazards_load())

# Wrapper to adapt 4 single-line entries to the existing Text-like API used by templates logic
class HazardWidget:
    def __init__(self, entries):
        self.entries = entries  # list of CTkEntry
    def get(self, *args, **kwargs):
        # emulate Text.get("1.0","end") => join lines with newline
        lines = [e.get().strip() for e in self.entries]
        return "\n".join(lines)
    def delete(self, *args, **kwargs):
        for e in self.entries:
            try:
                e.delete(0, "end")
            except Exception:
                try:
                    e.delete(0, tk.END)
                except Exception:
                    pass
    def insert(self, *args, **kwargs):
        # insert(index, text) or insert(text)
        if len(args) >= 2:
            text = args[1]
        elif len(args) == 1:
            text = args[0]
        else:
            text = kwargs.get('text', '')
        parts = str(text).splitlines()
        for i in range(len(self.entries)):
            try:
                self.entries[i].delete(0, "end")
            except Exception:
                pass
            if i < len(parts):
                try:
                    self.entries[i].insert(0, parts[i])
                except Exception:
                    pass

# Create 4 single-line entries with placeholders "Фактор 1".."Фактор 4" (без отдельных надписей)
hazard_entries = []
for i in range(4):
    ent = ctk.CTkEntry(tab_permit, width=200, font=DEFAULT_FONT, placeholder_text=f"Фактор {i+1}")
    ent.pack(fill="x", padx=6, pady=(8 if i==0 else 2,4))
    ent.bind("<KeyRelease>", lambda e: schedule_autosave())
    hazard_entries.append(ent)

hazards_widget = HazardWidget(hazard_entries)
widgets["hazards"] = {"type":"text", "widget": hazards_widget}
# Prefill hazards entries from defaults
try:
    hv = defaults.get('hazards', '')
    if hv:
        try:
            hazards_widget.insert(hv)
        except Exception:
            # fallback: populate entries manually
            parts = str(hv).splitlines()
            for i, ent in enumerate(getattr(hazards_widget, 'entries', [])[:4]):
                try:
                    ent.delete(0, 'end')
                    if i < len(parts):
                        ent.insert(0, parts[i])
                except Exception:
                    pass
except Exception:
    pass


def hazards_load(event=None):
    # Load selected hazards template (full content) into the 4 hazard entries.
    v = hazards_combo_var.get().strip()
    if not v:
        return
    # find template object by matching first-line (name) or by content
    lst = templates.get('fields', {}).get('hazards', [])
    chosen_content = None
    for it in lst:
        try:
            if isinstance(it, dict):
                name = (it.get('name') or '').strip()
                content = it.get('content','')
                if not name and content:
                    name = content.splitlines()[0] if content.splitlines() else ''
                if name == v:
                    chosen_content = content; break
                # also allow matching entire content stored as name-less dict
                if content.strip() == v:
                    chosen_content = content; break
            else:
                s = str(it) if it is not None else ''
                nm = s.splitlines()[0] if s.splitlines() else s
                if nm == v:
                    chosen_content = s; break
        except Exception:
            continue
    # fallback: if not found, assume combo contains the actual content (v)
    if chosen_content is None:
        chosen_content = v
    # Split into up to 4 lines
    parts = [line.rstrip() for line in str(chosen_content).splitlines()]
    # Normalize to 4 elements
    while len(parts) < 4:
        parts.append('')
    # Now put values into the widget stored in widgets['hazards']
    wmeta = widgets.get('hazards', {})
    w = wmeta.get('widget')
    if not w:
        return
    try:
        # If it's the HazardWidget wrapper (supports delete/insert like Text)
        # try to call delete/insert as Text-like
        try:
            w.delete('1.0', 'end')
        except Exception:
            try:
                w.delete(0, 'end')
            except Exception:
                pass
        try:
            # insert full content joined by newlines
            w.insert('1.0', '\n'.join(parts))
        except Exception:
            try:
                w.insert('\n'.join(parts))
            except Exception:
                # if it's a list of CTkEntry widgets
                if isinstance(w, list):
                    for i, ent in enumerate(w[:4]):
                        try:
                            ent.delete(0, 'end')
                        except Exception:
                            pass
                        try:
                            ent.insert(0, parts[i])
                        except Exception:
                            pass
                else:
                    # last resort: if w has attribute 'entries' (HazardWidget), populate them
                    try:
                        entries = getattr(w, 'entries', None)
                        if entries:
                            for i, ent in enumerate(entries[:4]):
                                try: ent.delete(0, 'end')
                                except Exception: pass
                                try: ent.insert(0, parts[i])
                                except Exception: pass
                    except Exception:
                        pass
    except Exception:
        # unexpected - try populate as list fallback
        try:
            if isinstance(w, list):
                for i, ent in enumerate(w[:4]):
                    try: ent.delete(0, 'end')
                    except Exception: pass
                    try: ent.insert(0, parts[i])
                    except Exception: pass
        except Exception:
            pass
    schedule_autosave()
def hazards_add():
    # save current hazards entries as a template: name = first non-empty line, content = joined 4 lines
    entries = widgets.get("hazards", {}).get("widget")
    if not entries:
        return
    try:
        # entries is HazardWidget in original code: use its get()
        txt = entries.get("1.0","end").strip()
    except Exception:
        try:
            txt = entries.get().strip()
        except Exception:
            # fallback: join list of entry widgets
            lst = []
            for e in entries:
                try:
                    lst.append(e.get().strip())
                except Exception:
                    lst.append("")
            txt = "\n".join(lst).strip()
    if not txt:
        messagebox.showwarning("Внимание","Нечего сохранять в шаблон.")
        return
    name = txt.splitlines()[0] if txt.splitlines() else txt
    templates.setdefault("fields",{}).setdefault("hazards",[])
    lst = templates["fields"]["hazards"]
    # avoid duplicate by exact content
    exists = False
    for it in lst:
        if isinstance(it, dict):
            if it.get("content","") == txt:
                exists = True; break
        else:
            if str(it) == txt:
                exists = True; break
    if not exists:
        lst.append({"name": name, "content": txt})
        save_json(TEMPLATES_FILE, templates)
    # refresh combo
    try:
        hazards_combo.set_values(_hazards_template_names())
    except Exception:
        refresh_all_template_combos()
    try:
        hazards_combo_var.set("")
    except Exception:
        pass
    schedule_autosave()

def hazards_del():
    v = hazards_combo_var.get().strip()
    if not v:
        messagebox.showwarning("Внимание","Выберите шаблон для удаления")
        return
    lst = templates.get("fields", {}).get("hazards", [])
    newlst = []
    removed = False
    for it in lst:
        if isinstance(it, dict):
            nm = (it.get("name") or "").strip() or (it.get("content","").splitlines()[0] if it.get("content","") else "")
            if nm == v:
                removed = True
                continue
        else:
            s = str(it) if it is not None else ""
            nm = s.splitlines()[0] if s.splitlines() else s
            if nm == v:
                removed = True
                continue
        newlst.append(it)
    if removed:
        templates.setdefault("fields",{})["hazards"] = newlst
        save_json(TEMPLATES_FILE, templates)
        try:
            hazards_combo.set_values(_hazards_template_names())
            hazards_combo_var.set("")
        except Exception:
            refresh_all_template_combos()
    else:
        messagebox.showinfo("Внимание","Шаблон не найден или уже удалён.")
    schedule_autosave()
def hazards_clear():
    try:
        w = widgets.get("hazards", {}).get("widget")
        if w:
            try:
                w.delete("1.0", "end")
            except Exception:
                try:
                    w.delete(0, "end")
                except Exception:
                    pass
    except Exception:
        pass
    schedule_autosave()

make_button(frame_tpl, text="Добавить шаблон", command=hazards_add).grid(row=0, column=2, padx=4)
make_button(frame_tpl, text="Удалить шаблон", command=hazards_del).grid(row=0, column=3, padx=4)
make_button(frame_tpl, text="Очистить", command=hazards_clear).grid(row=0, column=4, padx=6)



# Hazards block (многострочный) — оставляем полностью с логикой шаблонов


def refresh_permit_workers_display():
    # UI-отображение работников во вкладке "Наряд-допуск" удалено — функция оставлена безопасной заглушкой
    return
# TAB: Список
tabview.add("Список")
tab_spisok = tabview.tab("Список")

for k, meta in mapping["spisok"].items():
    # убираем поля position и place из UI (по запросу)
    if k in ("birth","pass","position","place"):
        continue
    ctk.CTkLabel(tab_spisok, text=meta["label"]).pack(anchor="w", padx=6, pady=2)
    if k == "workers":
        try:
            txt = ctk.CTkTextbox(tab_spisok, height=15, font=DEFAULT_FONT)
            txt.insert("1.0", defaults.get("spisok_workers",""))
            txt.pack(fill="both", padx=6, pady=2)
        except Exception:
            txt = scrolledtext.ScrolledText(tab_spisok, height=15, font=DEFAULT_FONT)
            txt.insert("1.0", defaults.get("spisok_workers",""))
            try:
                if ctk.get_appearance_mode().lower() == "dark":
                    try:
                        tmp = ctk.CTkFrame(tab_spisok)
                        bg = tmp.cget("fg_color"); tmp.destroy()
                    except Exception:
                        bg = "#2b2b2b"
                    fg = "#e6e6e6"
                    try:
                        txt.configure(bg=bg, fg=fg, insertbackground=fg, selectbackground="#4b4b4b")
                    except Exception:
                        pass
            except Exception:
                pass
            txt.pack(fill="both", padx=6, pady=2)
        def on_spisok_modified(e):
            try:
                e.widget.edit_modified(False)
            except Exception:
                pass
            schedule_autosave(); refresh_permit_workers_display()
        try:
            txt.bind("<<Modified>>", on_spisok_modified)
        except Exception:
            try:
                txt.bind("<KeyRelease>", lambda e: on_spisok_modified(e))
            except Exception:
                pass
        widgets["spisok_workers"] = {"type":"text","widget":txt}
        #add_mousewheel_support(txt)

        ctrl_frame = ctk.CTkFrame(tab_spisok); ctrl_frame.pack(anchor="w", padx=6, pady=4)
        def open_worker_selector_multi():
            sel = make_ctk_toplevel(root, "Выбрать работника(ов)", "720x520")
            try:
                # try to set Toplevel background to CTk frame background for consistent dark theme
                bg = None
                try:
                    tmp = ctk.CTkFrame(sel)
                    bg = tmp.cget("fg_color")
                    tmp.destroy()
                except Exception:
                    bg = None
                if bg:
                    try:
                        sel.configure(bg=bg)
                    except Exception:
                        pass
            except Exception:
                pass
            sel_frame_top = ctk.CTkFrame(sel); sel_frame_top.pack(fill="x", padx=8, pady=6)
            ctk.CTkLabel(sel_frame_top, text="Поиск / фильтр:").pack(side="left", padx=(0,6))
            search_var = tk.StringVar()
            search_entry = ctk.CTkEntry(sel_frame_top, textvariable=search_var, width=40)
            search_entry.pack(side="left", padx=(0,6))

            # Use CTkScrollableFrame with CheckBoxes for multi-select
            sel_container = ctk.CTkScrollableFrame(sel)
            sel_container.pack(fill="both", expand=True, padx=6, pady=6)
            sel_row_vars = []
            sel_row_frames = []
            sel_row_names = []

            def repopulate(filter_text=""):
                # clear existing rows
                for fr in list(sel_row_frames):
                    try: fr.destroy()
                    except Exception: pass
                sel_row_frames.clear(); sel_row_vars.clear(); sel_row_names.clear()
                f = (filter_text or "").strip().lower()
                for w in workers_db:
                    fio = w.get("fio","")
                    if not f or f in fio.lower():
                        row = ctk.CTkFrame(sel_container)
                        row.pack(fill="x", padx=4, pady=2)
                        var = tk.BooleanVar(value=False)
                        chk = ctk.CTkCheckBox(row, text=fio, variable=var)
                        chk.pack(fill="x", side="left", expand=True, padx=(6,2), pady=6)
                        # clicking row selects it (visual highlight)
                        def _make_onclick(idx):
                            return lambda e=None: None
                        try:
                            row.bind("<Button-1>", _make_onclick(len(sel_row_frames)))
                            chk.bind("<Button-1>", _make_onclick(len(sel_row_frames)))
                        except Exception:
                            pass
                        sel_row_frames.append(row)
                        sel_row_vars.append(var)
                        sel_row_names.append(fio)

            def on_search_change(*_):
                repopulate(search_var.get())
            search_var.trace_add("write", on_search_change)
            repopulate()





            def do_add_selected():
                picks = [name for i,name in enumerate(sel_row_names) if sel_row_vars[i].get()]
                if not picks:
                    messagebox.showwarning("Внимание","Выберите хотя бы одного работника")
                    return
                s = widgets["spisok_workers"]["widget"]
                existing = [ln.strip() for ln in s.get("1.0","end").splitlines() if ln.strip()]
                to_add = []
                for name in picks:
                    if name not in existing and name not in to_add:
                        to_add.append(name)
                if to_add:
                    if existing:
                        s.insert("end", "\n" + "\n".join(to_add))
                    else:
                        s.insert("1.0", "\n".join(to_add))
                    schedule_autosave()
                    refresh_permit_workers_display()
                sel.destroy()

            def do_add_all():
                all_names = sel_row_names[:]
                if not all_names:
                    messagebox.showwarning("Внимание","Список работников пуст")
                    return
                s = widgets["spisok_workers"]["widget"]
                existing = [ln.strip() for ln in s.get("1.0","end").splitlines() if ln.strip()]
                to_add = []
                for name in all_names:
                    if name not in existing and name not in to_add:
                        to_add.append(name)
                if not to_add:
                    messagebox.showinfo("Внимание", "Нечего добавлять — все работники уже в списке.")
                    return
                if existing:
                    s.insert("end", "\n" + "\n".join(to_add))
                else:
                    s.insert("1.0", "\n".join(to_add))
                schedule_autosave()
                refresh_permit_workers_display()
                sel.destroy()

            # Buttons
            btns = ctk.CTkFrame(sel)
            btns.pack(side="bottom", fill="x", padx=8, pady=(6,8))
            make_button(btns, text="Добавить выбранных", command=lambda: _do_add_selected_inner(sel, sel_row_names, sel_row_vars), width=18).pack(side="left", padx=8)
            make_button(btns, text="Добавить всех", command=lambda: _do_add_all_inner(sel, sel_row_names), width=18).pack(side="left", padx=8)
            make_button(btns, text="Отмена", command=sel.destroy, width=18).pack(side="right", padx=8)




            def do_add_all():
                all_names = [w.get('fio','') for w in workers_db]
                if not all_names:
                    messagebox.showwarning("Внимание","Список работников пуст")
                    return
                s = widgets["spisok_workers"]["widget"]
                existing = [ln.strip() for ln in s.get("1.0","end").splitlines() if ln.strip()]
                to_add = []
                for name in all_names:
                    if name not in existing and name not in to_add:
                        to_add.append(name)
                if not to_add:
                    messagebox.showinfo("Внимание", "Нечего добавлять — все работники уже в списке.")
                    return
                if existing:
                    s.insert("end", "\n" + "\n".join(to_add))
                else:
                    s.insert("1.0", "\n".join(to_add))
                schedule_autosave()
                refresh_permit_workers_display()
                sel.destroy()

            """
            Диалог сборки бригады: multi-select работников + ввод названия бригады.
            Кнопка 'Сохранить бригаду' сохраняет шаблон бригады и вставляет её в spisok.
            """

        make_button(ctrl_frame, text="Добавить работника", command=open_worker_selector_multi, width=BUTTON_WIDTH_LARGE, font=BUTTON_FONT_LARGE).pack(side="left", padx=6)
        make_button(ctrl_frame, text="Очистить список работников", command=lambda:(widgets["spisok_workers"]["widget"].delete("1.0","end"), schedule_autosave(), refresh_permit_workers_display()), width=BUTTON_WIDTH_LARGE, font=BUTTON_FONT_LARGE).pack(side="left", padx=6)
    else:
        ent = ctk.CTkEntry(tab_spisok, width=80)
        ent.insert(0, defaults.get(k,""))
        ent.pack(fill="x", padx=6, pady=2)
        try:

            ent.bind("<KeyRelease>", lambda e: schedule_autosave())
        except Exception:
            pass
        try:
            ent.bind("<FocusOut>", lambda e: schedule_autosave())
        except Exception:
            pass
        widgets[k] = {"type":"entry","widget":ent}
        #add_mousewheel_support(ent)




# --- Автосохранение: рекурсивно привязываем события сохранения ко всем полям во вкладке "Наряд-допуск" ---
def bind_autosave_for_permit_children():
    try:
        def _safe_bind_text(w):
            try:
                # для tk.Text и ScrolledText используем <<Modified>>
                w.bind("<<Modified>>", lambda e, ww=w: (ww.edit_modified(False), schedule_autosave()))
            except Exception:
                try:
                    w.bind("<KeyRelease>", lambda e, ww=w: schedule_autosave())
                except Exception:
                    pass
        def _safe_bind_entry(w):
            try:
                w.bind("<KeyRelease>", lambda e, ww=w: schedule_autosave())
            except Exception:
                try:
                    w.bind("<<Modified>>", lambda e, ww=w: schedule_autosave())
                except Exception:
                    pass
        def recurse(widget):
            try:
                children = widget.winfo_children()
            except Exception:
                return
            for ch in children:
                try:
                    # текстовые виджеты TK / ScrolledText
                    if isinstance(ch, tk.Text) or (hasattr(scrolledtext, 'ScrolledText') and isinstance(ch, scrolledtext.ScrolledText)):
                        _safe_bind_text(ch)
                    else:
                        # CTkEntry / CTkTextbox / tk.Entry
                        name = ch.__class__.__name__
                        if name in ("CTkEntry", "CTkTextbox") or isinstance(ch, tk.Entry):
                            _safe_bind_entry(ch)
                    # рекурсивно пройти глубже
                    recurse(ch)
                except Exception:
                    pass
        # Запускаем привязку от корня таба permit (если он существует)
        try:
            if 'tab_permit' in globals() and tab_permit is not None:
                recurse(tab_permit)
        except Exception:
            pass
    except Exception:
        pass

# Вызовим привязку (гарантируем, что все поля во вкладке теперь будут автосохраняться)
try:
    bind_autosave_for_permit_children()
except Exception:
    pass

# --- Конец вставки автосохранения ---

# TAB: Работники

tabview.add("Работники")
tab_workers = tabview.tab("Работники")
# CTk-based scrollable container replacing the old tk.Listbox for workers
workers_container = ctk.CTkScrollableFrame(tab_workers)
workers_container.pack(fill="both", expand=True, padx=6, pady=6)
# internal state for rows
workers_row_vars = []
workers_row_frames = []
workers_selected_idx = tk.IntVar(value=-1)

def _on_row_click(idx):
    try:
        workers_selected_idx.set(idx)
    except Exception:
        pass

def get_selected_worker_indices():
    # return list of indices selected via checkboxes; if none, fallback to single selected index
    sel = [i for i,var in enumerate(workers_row_vars) if var.get()]
    if sel:
        return sel
    s = workers_selected_idx.get()
    if s is not None and s >= 0 and s < len(workers_db):
        return [s]
    return []

def refresh_workers_listbox():
    # repopulate CTk rows
    for fr in list(workers_row_frames):
        try:
            fr.destroy()
        except Exception:
            pass
    workers_row_frames.clear()
    workers_row_vars.clear()
    # create a frame per worker with a CheckBox and clickable label area
    for i, w in enumerate(workers_db):
        fio = w.get("fio","")
        row = ctk.CTkFrame(workers_container)
        row.pack(fill="x", padx=4, pady=2)
        var = tk.BooleanVar(value=False)
        chk = ctk.CTkCheckBox(row, text=fio, variable=var)
        chk.pack(fill="x", side="left", expand=True, padx=(6,2), pady=6)
        # clicking the checkbox will toggle selection; but allow clicking row to set active index
        def _make_onclick(idx):
            return lambda e=None: _on_row_click(idx)
        # bind left click on the row to select it
        try:
            row.bind("<Button-1>", _make_onclick(i))
            chk.bind("<Button-1>", _make_onclick(i))
        except Exception:
            pass
        workers_row_frames.append(row)
        workers_row_vars.append(var)

# initial fill
refresh_workers_listbox()

def save_workers_db():
    save_json(WORKERS_FILE, workers_db); refresh_workers_listbox()
def open_worker_card(existing=None, index=None):
    # Create centered modal window scaled to screen size so fields/buttons are visible
    # Compute initial geometry to avoid brief flicker at default position. We'll refine height after layout.
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    w = min(1000, int(sw * 0.6))
    init_h = 480
    x = (sw - w) // 2
    y = max(0, (sh - init_h) // 2)
    win = make_ctk_toplevel(root, "Карточка работника", geometry=f"{w}x{init_h}+{x}+{y}")

    # fields: labels and entries made 2-2.5x larger (bigger width and larger font)
    fields = [
        ("fio", "ФИО (полностью)"),
        ("position", "Должность"),
        ("birth", "Дата рождения (дд.мм.гггг)"),
        ("pass", "Серия и номер"),
        ("place", "Кем выдан"),
        ("notes", "Примечание")
    ]
    entries = {}
    # choose larger font sizes
    label_font = ("TkDefaultFont", 14)
    entry_font = ("TkDefaultFont", 14)
    # make entries 25% larger than before (previously 200)
    entry_width = 250

    # Use a frame with some padding and vertical layout so controls scale nicely
    content = ctk.CTkFrame(win)
    content.pack(fill="both", expand=True, padx=12, pady=12)
    for key, label in fields:
        ctk.CTkLabel(content, text=label, font=label_font).pack(anchor="w", padx=6, pady=(8,2))
        ent = ctk.CTkEntry(content, width=entry_width, font=entry_font)
        ent.pack(anchor="w", padx=6, pady=(0,6))
        entries[key] = ent

    if existing:
        for k in entries:
            entries[k].delete(0, "end")
            entries[k].insert(0, existing.get(k, ""))

    def do_save():
        obj = {k: entries[k].get().strip() for k in entries}
        if not obj["fio"]:
            messagebox.showwarning("Внимание", "ФИО обязательно")
            return
        if existing is not None and index is not None:
            workers_db[index] = obj
        else:
            workers_db.append(obj)
        save_workers_db()
        refresh_permit_workers_display()
        win.grab_release()
        win.destroy()

    # Button frame pinned to bottom so it remains visible
    btn_frame = ctk.CTkFrame(win)
    btn_frame.pack(fill="x", side="bottom", pady=10, padx=10)
    make_button(btn_frame, text="Сохранить", command=do_save, width=20).pack(side="left", padx=8, pady=6)
    make_button(btn_frame, text="Отмена", command=lambda: (win.grab_release(), win.destroy()), width=14).pack(side="left", padx=8, pady=6)

    # Now that content is packed, compute required height and set geometry so buttons are visible without manual resize
    win.update_idletasks()
    req_h = win.winfo_reqheight()
    # ensure height does not exceed 90% of screen height and is at least 480
    h = max(480, min(int(sh * 0.9), req_h))
    y = max(0, (sh - h) // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")
    win.minsize(520, 420)
    win.transient(root)
    win.grab_set()

def new_worker(): open_worker_card()
def edit_worker():
    try: idx = (lambda s=get_selected_worker_indices(): s[0] if s else None)()
    except Exception: messagebox.showwarning("Внимание","Выберите работника для редактирования"); return
    open_worker_card(existing=workers_db[idx], index=idx)
def delete_worker():
    try: idx = (lambda s=get_selected_worker_indices(): s[0] if s else None)()
    except Exception: messagebox.showwarning("Внимание","Выберите работника для удаления"); return
    if messagebox.askyesno("Подтвердите","Удалить выбранного работника?"):
        workers_db.pop(idx); save_workers_db(); refresh_permit_workers_display()

def add_selected_to_spisok():
    idxs = get_selected_worker_indices()
    if not idxs:
        messagebox.showwarning("Внимание","Выберите работника(ов)")
        return
    picks = [workers_db[i].get("fio","") for i in idxs]
    s = widgets.get("spisok_workers",{}).get("widget")
    if s:
        existing = [ln.strip() for ln in s.get("1.0","end").splitlines() if ln.strip()]
        to_add = []
        for name in picks:
            if name not in existing and name not in to_add:
                to_add.append(name)
        if to_add:
            if existing:
                s.insert("end", "\n" + "\n".join(to_add))
            else:
                s.insert("1.0", "\n".join(to_add))
            schedule_autosave(); refresh_permit_workers_display()
        else:
            messagebox.showinfo("Внимание","Выбранные работники уже присутствуют в списке")
    else:
        messagebox.showinfo("Внимание","Поле 'Список' не найдено")

wb_frame = ctk.CTkFrame(tab_workers); wb_frame.pack(anchor="w", padx=6, pady=(4,120))
make_button(wb_frame, text="Новый", command=new_worker, width=18).pack(side="left", padx=6)
make_button(wb_frame, text="Редактировать", command=edit_worker, width=18).pack(side="left", padx=6)
make_button(wb_frame, text="Удалить", command=delete_worker, width=18).pack(side="left", padx=6)
make_button(wb_frame, text="Добавить в Список", command=add_selected_to_spisok, width=BUTTON_WIDTH_LARGE).pack(side="left", padx=6)

# Save profile / build context / generation
def save_profile(show_msg=True):
    new = {}
    fio_lines = fio_txt.get("1.0","end").strip().splitlines()
    new["fio"] = fio_lines[0] if len(fio_lines)>0 else ""
    new["fio2"] = fio_lines[1] if len(fio_lines)>1 else ""
    new["fio4"] = fio_lines[2] if len(fio_lines)>2 else ""
    if widgets.get("fio1"): new["fio1"] = widgets["fio1"]["widget"].get().strip()
    if widgets.get("fio3"): new["fio3"] = widgets["fio3"]["widget"].get().strip()
    for k,meta in mapping["shared"].items():
        if k in ("fio","fio2","fio4","fio1","fio3","aa","bb","dd","ee","numb"): continue
        if widgets.get(k):
            val = widgets[k]["widget"].get().strip()
            new[k] = val
    # Save spisok fields (including 'predmet')
    for k in mapping["spisok"].keys():
        if widgets.get(k):
            if widgets[k]["type"] == "text":
                w = widgets[k]["widget"]
                try:
                    if getattr(w, "_placeholder_active", False):
                        new[k] = ""
                    else:
                        new[k] = w.get("1.0","end").strip()
                except Exception:
                    new[k] = ""
            else:
                new[k] = widgets[k]["widget"].get().strip()
    
    # Сохраняем поля вкладки 'Наряд-допуск' (исключая w0..w11)
    try:
        for k, meta in mapping["permit"].items():
            if k.startswith("w"):
                continue
            if widgets.get(k):
                try:
                    if widgets[k]["type"] == "text":
                        new[k] = widgets[k]["widget"].get("1.0", "end").strip()
                    else:
                        new[k] = widgets[k]["widget"].get().strip()
                except Exception:
                    try:
                        # fallback for composite controls
                        w = widgets[k]["widget"]
                        if hasattr(w, "get"):
                            new[k] = w.get().strip()
                    except Exception:
                        new[k] = ""
    except Exception:
        pass
    save_json(DEFAULTS_FILE, new)
    save_json(WORKERS_FILE, workers_db)
    save_json(TEMPLATES_FILE, templates)
    save_brigades_db()
    if show_msg: messagebox.showinfo("OK","Профиль сохранён")

def build_ctx_common():
    ctx = {}
    fio_lines = fio_txt.get("1.0","end").strip().splitlines()
    ctx["fio"] = fio_lines[0] if len(fio_lines)>0 else ""
    ctx["fio2"] = fio_lines[1] if len(fio_lines)>1 else ""
    ctx["fio4"] = fio_lines[2] if len(fio_lines)>2 else ""
    if widgets.get("fio1"): ctx["fio1"] = widgets["fio1"]["widget"].get().strip()
    if widgets.get("fio3"): ctx["fio3"] = widgets["fio3"]["widget"].get().strip()
    for k,meta in mapping["shared"].items():
        if k in ("fio","fio2","fio4","fio1","fio3","aa","bb","dd","ee","numb"): continue
        if widgets.get(k): ctx[k] = widgets[k]["widget"].get().strip()
    for dkey in ("a","b","d","e"):
        raw = widgets.get(dkey,{}).get("widget").get().strip() if widgets.get(dkey) else ""
        day, month_year = parse_ddmmyyyy(raw)
        ctx[dkey] = day
        if dkey=="a": ctx["aa"] = month_year
        if dkey=="b": ctx["bb"] = month_year
        if dkey=="d": ctx["dd"] = month_year
        if dkey=="e": ctx["ee"] = month_year

    # ДОБАВЛЯЕМ permit-поля (не включая w0..w11)
    for k,meta in mapping["permit"].items():
        if k.startswith("w"):
            continue
        if widgets.get(k):
            if widgets[k]["type"] == "text":
                ctx[k] = widgets[k]["widget"].get("1.0", "end").strip()
            else:
                ctx[k] = widgets[k]["widget"].get().strip()

    # Разбиваем hazards на hazards1..hazards4
    hz_text = ctx.get("hazards","")
    hz_lines = hz_text.splitlines()
    for i in range(4):
        ctx[f"hazards{i+1}"] = hz_lines[i].strip() if i < len(hz_lines) else ""
        # --- FORCE FIX: ensure work_scope is taken from widgets ---
        try:
            w = widgets.get('work_scope')
            if w:
                if w.get('type') == 'text':
                    ctx['work_scope'] = w['widget'].get('1.0', 'end').strip()
                else:
                    ctx['work_scope'] = w['widget'].get().strip()
            # debug print
            try:
                print('DEBUG_FORCE: ctx["work_scope"] =', repr(ctx.get('work_scope','')))
            except Exception:
                pass
        except Exception:
            pass
        # --- END FORCE FIX ---


    return ctx


def build_ctx_spisok():
    """
    Построение контекста для шаблона 'список':
    - ctx['workers'] = list(dict...) — для таблиц Jinja
    - дополнительные переменные: worker, worker1, worker2...,
      а также workerN_position, workerN_birth, workerN_pass, workerN_place, workerN_notes
    - + переменные удобного доступа: position, position1..position11,
      birth, birth1..birth11, pass, pass1..pass11, place, place1..place11
    Кроме того — перенос всех полей mapping['spisok'] (например 'predmet') в контекст,
    а также вложенный словарь ctx['spisok'] для совместимости с шаблонами.
    """
    ctx = build_ctx_common()

    # Сначала скопируем UI-поля spisok (например 'predmet')
    spisok_fields = {}
    for k in mapping["spisok"].keys():
        if widgets.get(k):
            if widgets[k]["type"] == "text":
                val = widgets[k]["widget"].get("1.0","end").strip()
            else:
                val = widgets[k]["widget"].get().strip()
        else:
            # если поле отсутствует в UI — берем значение из defaults если есть, иначе пустую строку
            val = defaults.get(k, "")
        # сохраняем как в ctx по имени, так и в отдельном словаре spisok_fields
        ctx[k] = val
        spisok_fields[k] = val

    # Также добавим вложенный словарь для шаблонов, которые обращаются как {{ spisok.predmet }}
    ctx["spisok"] = spisok_fields

    sp = widgets.get("spisok_workers",{}).get("widget").get("1.0","end").strip() if widgets.get("spisok_workers") else ""
    workers = []
    for ln in [l.strip() for l in sp.splitlines() if l.strip()]:
        found = next((w for w in workers_db if w.get("fio","") == ln), None)
        if found:
            workers.append({
                "fio": found.get("fio",""),
                "position": found.get("position",""),
                "birth": found.get("birth",""),
                "pass": found.get("pass",""),
                "place": found.get("place",""),
                "notes": found.get("notes","")
            })
        else:
            workers.append({"fio": ln})

    # основной список (для Jinja-таблиц типа {% for w in workers %} ...)
    ctx["workers"] = workers

    # доп. переменные для шаблонов, использующих {worker}, {worker1} и т.д.
    for i, w in enumerate(workers):
        fio = w.get("fio", "")
        if i == 0:
            ctx["worker"] = fio
            ctx["worker0"] = fio
            prefix = "worker"
        else:
            ctx[f"worker{i}"] = fio
            prefix = f"worker{i}"
        ctx[f"{prefix}_position"] = w.get("position", "")
        ctx[f"{prefix}_birth"] = w.get("birth", "")
        ctx[f"{prefix}_pass"] = w.get("pass", "")
        ctx[f"{prefix}_place"] = w.get("place", "")
        ctx[f"{prefix}_notes"] = w.get("notes", "")

    # --- Новые удобные переменные: position, position1..position11 и т.д. ---
    # Для совместимости: первая (0-я) позиция -- без индекса; далее 1..11
    max_slots = 12  # создаём position .. position11 (всего 12 слотов)
    for i in range(max_slots):
        if i < len(workers):
            p = workers[i].get("position", "")
            b = workers[i].get("birth", "")
            pa = workers[i].get("pass", "")
            pl = workers[i].get("place", "")
        else:
            p = b = pa = pl = ""
        if i == 0:
            ctx["position"] = p
            ctx["birth"] = b
            ctx["pass"] = pa
            ctx["place"] = pl
        else:
            ctx[f"position{i}"] = p
            ctx[f"birth{i}"] = b
            ctx[f"pass{i}"] = pa
            ctx[f"place{i}"] = pl

    return ctx





def edit_numb_dialog():
    d = make_ctk_toplevel(root, "Редактировать номер")
    # title set by helper документа")
    d.transient(root)
    d.grab_set()
    # center and size so buttons are visible immediately
    d.update_idletasks()
    sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
    w = 600; h = 340
    x = (sw - w)//2; y = (sh - h)//2
    d.geometry(f"{w}x{h}+{x}+{y}")
    d.minsize(520, 320)
    d.resizable(False, False)

    # content frame holds the labels/entries and expands
    content = ctk.CTkFrame(d)
    content.pack(fill="both", expand=True, padx=12, pady=(12,6))

    ctk.CTkLabel(content, text="Введите номер (число):").pack(anchor="w", padx=6, pady=(4,2))
    num_ent = ctk.CTkEntry(content, width=260, font=DEFAULT_FONT)
    num_ent.insert(0, str(counters.get("numb", 1606)))
    num_ent.pack(anchor="w", padx=6, pady=(0,8), fill="x")

    ctk.CTkLabel(content, text="Введите буквенное обозначение (опционально):").pack(anchor="w", padx=6, pady=(6,2))
    suff_ent = ctk.CTkEntry(content, width=120, font=DEFAULT_FONT)
    suff_ent.insert(0, str(counters.get("numb_suffix", "")))
    suff_ent.pack(anchor="w", padx=6, pady=(0,8), fill="x")

    # Put buttons in bottom bar so they are always visible; styled as a frame (grey bar)
    btnf = ctk.CTkFrame(d)
    btnf.pack(fill="x", side="bottom", padx=0, pady=0)
    inner = ctk.CTkFrame(btnf)
    inner.pack(fill="x", padx=12, pady=12)
    # Left-aligned Save/Cancel with generous padding so buttons are fully visible
    make_button(inner, text="Сохранить", command=lambda: do_save_numb(num_ent, suff_ent, d), width=14).pack(side="left", padx=12, pady=8)
    make_button(inner, text="Отмена", command=lambda: (d.grab_release(), d.destroy()), width=14).pack(side="left", padx=12, pady=8)

    # Ensure focus is on the number entry
    try:
        num_ent.focus_set()
    except Exception:
        pass


def do_save_numb(num_ent, suff_ent, dlg):
    try:
        v = int(num_ent.get().strip())
    except Exception:
        messagebox.showwarning("Ошибка", "Номер должен быть целым числом")
        return
    suff = suff_ent.get().strip()
    counters["numb"] = v
    counters["numb_suffix"] = suff
    save_json(COUNTERS_FILE, counters)
    # update main display (ent_num may be entry or label variable)
    try:
        if 'ent_num_var' in globals():
            ent_num_var.set(str(counters.get("numb")) + (("-"+suff) if suff else ""))
        else:
            ent_num.configure(state="normal")
            ent_num.delete(0, "end")
            ent_num.insert(0, str(counters.get("numb")) + (("-"+suff) if suff else ""))
            ent_num.configure(state="readonly")
    except Exception:
        pass
    try:
        dlg.grab_release()
        dlg.destroy()
    except Exception:
        pass

def generate_docx_all():
    save_profile(False)
    outs = []
    errors = []
    files_map = {"permit":TEMPLATE_PERMIT,"spisok":TEMPLATE_SPISOK,"order":TEMPLATE_ORDER,"pb_order":TEMPLATE_PB_ORDER}

    # резервируем текущий номер, но не инкрементируем ещё в файле.
    reserved_numb = counters.get("numb", 1606)

    for key, path in files_map.items():
        if not path.exists(): continue

        if key == "spisok":
            ctx = build_ctx_spisok()
        else:
            ctx = build_ctx_common()

        # всем генерируемым файлам даём один и тот же номер
        ctx["numb"] = str(reserved_numb) + (("-" + str(counters.get("numb_suffix",""))) if counters.get("numb_suffix","") else "")

        if key == "permit":
            sp_ctx = build_ctx_spisok()
            sp_workers = sp_ctx.get("workers", [])
            short_list = [short_name(w.get("fio","")) for w in sp_workers]
            max_slots = 12  # w0..w11
            for i in range(max_slots):
                ctx[f"w{i}"] = short_list[i] if i < len(short_list) else ""
            ctx["w"] = ctx.get("w0", "")

        # human-readable names for final files
        human_names = {"permit":"наряд-допуск", "spisok":"список", "order":"приказ", "pb_order":"приказ-пб"}
        human = human_names.get(key, path.stem)
        
        # use number already placed into ctx (falls back to reserved_numb if missing)
        safe_numb = ctx.get("numb", str(reserved_numb))
        
        # build final filename like: "наряд-допуск (1606-А).docx"
        out_name = f"{human} ({safe_numb}).docx"
        
        # sanitize filename (replace forbidden chars)
        out_name = re.sub(r'[\\/:*?"<>|]', '_', out_name)
        
        out_full = get_output_dir() / out_name
        try:
            render_docx_safely(path, ctx, str(out_full))
            outs.append(str(out_full))
        except Exception as e:
            tb = traceback.format_exc()
            errors.append((str(path), str(e), tb))

    # Если создан хотя бы один файл — увеличиваем счётчик на +1 и сохраняем изменения.
    if outs:
        counters["numb"] = reserved_numb + 1
        save_json(COUNTERS_FILE, counters)
        try:
            ent_num.configure(state="normal")
            ent_num.delete(0,"end")
            ent_num.insert(0, str(counters.get("numb", reserved_numb + 1)) + (("-"+str(counters.get("numb_suffix",""))) if counters.get("numb_suffix","") else ""))
            ent_num.configure(state="readonly")
        except Exception:
            pass

        messagebox.showinfo("OK", "Созданы: " + ", ".join(outs))

    if errors:
        short_msgs = []
        for p, msg, tb in errors:
            short_msgs.append(f"{Path(p).name}: {msg}")
        errmsg = "Некоторые шаблоны не сгенерированы:\n" + "\n".join(short_msgs)
        log_path = APPDIR / f"docgen_error_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        with open(log_path, "w", encoding="utf-8") as lf:
            for p, msg, tb in errors:
                lf.write(f"=== {p} ===\n{msg}\n{tb}\n\n")
        errmsg += f"\nПодробный лог: {log_path}"
        messagebox.showerror("Ошибки генерации", errmsg)

# --- Output folder selection utilities ---
def get_output_dir():
    od = settings.get("output_dir")
    if not od:
        od = str(APPDIR)
        settings["output_dir"] = od
        save_json(SETTINGS_FILE, settings)
    return Path(od)

def choose_output_folder():
    current = get_output_dir()
    sel = filedialog.askdirectory(initialdir=str(current), title="Выберите папку для сохранения документов")
    if sel:
        settings["output_dir"] = sel
        save_json(SETTINGS_FILE, settings)
        update_output_dir_label()

# -------------------------
# Открытие исходных шаблонов (новая функциональность)
# -------------------------
def _open_path_with_default_app(pth: Path):
    """Открыть файл в системном приложении (Windows/Mac/Linux)."""
    try:
        # Windows
        if hasattr(os, "startfile"):
            os.startfile(str(pth))
            return
    except Exception:
        pass
    try:
        # macOS
        import subprocess, sys as _sys
        import platform
        if _sys.platform == "darwin":
            subprocess.call(["open", str(pth)])
            return
        # Linux (xdg-open)
        subprocess.call(["xdg-open", str(pth)])
    except Exception as e:
        try:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\\n{pth}\\n\\n{e}")
        except Exception:
            print(f"Failed to open {pth}: {e}")

def open_template_by_key(key):
    """Открывает соответствующий шаблон .docx:
      'наряд-допуск' -> TEMPLATE_PERMIT
      'список'       -> TEMPLATE_SPISOK
      'приказ'       -> TEMPLATE_ORDER
      'приказ-пб'    -> TEMPLATE_PB_ORDER
    """
    mapping = {
        "наряд-допуск": TEMPLATE_PERMIT,
        "список": TEMPLATE_SPISOK,
        "приказ": TEMPLATE_ORDER,
        "приказ-пб": TEMPLATE_PB_ORDER
    }
    path = mapping.get(key)
    if not path:
        try:
            messagebox.showerror("Ошибка", f"Неизвестный тип: {key}")
            return
        except Exception:
            print(f"Unknown key: {key}")
    if not path.exists():
        try:
            messagebox.showwarning("Файл не найден", f"Шаблон не найден:\\n{path}\\n\\nУбедитесь, что файл существует в каталоге приложения.")
            return
        except Exception:
            print(f"Template not found: {path}")
    try:
        _open_path_with_default_app(path)
    except Exception as e:
        try:
            messagebox.showerror("Ошибка", f"Не удалось открыть {path}:\\n{e}")
        except Exception:
            print(f"Failed to open {path}: {e}")

def open_source_selector():
    """Диалог выбора исходника — список из 4 элементов, каждый открывает свой .docx."""
    dlg = make_ctk_toplevel(root, "Изменить исходник")
    # title set by helper
    dlg.geometry("420x260")
    dlg.transient(root)
    ctk.CTkLabel(dlg, text="Выберите исходник и нажмите «Открыть»:").pack(anchor="w", padx=10, pady=(10,6))
    lb = tk.Listbox(dlg, font=DEFAULT_FONT, height=6, exportselection=False)
    options = ["наряд-допуск", "список", "приказ", "приказ-пб"]
    for opt in options:
        lb.insert("end", opt)
    lb.pack(fill="both", expand=True, padx=10, pady=(0,8))

    def do_open(_evt=None):
        try:
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("Внимание", "Выберите пункт списка")
                return
            key = lb.get(sel[0])
            open_template_by_key(key)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # двойной клик = открыть
    lb.bind("<Double-1>", do_open)

    btns = ctk.CTkFrame(dlg)
    btns.pack(fill="x", padx=10, pady=(0,10))
    make_button(btns, text="Открыть", command=do_open, width=20).pack(side="left", padx=(0,6))
    make_button(btns, text="Отмена", command=dlg.destroy, width=20).pack(side="right")



def update_output_dir_label():
    lbl_text = str(get_output_dir())
    p = Path(lbl_text)
    parts = p.parts
    if len(parts) > 4:
        short = os.path.join("...", *parts[-3:])
    else:
        short = lbl_text
    try:
        output_dir_label.configure(text=short)
    except Exception:
        pass
    try:
        output_dir_label_overlay.configure(text=short)
    except Exception:
        pass


# bottom controls (кнопки подогнаны шириной, чтобы текст помещался)
frame_bot = ctk.CTkFrame(root); frame_bot.pack(side='bottom', fill="x", padx=0, pady=6)

# Ensure bottom control frame is above other widgets
try:
    frame_bot.lift()
except Exception:
    pass

# Theme switch placed in bottom controls (right side)
def _toggle_theme_btn():
    """Toggle appearance mode between Light and Dark and persist in settings.json"""
    try:
        current = ctk.get_appearance_mode()
        new = 'Dark' if current.lower() != 'dark' else 'Light'
        ctk.set_appearance_mode(new)
        settings['appearance_mode'] = new
        save_json(SETTINGS_FILE, settings)
    except Exception:
        pass
try:
    theme_var = tk.BooleanVar(value=(ctk.get_appearance_mode().lower()=='dark'))
    theme_switch = ctk.CTkSwitch(frame_bot, text='Тёмная тема', command=_toggle_theme_btn, variable=theme_var)
    theme_switch.pack(side='right', padx=6)
except Exception:
    pass

# Make sure bottom frame stays on top after geometry changes
def _keep_bottom_on_top(event=None):
    try:
        frame_bot.lift()
    except Exception:
        pass

root.bind('<Configure>', _keep_bottom_on_top)

# folder choose button + label
folder_frame = ctk.CTkFrame(frame_bot); folder_frame.pack(side="left", padx=(0,8))
make_button(folder_frame, text="📁", command=choose_output_folder, width=4).pack(side="left")
output_dir_label = ctk.CTkLabel(folder_frame, text=""); output_dir_label.pack(side="left", padx=(6,0))
update_output_dir_label()

ctk.CTkCheckBox(frame_bot, text="Автосохранение", variable=autosave_var).pack(side="left")

# <-- новая кнопка для открытия исходников
make_button(frame_bot, text="Изменить исходник", command=open_source_selector, width=24).pack(side="left", padx=6)

make_button(frame_bot, text="Сгенерировать Все", command=generate_docx_all, width=24).pack(side="left", padx=6)
make_button(frame_bot, text="Сохранить профиль", command=lambda: save_profile(True), width=20).pack(side="left", padx=6)
make_button(frame_bot, text="Выйти", command=lambda:(autosave_now(), root.destroy())).pack(side="right")


# folder choose button + label

def clear_all_sel():
    for info in widgets.values():
        w = info["widget"]
        try:
            if isinstance(w, (tk.Text, scrolledtext.ScrolledText)): w.tag_remove("sel","1.0","end")
            else: w.select_clear()
        except Exception:
            pass
clear_all_sel()

# -------------------------
# Автосохранение: привязать ко всем полям дополнительно,
# на случай если где-то привязки не были сделаны ранее.
# -------------------------
for info in widgets.values():
    w = info["widget"]
    try:
        if isinstance(w, (tk.Text, scrolledtext.ScrolledText)):
            # use default argument to capture current widget
            w.bind("<<Modified>>", lambda e, widget=w: (widget.edit_modified(False), schedule_autosave()))
        else:
            w.bind("<KeyRelease>", lambda e: schedule_autosave())
    except Exception:
        pass

# initial refreshes
refresh_workers_listbox()
refresh_permit_workers_display()

root.protocol("WM_DELETE_WINDOW", lambda:(autosave_now(), root.destroy()))
root.after(1000, autosave_now)

# Final ensure bottom controls are on top
try:
    frame_bot.lift()
except Exception:
    pass


# Overlay bottom controls: ensure visibility above notebook and fixed at bottom
try:
    overlay_bot = ctk.CTkFrame(root, corner_radius=6)
    overlay_bot.place(relx=0.0, rely=1.0, relwidth=1.0, anchor='sw')
    overlay_bot.lift()

    # Folder choose + label
    of_frame = ctk.CTkFrame(overlay_bot, corner_radius=6); of_frame.pack(side='left', padx=(6,8), pady=6)
    make_button(of_frame, text="📁", command=choose_output_folder, width=4).pack(side="left")
    try:
        output_dir_label_overlay = ctk.CTkLabel(of_frame, text=output_dir_label.cget('text'))
    except Exception:
        output_dir_label_overlay = ctk.CTkLabel(of_frame, text="")
    output_dir_label_overlay.pack(side='left', padx=(6,0))

    # Autosave checkbox
    try:
        autosave_chk_overlay = ctk.CTkCheckBox(overlay_bot, text="Автосохранение", variable=autosave_var)
        autosave_chk_overlay.pack(side='left', padx=6)
    except Exception:
        pass

    # Other action buttons
    make_button(overlay_bot, text="Изменить исходник", command=open_source_selector, width=20).pack(side="left", padx=6)
    make_button(overlay_bot, text="Сгенерировать Все", command=generate_docx_all, width=20).pack(side="left", padx=6)
    make_button(overlay_bot, text="Сохранить профиль", command=lambda: save_profile(True), width=16).pack(side="left", padx=6)
    make_button(overlay_bot, text="Выйти", command=lambda:(autosave_now(), root.destroy()), width=10).pack(side="right", padx=6)    # Theme switch on overlay right
    try:
        try:
            theme_var  # noqa: F821
        except NameError:
            theme_var = tk.BooleanVar(value=(ctk.get_appearance_mode().lower()=='dark'))
        theme_switch_overlay = ctk.CTkSwitch(overlay_bot, text='Тёмная тема', command=_toggle_theme_btn, variable=theme_var)
        theme_switch_overlay.pack(side='right', padx=6)
    except Exception:
        pass
    # Keep overlay on top
    def _keep_overlay_on_top(e=None):
        try:
            overlay_bot.lift()
        except Exception:
            pass
    root.bind('<Configure>', _keep_overlay_on_top)
except Exception:
    pass



# Schedule removal of stray top widgets after UI is built
try:
    root.after(150, _remove_top_rects)
except Exception:
    pass

root.mainloop()