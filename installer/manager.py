"""
1C Bridge Manager — GUI для управления databases.json.

Запуск:
    pythonw.exe manager.py            (без чёрного окна консоли)
    python.exe  manager.py            (с консолью — для отладки)

Layout:
    +--------------------+----------------------------------+
    | [+] [-] [по умолч] | Краткое имя:  [_______________] |
    | ┌──────────────┐   | Описание:     [_______________] |
    | │ ut [default] │   | Платформа:    [V83.COMConnector]|
    | │ bp           │   | Тип:          (•) серверная    |
    | │ zup          │   |               ( ) файловая      |
    | └──────────────┘   | Сервер:       [_______________] |
    |                    | Имя ИБ:       [_______________] |
    |                    | [✓] Аутентификация Windows      |
    |                    | Логин/Пароль: ...               |
    |                    | Заметки для Claude:             |
    |                    | ┌─────────────────────────────┐ |
    |                    | │ Что в этой базе можно найти │ |
    |                    | └─────────────────────────────┘ |
    |                    | [Тест подключения] [Сохранить] |
    +--------------------+----------------------------------+
"""

from __future__ import annotations

import json
import os
import sys
import threading
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext


# ---------------------------------------------------------------------------
# Конфигурация
# ---------------------------------------------------------------------------

def find_databases_file() -> Path:
    """Ищем databases.json: переменная окружения → корень установки → CWD."""
    env_path = os.environ.get("ONEC_DATABASES_FILE", "").strip()
    if env_path:
        return Path(env_path)
    # manager.py обычно лежит в C:\Program Files\1cMcpBridge\manager\manager.py
    here = Path(__file__).resolve().parent
    candidates = [
        here.parent / "databases.json",
        here / "databases.json",
        Path("C:/Program Files/1cMcpBridge/databases.json"),
    ]
    for c in candidates:
        if c.exists():
            return c
    return candidates[0]  # default


DB_FILE = find_databases_file()


def load_config() -> dict:
    if not DB_FILE.exists():
        return {"version": 1, "default_database": "", "databases": {}}
    try:
        data = json.loads(DB_FILE.read_text(encoding="utf-8"))
        data.setdefault("version", 1)
        data.setdefault("default_database", "")
        data.setdefault("databases", {})
        return data
    except Exception as e:
        messagebox.showerror(
            "Ошибка чтения",
            f"Не удалось прочитать {DB_FILE}:\n{e}\n\nСоздаю новый конфиг."
        )
        return {"version": 1, "default_database": "", "databases": {}}


def save_config(config: dict) -> None:
    config["version"] = 1
    keys = list(config["databases"].keys())
    if not keys:
        config["default_database"] = ""
    elif config["default_database"] not in config["databases"]:
        config["default_database"] = keys[0]

    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    DB_FILE.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


# ---------------------------------------------------------------------------
# Поиск платформ 1С
# ---------------------------------------------------------------------------

def find_platforms() -> list[dict]:
    """Сканируем стандартные пути установки 1С."""
    found = []
    roots = [
        Path("C:/Program Files/1cv8"),
        Path("C:/Program Files (x86)/1cv8"),
    ]
    for root in roots:
        if not root.exists():
            continue
        for d in root.iterdir():
            if not d.is_dir():
                continue
            dll = d / "bin" / "comcntr.dll"
            try:
                if not dll.exists():
                    continue
            except (PermissionError, OSError):
                # Нет прав читать атрибуты — пропускаем эту версию
                continue
            version = d.name
            parts = version.split(".")
            if len(parts) < 2:
                continue
            try:
                major = int(parts[0] + parts[1])
            except ValueError:
                continue
            found.append({
                "version": version,
                "progid": f"V{major}.COMConnector",
                "dll_path": str(dll),
            })
    return found


# ---------------------------------------------------------------------------
# Тест подключения
# ---------------------------------------------------------------------------

def test_connection(progid: str, conn_str: str, dll_path: str = "") -> tuple[bool, str]:
    """Возвращает (успех, сообщение)."""
    import subprocess

    # Регистрируем коннектор если нужно (требует админа)
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        try:
            connector = win32com.client.Dispatch(progid)
        except Exception:
            if not dll_path or not Path(dll_path).exists():
                return False, f"ProgID {progid} не зарегистрирован, и comcntr.dll не найден."
            # Пытаемся регистрировать
            r = subprocess.run(
                ["regsvr32", "/s", dll_path],
                capture_output=True, text=True
            )
            if r.returncode != 0:
                return False, (
                    f"Не удалось зарегистрировать {dll_path}.\n"
                    f"Запусти Manager от имени администратора."
                )
            # И массовая регистрация остальных DLL — лечит TYPE_E_LIBNOTREGISTERED
            bin_dir = Path(dll_path).parent
            for d in bin_dir.glob("*.dll"):
                if d.name.lower() == "comcntr.dll":
                    continue
                subprocess.run(["regsvr32", "/s", str(d)], capture_output=True)
            connector = win32com.client.Dispatch(progid)

        ib = connector.Connect(conn_str)
        try:
            name = ib.Метаданные.Имя
            return True, f"Подключение успешно. Имя конфигурации: {name}"
        except Exception:
            return True, "Подключение успешно (имя конфигурации недоступно)."

    except Exception as e:
        return False, f"Ошибка подключения:\n{e}"


# ---------------------------------------------------------------------------
# Основное окно
# ---------------------------------------------------------------------------

class ManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("1C Bridge Manager")
        self.geometry("960x640")
        self.minsize(820, 540)

        # Иконка из корня установки, если есть
        try:
            icon_path = Path(__file__).resolve().parent.parent / "assets" / "icon.ico"
            if icon_path.exists():
                self.iconbitmap(default=str(icon_path))
        except Exception:
            pass

        self.config_data = load_config()
        self.current_key: str | None = None
        self.dirty = False

        self._build_ui()
        self._refresh_list()

        # При закрытии — спросить про несохранённые
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ----- Layout -----
    def _build_ui(self):
        # Toolbar сверху над всем
        toolbar = ttk.Frame(self, padding=(8, 6))
        toolbar.pack(fill=tk.X)
        ttk.Label(toolbar, text=f"Файл: {DB_FILE}", foreground="#666").pack(side=tk.LEFT)
        ttk.Button(toolbar, text="Открыть в Notepad", command=self._open_in_editor).pack(side=tk.RIGHT, padx=2)

        # Основной paned: список слева, форма справа
        main = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        # Левая панель: список баз
        left = ttk.Frame(main, padding=4)
        main.add(left, weight=1)

        list_toolbar = ttk.Frame(left)
        list_toolbar.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(list_toolbar, text="+ Добавить", command=self._on_add).pack(side=tk.LEFT, padx=2)
        ttk.Button(list_toolbar, text="– Удалить", command=self._on_delete).pack(side=tk.LEFT, padx=2)
        ttk.Button(list_toolbar, text="✓ По умолчанию", command=self._on_set_default).pack(side=tk.LEFT, padx=2)

        list_frame = ttk.Frame(left)
        list_frame.pack(fill=tk.BOTH, expand=True)
        self.listbox = tk.Listbox(list_frame, exportselection=False, font=("Segoe UI", 10))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb = ttk.Scrollbar(list_frame, command=self.listbox.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        # Правая панель: форма
        right = ttk.Frame(main, padding=8)
        main.add(right, weight=3)

        # Сетка с полями
        row = 0
        ttk.Label(right, text="Краткое имя:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_key = tk.StringVar()
        self.entry_key = ttk.Entry(right, textvariable=self.var_key, width=20)
        self.entry_key.grid(row=row, column=1, sticky="ew", pady=2)
        ttk.Label(right, text="(латиницей: ut, bp, zup)", foreground="#888").grid(row=row, column=2, sticky="w", padx=6)
        row += 1

        ttk.Label(right, text="Описание:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_description = tk.StringVar()
        ttk.Entry(right, textvariable=self.var_description).grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        row += 1

        ttk.Label(right, text="Платформа 1С:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_progid = tk.StringVar()
        self.combo_progid = ttk.Combobox(right, textvariable=self.var_progid, state="readonly", width=30)
        self.combo_progid.grid(row=row, column=1, sticky="w", pady=2)
        self.platforms = find_platforms()
        progid_options = [f"{p['progid']}  ({p['version']})" for p in self.platforms]
        if not progid_options:
            progid_options = ["V83.COMConnector", "V85.COMConnector"]
            self.combo_progid["state"] = "normal"
        self.combo_progid["values"] = progid_options
        if progid_options:
            self.combo_progid.set(progid_options[-1])
        row += 1

        ttk.Label(right, text="Тип базы:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        type_frame = ttk.Frame(right)
        type_frame.grid(row=row, column=1, columnspan=2, sticky="w", pady=2)
        self.var_type = tk.StringVar(value="server")
        ttk.Radiobutton(type_frame, text="Серверная", variable=self.var_type, value="server",
                        command=self._on_type_change).pack(side=tk.LEFT)
        ttk.Radiobutton(type_frame, text="Файловая", variable=self.var_type, value="file",
                        command=self._on_type_change).pack(side=tk.LEFT, padx=12)
        row += 1

        # Серверные поля
        ttk.Label(right, text="Сервер (Srvr):", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_server = tk.StringVar(value="127.0.0.1")
        self.entry_server = ttk.Entry(right, textvariable=self.var_server)
        self.entry_server.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        row += 1

        ttk.Label(right, text="Имя ИБ (Ref):", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_ref = tk.StringVar()
        self.entry_ref = ttk.Entry(right, textvariable=self.var_ref)
        self.entry_ref.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        row += 1

        # Файловые поля
        self.lbl_file = ttk.Label(right, text="Путь к базе:", anchor="w")
        self.lbl_file.grid(row=row, column=0, sticky="w", pady=2)
        self.var_file = tk.StringVar()
        file_frame = ttk.Frame(right)
        file_frame.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        self.entry_file = ttk.Entry(file_frame, textvariable=self.var_file)
        self.entry_file.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="...", width=3, command=self._browse_file).pack(side=tk.RIGHT, padx=(4, 0))
        row += 1

        # Аутентификация
        ttk.Label(right, text="Аутентификация:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_os_auth = tk.BooleanVar(value=True)
        ttk.Checkbutton(right, text="Средствами Windows (текущий пользователь)",
                        variable=self.var_os_auth, command=self._on_auth_change).grid(
            row=row, column=1, columnspan=2, sticky="w", pady=2)
        row += 1

        ttk.Label(right, text="Логин 1С:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_user = tk.StringVar()
        self.entry_user = ttk.Entry(right, textvariable=self.var_user)
        self.entry_user.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        row += 1

        ttk.Label(right, text="Пароль:", anchor="w").grid(row=row, column=0, sticky="w", pady=2)
        self.var_password = tk.StringVar()
        self.entry_password = ttk.Entry(right, textvariable=self.var_password, show="•")
        self.entry_password.grid(row=row, column=1, columnspan=2, sticky="ew", pady=2)
        row += 1

        # Notes — большой Text
        ttk.Label(right, text="Заметки для Claude\n(что в этой базе):",
                  anchor="w", justify="left").grid(row=row, column=0, sticky="nw", pady=(8, 2))
        self.text_notes = scrolledtext.ScrolledText(right, height=6, wrap=tk.WORD, font=("Segoe UI", 9))
        self.text_notes.grid(row=row, column=1, columnspan=2, sticky="ewns", pady=(8, 2))
        row += 1

        # Кнопки внизу формы
        btn_frame = ttk.Frame(right)
        btn_frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(12, 4))
        self.btn_test = ttk.Button(btn_frame, text="Тест подключения", command=self._on_test)
        self.btn_test.pack(side=tk.LEFT)
        self.btn_save = ttk.Button(btn_frame, text="Сохранить", command=self._on_save)
        self.btn_save.pack(side=tk.RIGHT)

        # Растягиваем колонку с полями
        right.columnconfigure(1, weight=1)
        right.rowconfigure(row - 1, weight=1)  # notes растягивается

        # Статус-бар
        self.status_var = tk.StringVar(value="Готово.")
        status = ttk.Label(self, textvariable=self.status_var, anchor="w",
                          relief=tk.SUNKEN, padding=(8, 2))
        status.pack(fill=tk.X, side=tk.BOTTOM)

        # Подсказка про перезапуск Claude Desktop
        hint = ttk.Label(self, text="После любых изменений нужно перезапустить Claude Desktop (Quit из трея → запустить снова).",
                        foreground="#a05a00", padding=(8, 4))
        hint.pack(fill=tk.X, side=tk.BOTTOM)

        self._on_type_change()
        self._on_auth_change()
        self._set_form_enabled(False)

    # ----- Список баз -----
    def _refresh_list(self):
        self.listbox.delete(0, tk.END)
        default = self.config_data.get("default_database", "")
        for key in sorted(self.config_data["databases"].keys()):
            cfg = self.config_data["databases"][key]
            marker = "  ●" if key == default else "   "
            display = f"{marker} {key}"
            desc = cfg.get("description", "")
            if desc:
                display += f"  — {desc}"
            self.listbox.insert(tk.END, display)
        self.status_var.set(f"Баз в списке: {self.listbox.size()}")

    def _selected_key(self) -> str | None:
        sel = self.listbox.curselection()
        if not sel:
            return None
        text = self.listbox.get(sel[0])
        # формат: "  ● key — desc"
        parts = text.lstrip(" ●").strip().split("  —", 1)
        return parts[0].strip()

    def _on_select(self, event=None):
        if self.dirty:
            if not messagebox.askyesno("Несохранённые изменения",
                                        "Есть несохранённые изменения. Отбросить?"):
                # Возвращаем выделение
                if self.current_key:
                    keys = sorted(self.config_data["databases"].keys())
                    if self.current_key in keys:
                        idx = keys.index(self.current_key)
                        self.listbox.selection_clear(0, tk.END)
                        self.listbox.selection_set(idx)
                return

        key = self._selected_key()
        if not key:
            self._set_form_enabled(False)
            return
        self.current_key = key
        self._load_into_form(self.config_data["databases"][key])
        self._set_form_enabled(True)
        self.dirty = False

    def _load_into_form(self, cfg: dict):
        self.var_key.set(self.current_key or "")
        self.var_description.set(cfg.get("description", ""))

        # ProgID
        progid = cfg.get("progid", "")
        for opt in self.combo_progid["values"]:
            if opt.startswith(progid):
                self.combo_progid.set(opt)
                break
        else:
            self.combo_progid.set(progid)

        # Парсим connection_string
        conn = cfg.get("connection_string", "")
        is_file = "File=" in conn
        self.var_type.set("file" if is_file else "server")

        def extract(key):
            import re
            m = re.search(rf'{key}="([^"]*)"', conn)
            return m.group(1) if m else ""

        self.var_file.set(extract("File"))
        self.var_server.set(extract("Srvr") or "127.0.0.1")
        self.var_ref.set(extract("Ref"))
        self.var_user.set(extract("Usr"))
        self.var_password.set(extract("Pwd"))
        self.var_os_auth.set(not extract("Usr"))

        # Notes
        self.text_notes.delete("1.0", tk.END)
        self.text_notes.insert("1.0", cfg.get("notes", ""))

        self._on_type_change()
        self._on_auth_change()

    def _set_form_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for w in [self.entry_key, self.combo_progid, self.entry_server, self.entry_ref,
                  self.entry_file, self.entry_user, self.entry_password,
                  self.btn_test, self.btn_save]:
            try:
                if w == self.combo_progid:
                    w["state"] = "readonly" if enabled and self.platforms else state
                else:
                    w["state"] = state
            except tk.TclError:
                pass
        try:
            self.text_notes["state"] = state
        except tk.TclError:
            pass

    # ----- Тип / аутентификация -----
    def _on_type_change(self):
        is_file = self.var_type.get() == "file"
        for w, show in [(self.entry_server, not is_file),
                        (self.entry_ref, not is_file)]:
            if show:
                w.configure(state="normal")
            else:
                w.configure(state="disabled")
        self.entry_file.configure(state="normal" if is_file else "disabled")

    def _on_auth_change(self):
        os_auth = self.var_os_auth.get()
        for w in [self.entry_user, self.entry_password]:
            w.configure(state="disabled" if os_auth else "normal")

    def _browse_file(self):
        path = filedialog.askdirectory(title="Выберите каталог файловой базы 1С")
        if path:
            self.var_file.set(path)

    # ----- Сборка connection_string -----
    def _build_connstr(self) -> str:
        if self.var_type.get() == "file":
            base = f'File="{self.var_file.get()}"'
        else:
            base = f'Srvr="{self.var_server.get()}";Ref="{self.var_ref.get()}"'
        if not self.var_os_auth.get():
            base += f';Usr="{self.var_user.get()}";Pwd="{self.var_password.get()}"'
        return base

    def _selected_progid_and_dll(self) -> tuple[str, str]:
        text = self.combo_progid.get()
        # "V83.COMConnector  (8.3.27.1859)"
        if "  (" in text:
            progid, version = text.split("  (")
            progid = progid.strip()
            version = version.rstrip(")").strip()
            for p in self.platforms:
                if p["version"] == version:
                    return progid, p["dll_path"]
            return progid, ""
        return text.strip(), ""

    # ----- Действия -----
    def _on_add(self):
        # Сначала спросим имя
        from tkinter import simpledialog
        key = simpledialog.askstring("Добавить базу",
                                      "Краткое имя новой базы (латиницей):",
                                      parent=self)
        if not key:
            return
        key = key.strip()
        import re
        if not re.match(r"^[a-zA-Z0-9_]+$", key):
            messagebox.showerror("Ошибка", "Только латинские буквы, цифры, _.")
            return
        if key in self.config_data["databases"]:
            messagebox.showerror("Ошибка", f"База '{key}' уже существует.")
            return

        self.config_data["databases"][key] = {
            "description": "",
            "progid": "V83.COMConnector",
            "connection_string": 'Srvr="127.0.0.1";Ref=""',
            "notes": "",
        }
        if not self.config_data["default_database"]:
            self.config_data["default_database"] = key
        save_config(self.config_data)
        self._refresh_list()

        # Выделяем новую
        keys = sorted(self.config_data["databases"].keys())
        idx = keys.index(key)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(idx)
        self.listbox.event_generate("<<ListboxSelect>>")
        self.entry_key.focus_set()

    def _on_delete(self):
        key = self._selected_key()
        if not key:
            return
        if not messagebox.askyesno("Удалить",
                                    f"Удалить базу '{key}' из списка?"):
            return
        self.config_data["databases"].pop(key, None)
        save_config(self.config_data)
        self.current_key = None
        self.dirty = False
        self._refresh_list()
        self._set_form_enabled(False)
        self.status_var.set(f"База '{key}' удалена.")

    def _on_set_default(self):
        key = self._selected_key()
        if not key:
            messagebox.showinfo("Подсказка", "Выбери базу в списке слева.")
            return
        self.config_data["default_database"] = key
        save_config(self.config_data)
        self._refresh_list()
        self.status_var.set(f"База по умолчанию: '{key}'")

    def _on_save(self):
        key = self.var_key.get().strip()
        import re
        if not re.match(r"^[a-zA-Z0-9_]+$", key):
            messagebox.showerror("Ошибка", "Краткое имя должно содержать только латинские буквы, цифры и _.")
            return

        # Если поменяли key — переименовать
        old_key = self.current_key
        if old_key and old_key != key:
            if key in self.config_data["databases"]:
                messagebox.showerror("Ошибка", f"База '{key}' уже существует.")
                return
            self.config_data["databases"][key] = self.config_data["databases"].pop(old_key)
            if self.config_data["default_database"] == old_key:
                self.config_data["default_database"] = key

        progid, dll = self._selected_progid_and_dll()
        cfg = {
            "description": self.var_description.get().strip(),
            "progid": progid,
            "connection_string": self._build_connstr(),
            "notes": self.text_notes.get("1.0", "end-1c").strip(),
        }
        if dll:
            cfg["dll_path"] = dll
        self.config_data["databases"][key] = cfg
        self.current_key = key
        save_config(self.config_data)
        self.dirty = False
        self._refresh_list()

        # Восстанавливаем выделение
        keys = sorted(self.config_data["databases"].keys())
        if key in keys:
            self.listbox.selection_clear(0, tk.END)
            self.listbox.selection_set(keys.index(key))

        self.status_var.set(f"База '{key}' сохранена. Перезапусти Claude Desktop.")
        messagebox.showinfo(
            "Сохранено",
            f"База '{key}' сохранена.\n\n"
            "Чтобы изменения вступили в силу — перезапусти Claude Desktop:\n"
            "  правый клик по иконке в трее → Quit → запустить снова."
        )

    def _on_test(self):
        progid, dll = self._selected_progid_and_dll()
        connstr = self._build_connstr()
        self.status_var.set("Проверяю подключение...")
        self.btn_test["state"] = "disabled"
        self.update()

        def worker():
            ok, msg = test_connection(progid, connstr, dll)
            self.after(0, lambda: self._show_test_result(ok, msg))

        threading.Thread(target=worker, daemon=True).start()

    def _show_test_result(self, ok: bool, msg: str):
        self.btn_test["state"] = "normal"
        if ok:
            self.status_var.set("Тест: успех.")
            messagebox.showinfo("Тест подключения", msg)
        else:
            self.status_var.set("Тест: неудача.")
            messagebox.showerror("Тест подключения", msg)

    def _open_in_editor(self):
        if not DB_FILE.exists():
            save_config(self.config_data)
        os.startfile(str(DB_FILE))  # type: ignore[attr-defined]

    def _on_close(self):
        # Проверяем dirty
        if self.dirty:
            r = messagebox.askyesnocancel(
                "Несохранённые изменения",
                "В форме есть изменения. Сохранить перед выходом?"
            )
            if r is None:
                return
            if r:
                self._on_save()
        self.destroy()


def main():
    # Включаем отслеживание dirty при изменении полей
    app = ManagerApp()

    def mark_dirty(*_):
        app.dirty = True

    # Привязываем к каждому изменению
    for var in [app.var_key, app.var_description, app.var_progid, app.var_type,
                app.var_server, app.var_ref, app.var_file,
                app.var_user, app.var_password, app.var_os_auth]:
        var.trace_add("write", mark_dirty)
    app.text_notes.bind("<KeyRelease>", lambda e: mark_dirty())

    app.mainloop()


if __name__ == "__main__":
    main()
