from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import messagebox

from windows10_low_ram_profile import (
    BACKUP_FILE,
    apply_profile,
    find_open_shell,
    profile_descriptions,
    restore_profile,
)


BG = "#071311"
PANEL = "#0d1c1a"
PANEL_ALT = "#102725"
BORDER = "#1f4e46"
TEXT = "#ecfff7"
TEXT_SOFT = "#a6d5c5"
TEXT_MUTED = "#7ca394"
ACCENT = "#d0a94a"
ACCENT_DARK = "#725b24"


class OptimizationUI:
    """Самостоятелен интерфейс за бъдещите инструменти за оптимизация."""

    MENU_ITEMS = (
        ("windows10", "WINDOWS 10", "Оптимизация за Windows 10"),
        ("windows11", "WINDOWS 11", "Оптимизация за Windows 11"),
    )

    PAGE_CARDS = {
        "windows10": (
            ("RAM профил за слаби системи", "Специален Windows 10 профил за компютри с 2 GB или 4 GB RAM памет."),
            ("Почистване", "Временни файлове, системен кеш, Windows Update cache и кошче."),
            ("Startup приложения", "Преглед и управление на програмите, стартиращи с Windows 10."),
            ("Услуги и процеси", "Оценка на ненужното фоново натоварване в Windows 10."),
            ("Поверителност", "Диагностични данни, разрешения и фонови приложения."),
            ("Защита и Restore Point", "Точка за възстановяване преди всяка група промени."),
        ),
        "windows11": (
            ("Бърза оптимизация", "Безопасен Windows 11 профил с предварителен преглед на действията."),
            ("Почистване", "Временни файлове, Delivery Optimization, update cache и кошче."),
            ("Startup приложения", "Преглед и управление на програмите, стартиращи с Windows 11."),
            ("Услуги и Widgets", "Оценка на фонови услуги, Widgets и допълнителни компоненти."),
            ("Поверителност", "Диагностични данни, advertising ID, разрешения и background access."),
            ("Защита и Restore Point", "Точка за възстановяване и история на приложените промени."),
        ),
    }

    def __init__(self, root: tk.Tk, on_back: Callable[[], None]) -> None:
        self.root = root
        self.on_back = on_back
        self.current_page = "windows10"
        self.nav_buttons: dict[str, tk.Button] = {}
        self.root.title("WGA Optimization")
        self.root.overrideredirect(False)
        self.root.configure(bg=BG)
        self.root.geometry("1220x820")
        self.root.minsize(1060, 720)
        self._build()
        self.show_page("windows10")

    def _build(self) -> None:
        self.shell = tk.Frame(self.root, bg=BG)
        self.shell.pack(fill="both", expand=True)

        header = tk.Frame(self.shell, bg=PANEL, height=92, highlightbackground=BORDER, highlightthickness=1)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="WGA ОПТИМИЗАЦИЯ",
            bg=PANEL,
            fg=TEXT,
            font=("Segoe UI Semibold", 23),
        ).pack(anchor="w", padx=26, pady=(14, 0))
        tk.Label(
            header,
            text="Контролен център за безопасно почистване, настройване и ускоряване на Windows",
            bg=PANEL,
            fg=TEXT_SOFT,
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=27, pady=(2, 0))
        tk.Button(
            header,
            text="Назад към модулите",
            command=self.on_back,
            bg=ACCENT_DARK,
            activebackground=ACCENT,
            fg="#fff8df",
            activeforeground="#071311",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=18,
            pady=9,
        ).place(relx=1.0, x=-26, y=24, anchor="ne")

        body = tk.Frame(self.shell, bg=BG)
        body.pack(fill="both", expand=True)

        sidebar = tk.Frame(body, bg=PANEL, width=285, highlightbackground=BORDER, highlightthickness=1)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(
            sidebar,
            text="OPTIMIZATION",
            bg=PANEL,
            fg=ACCENT,
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", padx=22, pady=(24, 14))

        for page_id, title, subtitle in self.MENU_ITEMS:
            button = tk.Button(
                sidebar,
                text=f"{title}\n{subtitle}",
                command=lambda target=page_id: self.show_page(target),
                bg=PANEL,
                activebackground=PANEL_ALT,
                fg=TEXT,
                activeforeground=TEXT,
                justify="left",
                anchor="w",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Segoe UI Semibold", 10),
                padx=22,
                pady=10,
            )
            button.pack(fill="x", padx=10, pady=3)
            self.nav_buttons[page_id] = button

        tk.Label(
            sidebar,
            text="WGA Optimization\nЛокален модул",
            bg=PANEL,
            fg=TEXT_MUTED,
            justify="left",
            font=("Segoe UI", 9),
        ).pack(side="bottom", anchor="w", padx=22, pady=22)

        self.content = tk.Frame(body, bg=BG)
        self.content.pack(side="left", fill="both", expand=True, padx=28, pady=24)

    def show_page(self, page_id: str) -> None:
        self.current_page = page_id if page_id in self.PAGE_CARDS else "windows10"
        for child in self.content.winfo_children():
            child.destroy()
        for item_id, button in self.nav_buttons.items():
            active = item_id == self.current_page
            button.configure(
                bg=PANEL_ALT if active else PANEL,
                fg=ACCENT if active else TEXT,
            )

        title = next(item[1] for item in self.MENU_ITEMS if item[0] == self.current_page)
        tk.Label(
            self.content,
            text=title,
            bg=BG,
            fg=TEXT,
            font=("Segoe UI Semibold", 24),
        ).pack(anchor="w")
        tk.Label(
            self.content,
            text="Модулът е подготвен за добавяне на реални проверки и контролирани действия.",
            bg=BG,
            fg=TEXT_MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 22))

        banner = tk.Frame(
            self.content,
            bg="#2b2514",
            highlightbackground=ACCENT_DARK,
            highlightthickness=1,
        )
        banner.pack(fill="x", pady=(0, 22))
        tk.Label(
            banner,
            text="БЕЗОПАСЕН РЕЖИМ",
            bg="#2b2514",
            fg=ACCENT,
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", padx=18, pady=(14, 2))
        tk.Label(
            banner,
            text="Няма да се изпълняват промени без предварителен преглед и потвърждение.",
            bg="#2b2514",
            fg="#e8dba8",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=18, pady=(0, 14))

        cards = tk.Frame(self.content, bg=BG)
        cards.pack(fill="both", expand=True)
        for column in range(3):
            cards.grid_columnconfigure(column, weight=1, uniform="optimization")
        cards.grid_rowconfigure(0, weight=1)
        cards.grid_rowconfigure(1, weight=1)

        for column, (card_title, description) in enumerate(self.PAGE_CARDS[self.current_page]):
            card = tk.Frame(
                cards,
                bg=PANEL,
                highlightbackground=BORDER,
                highlightthickness=1,
            )
            row = column // 3
            card_column = column % 3
            card.grid(row=row, column=card_column, sticky="nsew", padx=8, pady=8)
            tk.Label(
                card,
                text=card_title,
                bg=PANEL,
                fg=TEXT,
                wraplength=220,
                justify="left",
                font=("Segoe UI Semibold", 15),
            ).pack(anchor="w", padx=20, pady=(28, 12))
            tk.Label(
                card,
                text=description,
                bg=PANEL,
                fg=TEXT_SOFT,
                wraplength=220,
                justify="left",
                font=("Segoe UI", 10),
            ).pack(anchor="w", padx=20)
            is_low_ram_profile = self.current_page == "windows10" and column == 0
            tk.Button(
                card,
                text="ОПТИМИЗАЦИЯ ЗА 2 И 4 GB RAM ПАМЕТ" if is_low_ram_profile else "ОТВОРИ",
                command=self._open_low_ram_profile if is_low_ram_profile else lambda name=card_title: self._not_implemented(name),
                bg=ACCENT_DARK,
                activebackground=ACCENT,
                fg="#fff8df",
                activeforeground="#071311",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Segoe UI Semibold", 10),
                padx=22,
                pady=9,
                wraplength=210,
            ).pack(side="bottom", anchor="w", padx=20, pady=24)

    def _open_low_ram_profile(self) -> None:
        open_shell_status = (
            "Open-Shell е намерен и XP менюто ще бъде приложено."
            if find_open_shell()
            else "Open-Shell не е намерен. Windows оптимизациите могат да се приложат, но XP менюто ще бъде пропуснато."
        )
        details = "\n".join(f"• {item}" for item in profile_descriptions())
        action = "ВЪРНИ НАСТРОЙКИТЕ" if BACKUP_FILE.exists() else "ПРИЛОЖИ ПРОФИЛА"
        prompt = (
            "ПРОФИЛ WINDOWS 10 ЗА 2/4 GB RAM\n\n"
            f"{open_shell_status}\n\n"
            f"{details}\n\n"
            "Няма да бъдат спирани Windows Update, Defender, Firewall или критични услуги.\n\n"
            f"Изберете „Да“, за да изпълните: {action}."
        )
        if not messagebox.askyesno("WGA безопасна оптимизация", prompt, parent=self.root):
            return
        try:
            messages = restore_profile() if BACKUP_FILE.exists() else apply_profile()
        except Exception as exc:
            messagebox.showerror("WGA оптимизация", str(exc), parent=self.root)
            return
        messagebox.showinfo("WGA оптимизация", "\n\n".join(messages), parent=self.root)

    def _not_implemented(self, feature: str) -> None:
        messagebox.showinfo(
            feature,
            "Екранът е подготвен. Реалната функция ще бъде добавена в следваща стъпка.",
            parent=self.root,
        )
