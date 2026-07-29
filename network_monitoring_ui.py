from __future__ import annotations

import tkinter as tk
from collections.abc import Callable
from tkinter import messagebox


BG = "#071311"
PANEL = "#0d1c1a"
PANEL_ALT = "#102725"
BORDER = "#1f4e46"
TEXT = "#ecfff7"
TEXT_SOFT = "#a6d5c5"
TEXT_MUTED = "#7ca394"
ACCENT = "#2f8fff"
ACCENT_DARK = "#1d527f"


class NetworkMonitoringUI:
    """Самостоятелен интерфейс за бъдещите инструменти за мрежово наблюдение."""

    MENU_ITEMS = (
        ("overview", "Обзор", "Мрежов статус"),
        ("live", "Live Monitoring", "Трафик в реално време"),
        ("speed", "Speed Test", "Скорост на връзката"),
        ("latency", "Ping и Latency", "Забавяне и загуба"),
        ("devices", "Устройства", "Хостове в локалната мрежа"),
        ("connections", "Връзки", "Активни TCP/UDP сесии"),
        ("dns", "DNS инструменти", "Проверка и диагностика"),
    )

    PAGE_CARDS = {
        "overview": (
            ("Интернет връзка", "Ще показва достъпност, активен адаптер и текущ IP адрес."),
            ("Мрежова активност", "Обобщение на входящия и изходящия трафик."),
            ("Състояние", "Предупреждения за забавяне, загуба на пакети и прекъсвания."),
        ),
        "live": (
            ("Download", "Графика на входящия трафик в реално време."),
            ("Upload", "Графика на изходящия трафик в реално време."),
            ("Използване", "Приложения и процеси с активни мрежови връзки."),
        ),
        "speed": (
            ("Download тест", "Измерване на скоростта за изтегляне."),
            ("Upload тест", "Измерване на скоростта за качване."),
            ("История", "Сравнение на последните измервания."),
        ),
        "latency": (
            ("Gateway", "Ping към локалния gateway и оценка на LAN връзката."),
            ("Internet", "Latency към избрани надеждни интернет endpoints."),
            ("Packet Loss", "Проверка за изгубени пакети и нестабилност."),
        ),
        "devices": (
            ("Открити устройства", "Списък на активните устройства в локалната мрежа."),
            ("Нови устройства", "Сигнал при появяване на непознато устройство."),
            ("Детайли", "IP, MAC, име и производител, когато са достъпни."),
        ),
        "connections": (
            ("TCP връзки", "Активни и слушащи TCP endpoints."),
            ("UDP endpoints", "Активни UDP портове на текущия компютър."),
            ("Процеси", "Свързване на мрежовите връзки с локалните процеси."),
        ),
        "dns": (
            ("DNS статус", "Текущи DNS сървъри и време за отговор."),
            ("Lookup", "Проверка на домейн, IPv4 и IPv6 записи."),
            ("Диагностика", "Откриване на DNS проблеми и кеширани грешки."),
        ),
    }

    def __init__(self, root: tk.Tk, on_back: Callable[[], None]) -> None:
        self.root = root
        self.on_back = on_back
        self.current_page = "overview"
        self.nav_buttons: dict[str, tk.Button] = {}
        self.root.title("WGA Network Monitoring")
        self.root.overrideredirect(False)
        self.root.configure(bg=BG)
        self.root.geometry("1220x820")
        self.root.minsize(1060, 720)
        self._build()
        self.show_page("overview")

    def _build(self) -> None:
        self.shell = tk.Frame(self.root, bg=BG)
        self.shell.pack(fill="both", expand=True)

        header = tk.Frame(self.shell, bg=PANEL, height=92, highlightbackground=BORDER, highlightthickness=1)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="WGA NETWORK MONITORING",
            bg=PANEL,
            fg=TEXT,
            font=("Segoe UI Semibold", 23),
        ).pack(anchor="w", padx=26, pady=(14, 0))
        tk.Label(
            header,
            text="Наблюдение, диагностика и анализ на локалната и интернет връзката",
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
            fg="#eef7ff",
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
            text="NETWORK TOOLS",
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
            text="WGA Network Monitoring\nЛокален модул",
            bg=PANEL,
            fg=TEXT_MUTED,
            justify="left",
            font=("Segoe UI", 9),
        ).pack(side="bottom", anchor="w", padx=22, pady=22)

        self.content = tk.Frame(body, bg=BG)
        self.content.pack(side="left", fill="both", expand=True, padx=28, pady=24)

    def show_page(self, page_id: str) -> None:
        self.current_page = page_id if page_id in self.PAGE_CARDS else "overview"
        for child in self.content.winfo_children():
            child.destroy()
        for item_id, button in self.nav_buttons.items():
            active = item_id == self.current_page
            button.configure(bg=PANEL_ALT if active else PANEL, fg=ACCENT if active else TEXT)

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
            text="Самостоятелен модул за бъдещи live мрежови проверки и диагностични инструменти.",
            bg=BG,
            fg=TEXT_MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 22))

        banner = tk.Frame(
            self.content,
            bg="#10253a",
            highlightbackground=ACCENT_DARK,
            highlightthickness=1,
        )
        banner.pack(fill="x", pady=(0, 22))
        tk.Label(
            banner,
            text="MONITORING READY",
            bg="#10253a",
            fg="#70b7ff",
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", padx=18, pady=(14, 2))
        tk.Label(
            banner,
            text="Наблюдението ще бъде локално и няма да изпраща мрежови данни към външни услуги.",
            bg="#10253a",
            fg="#c5e3ff",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=18, pady=(0, 14))

        cards = tk.Frame(self.content, bg=BG)
        cards.pack(fill="both", expand=True)
        for column in range(3):
            cards.grid_columnconfigure(column, weight=1, uniform="network")
        cards.grid_rowconfigure(0, weight=1)

        for column, (card_title, description) in enumerate(self.PAGE_CARDS[self.current_page]):
            card = tk.Frame(cards, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
            card.grid(row=0, column=column, sticky="nsew", padx=8)
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
            tk.Button(
                card,
                text="ОТВОРИ",
                command=lambda name=card_title: self._not_implemented(name),
                bg=ACCENT_DARK,
                activebackground=ACCENT,
                fg="#eef7ff",
                activeforeground="#071311",
                relief="flat",
                bd=0,
                cursor="hand2",
                font=("Segoe UI Semibold", 10),
                padx=22,
                pady=9,
            ).pack(side="bottom", anchor="w", padx=20, pady=24)

    def _not_implemented(self, feature: str) -> None:
        messagebox.showinfo(
            feature,
            "Екранът е подготвен. Реалната функция ще бъде добавена в следваща стъпка.",
            parent=self.root,
        )
