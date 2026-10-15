from __future__ import annotations

import ctypes
import json
import platform
import subprocess
import sys
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox, ttk


BG = "#071311"
PANEL = "#0d1c1a"
PANEL_ALT = "#102725"
BORDER = "#1f4e46"
TEXT = "#ecfff7"
TEXT_SOFT = "#a6d5c5"
TEXT_MUTED = "#7ca394"
ACCENT = "#2f8fff"
ACCENT_DARK = "#1d527f"
PROJECT_ROOT = Path(__file__).resolve().parent
AGENT_SCRIPT = PROJECT_ROOT / "wga_agent.py"
AGENT_EXE = PROJECT_ROOT / "wga_agent.exe"
AGENT_STATUS_CANDIDATES = (
    Path("C:/ProgramData/WGA/Agent/wga_agent_status.json"),
    PROJECT_ROOT / "wga_agent_status.json",
)
AGENT_REALTIME_CANDIDATES = (
    Path("C:/ProgramData/WGA/Agent/wga_realtime_activity.json"),
    PROJECT_ROOT / "wga_realtime_activity.json",
)
MACHINE_SNAPSHOT_DIRS = (
    Path("C:/ProgramData/WGA/Monitor/machines"),
    PROJECT_ROOT / "machines",
    PROJECT_ROOT / "agent_snapshots",
)


class NetworkMonitoringUI:
    """Самостоятелен интерфейс за бъдещите инструменти за мрежово наблюдение."""

    MENU_ITEMS = (
        ("overview", "Обзор", "Мрежов статус"),
        ("machines", "Machines", "Компютри и Remote Desktop"),
        ("live", "Live Monitoring", "Трафик в реално време"),
        ("speed", "Speed Test", "Скорост на връзката"),
        ("latency", "Ping и Latency", "Забавяне и загуба"),
        ("devices", "Устройства", "Хостове в локалната мрежа"),
        ("connections", "Връзки", "Активни TCP/UDP сесии"),
        ("dns", "DNS инструменти", "Проверка и диагностика"),
        ("agent", "Agent", "Инсталация на мониторинг агент"),
    )

    PAGE_CARDS = {
        "overview": (
            ("Интернет връзка", "Ще показва достъпност, активен адаптер и текущ IP адрес."),
            ("Мрежова активност", "Обобщение на входящия и изходящия трафик."),
            ("Състояние", "Предупреждения за забавяне, загуба на пакети и прекъсвания."),
        ),
        "machines": (
            ("Компютри", "Списък с WGA Agent машини и последен live статус."),
            ("Remote Desktop", "Двойно натискане отваря mstsc към избраната машина."),
            ("Shadow View", "Когато има session ID, WGA използва mstsc /shadow за гледане на текущата сесия."),
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
        "agent": (
            ("Инсталиране", "Инсталира локален WGA Agent и пита дали да смени името на компютъра."),
            ("SYSTEM режим", "Фоновата maintenance задача работи като SYSTEM и не зависи от правата на текущия потребител."),
            ("Admin Console", "Ctrl+Alt+W отваря статус, RDP, логове, Stop Agent и Kill Hung Process."),
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

        if self.current_page == "agent":
            self._show_agent_page()
            return
        if self.current_page == "machines":
            self._show_machines_page()
            return

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

    def _read_json_file(self, path: Path) -> dict[str, object]:
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            return data if isinstance(data, dict) else {}
        except (OSError, json.JSONDecodeError):
            return {}

    def _machine_snapshots(self) -> list[dict[str, object]]:
        snapshots: list[dict[str, object]] = []
        seen_paths: set[Path] = set()

        for path in (*AGENT_REALTIME_CANDIDATES, *AGENT_STATUS_CANDIDATES):
            if path.exists() and path not in seen_paths:
                seen_paths.add(path)
                data = self._read_json_file(path)
                if data:
                    snapshots.append(self._normalize_machine_snapshot(data, path))

        for directory in MACHINE_SNAPSHOT_DIRS:
            if not directory.exists():
                continue
            for path in sorted(directory.glob("*.json")):
                if path in seen_paths:
                    continue
                seen_paths.add(path)
                data = self._read_json_file(path)
                if data:
                    snapshots.append(self._normalize_machine_snapshot(data, path))

        deduped: dict[str, dict[str, object]] = {}
        for machine in snapshots:
            key = str(machine.get("hostname") or machine.get("primary_ip") or machine.get("source"))
            previous = deduped.get(key)
            if previous is None or str(machine.get("last_seen", "")) >= str(previous.get("last_seen", "")):
                deduped[key] = machine
        return sorted(deduped.values(), key=lambda item: str(item.get("hostname", "")).lower())

    def _normalize_machine_snapshot(self, data: dict[str, object], source: Path) -> dict[str, object]:
        activity = data.get("realtime_activity", data)
        if not isinstance(activity, dict):
            activity = {}
        rdp = activity.get("remote_desktop", {})
        if not isinstance(rdp, dict):
            rdp = {}
        foreground = activity.get("foreground_window", {})
        if not isinstance(foreground, dict):
            foreground = {}
        local_ips = rdp.get("local_ips") if isinstance(rdp.get("local_ips"), list) else data.get("local_ips", [])
        if not isinstance(local_ips, list):
            local_ips = []
        sessions = rdp.get("sessions", []) if isinstance(rdp.get("sessions"), list) else []
        active_session = next(
            (
                session for session in sessions
                if isinstance(session, dict)
                and session.get("id") not in {None, 0}
                and str(session.get("state", "")).lower() in {"active", "disc"}
            ),
            None,
        )
        hostname = str(rdp.get("hostname") or data.get("hostname") or data.get("node") or source.stem)
        return {
            "source": str(source),
            "hostname": hostname,
            "primary_ip": str(local_ips[0]) if local_ips else "",
            "ips": [str(item) for item in local_ips],
            "last_seen": str(activity.get("timestamp") or data.get("timestamp") or ""),
            "user": str(activity.get("interactive_user") or ""),
            "foreground_process": str(foreground.get("process_name") or ""),
            "foreground_title": str(foreground.get("window_title") or ""),
            "rdp_enabled": bool(rdp.get("enabled")),
            "nla_enabled": bool(rdp.get("nla_enabled")),
            "sessions": sessions,
            "active_session": active_session,
            "connect_host": str(local_ips[0] if local_ips else hostname),
        }

    def _show_machines_page(self) -> None:
        tk.Label(
            self.content,
            text="Machines",
            bg=BG,
            fg=TEXT,
            font=("Segoe UI Semibold", 24),
        ).pack(anchor="w")
        tk.Label(
            self.content,
            text="Компютри с WGA Agent snapshots. Двойно натискане отваря Remote Desktop към избраната машина.",
            bg=BG,
            fg=TEXT_MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 18))

        toolbar = tk.Frame(self.content, bg=BG)
        toolbar.pack(fill="x", pady=(0, 10))
        tk.Button(
            toolbar,
            text="ОБНОВИ",
            command=lambda: self.show_page("machines"),
            bg="#173c36",
            activebackground="#1f554c",
            fg="#d9fff1",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=16,
            pady=8,
        ).pack(side="left")
        tk.Button(
            toolbar,
            text="ОТВОРИ RDP",
            command=lambda: self._open_selected_machine(prefer_shadow=False),
            bg=ACCENT_DARK,
            activebackground=ACCENT,
            fg="#eef7ff",
            activeforeground="#071311",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=16,
            pady=8,
        ).pack(side="left", padx=(8, 0))
        tk.Button(
            toolbar,
            text="SHADOW VIEW",
            command=lambda: self._open_selected_machine(prefer_shadow=True),
            bg="#245634",
            activebackground="#2f7044",
            fg="#f3fff5",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=16,
            pady=8,
        ).pack(side="left", padx=(8, 0))

        wrapper = tk.Frame(self.content, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        wrapper.pack(fill="both", expand=True)

        columns = ("hostname", "ip", "user", "foreground", "rdp", "session", "last_seen")
        self.machine_tree = ttk.Treeview(wrapper, columns=columns, show="headings", height=14)
        headings = {
            "hostname": "Computer",
            "ip": "IP",
            "user": "User",
            "foreground": "Active App",
            "rdp": "RDP",
            "session": "Session",
            "last_seen": "Last Seen",
        }
        widths = {
            "hostname": 160,
            "ip": 120,
            "user": 150,
            "foreground": 220,
            "rdp": 80,
            "session": 90,
            "last_seen": 170,
        }
        for column in columns:
            self.machine_tree.heading(column, text=headings[column])
            self.machine_tree.column(column, width=widths[column], anchor="w", stretch=column in {"foreground", "last_seen"})
        self.machine_tree.pack(fill="both", expand=True, padx=14, pady=(14, 8))
        self.machine_tree.bind("<Double-1>", lambda _event: self._open_selected_machine(prefer_shadow=True))

        self.machine_details = tk.Text(
            wrapper,
            bg="#07100f",
            fg="#d8fff1",
            insertbackground="#d8fff1",
            relief="flat",
            wrap="word",
            font=("Consolas", 10),
            height=7,
            padx=12,
            pady=10,
        )
        self.machine_details.pack(fill="x", padx=14, pady=(0, 14))
        self.machine_tree.bind("<<TreeviewSelect>>", lambda _event: self._update_machine_details())

        self.machines = self._machine_snapshots()
        for index, machine in enumerate(self.machines):
            active_session = machine.get("active_session")
            session_text = ""
            if isinstance(active_session, dict):
                session_text = f"{active_session.get('id', '')} {active_session.get('state', '')}".strip()
            foreground = machine.get("foreground_process") or "N/A"
            if machine.get("foreground_title"):
                foreground = f"{foreground} - {machine.get('foreground_title')}"
            self.machine_tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    machine.get("hostname", ""),
                    machine.get("primary_ip", ""),
                    machine.get("user", ""),
                    foreground,
                    "On" if machine.get("rdp_enabled") else "Off",
                    session_text,
                    machine.get("last_seen", ""),
                ),
            )
        if self.machines:
            self.machine_tree.selection_set("0")
            self._update_machine_details()
        else:
            self.machine_details.insert(
                "end",
                "Няма намерени machine snapshots.\n\n"
                "Очаквани места:\n"
                f"- {AGENT_REALTIME_CANDIDATES[0]}\n"
                f"- {MACHINE_SNAPSHOT_DIRS[0]}\\*.json\n"
                f"- {MACHINE_SNAPSHOT_DIRS[1]}\\*.json\n",
            )
            self.machine_details.configure(state="disabled")

    def _selected_machine(self) -> dict[str, object] | None:
        if not hasattr(self, "machine_tree"):
            return None
        selection = self.machine_tree.selection()
        if not selection:
            return None
        try:
            return self.machines[int(selection[0])]
        except (ValueError, IndexError, AttributeError):
            return None

    def _update_machine_details(self) -> None:
        machine = self._selected_machine()
        if machine is None:
            return
        lines = [
            f"Computer: {machine.get('hostname', 'N/A')}",
            f"IP addresses: {', '.join(machine.get('ips', [])) if isinstance(machine.get('ips'), list) else 'N/A'}",
            f"User: {machine.get('user', '') or 'N/A'}",
            f"Foreground: {machine.get('foreground_process', '') or 'N/A'}",
            f"Window: {machine.get('foreground_title', '') or 'N/A'}",
            f"RDP enabled: {'Yes' if machine.get('rdp_enabled') else 'No'}",
            f"NLA enabled: {'Yes' if machine.get('nla_enabled') else 'No'}",
            f"Source: {machine.get('source', '')}",
            "",
            "Double-click opens Shadow View when a session is available; otherwise it opens normal RDP.",
        ]
        self.machine_details.configure(state="normal")
        self.machine_details.delete("1.0", "end")
        self.machine_details.insert("end", "\n".join(lines))
        self.machine_details.configure(state="disabled")

    def _open_selected_machine(self, *, prefer_shadow: bool) -> None:
        machine = self._selected_machine()
        if machine is None:
            messagebox.showinfo("Remote Desktop", "Изберете машина от списъка.", parent=self.root)
            return
        host = str(machine.get("connect_host") or machine.get("hostname") or "").strip()
        if not host:
            messagebox.showerror("Remote Desktop", "Няма hostname или IP за избраната машина.", parent=self.root)
            return

        command = ["mstsc", f"/v:{host}"]
        active_session = machine.get("active_session")
        if prefer_shadow and isinstance(active_session, dict):
            session_id = active_session.get("id")
            if isinstance(session_id, int) and session_id > 0:
                command = ["mstsc", f"/shadow:{session_id}", f"/v:{host}", "/control", "/prompt"]

        try:
            subprocess.Popen(command)
        except OSError as exc:
            messagebox.showerror("Remote Desktop", f"Неуспешно стартиране на mstsc:\n{exc}", parent=self.root)

    def _show_agent_page(self) -> None:
        tk.Label(
            self.content,
            text="Agent",
            bg=BG,
            fg=TEXT,
            font=("Segoe UI Semibold", 24),
        ).pack(anchor="w")
        tk.Label(
            self.content,
            text="Локален мониторинг агент за статус на компютъра и бъдещо управление от WGA.",
            bg=BG,
            fg=TEXT_MUTED,
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(3, 22))

        actions = tk.Frame(self.content, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        actions.pack(fill="x", pady=(0, 18))
        tk.Label(
            actions,
            text="Инсталация на агент",
            bg=PANEL,
            fg=TEXT,
            font=("Segoe UI Semibold", 15),
        ).pack(anchor="w", padx=20, pady=(18, 6))
        tk.Label(
            actions,
            text="При инсталация агентът ще попита дали да смени името на компютъра. За смяната са нужни Administrator права и рестарт.",
            bg=PANEL,
            fg=TEXT_SOFT,
            wraplength=780,
            justify="left",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=20)
        tk.Label(
            actions,
            text="След инсталация агентът работи тихо във фон като SYSTEM. Ctrl+Alt+W отваря admin-only console със live activity, Enable RDP, mstsc/shadow команди, Stop Agent и remote kill само за Not Responding програми.",
            bg=PANEL,
            fg=TEXT_MUTED,
            wraplength=780,
            justify="left",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=20, pady=(8, 0))

        button_row = tk.Frame(actions, bg=PANEL)
        button_row.pack(fill="x", padx=20, pady=18)
        tk.Button(
            button_row,
            text="ИНСТАЛИРАЙ AGENT",
            command=self._run_agent_installer,
            bg=ACCENT_DARK,
            activebackground=ACCENT,
            fg="#eef7ff",
            activeforeground="#071311",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=22,
            pady=10,
        ).pack(side="left")
        tk.Button(
            button_row,
            text="ОБНОВИ STATUS",
            command=lambda: self.show_page("agent"),
            bg="#173c36",
            activebackground="#1f554c",
            fg="#d9fff1",
            activeforeground="#ffffff",
            relief="flat",
            bd=0,
            cursor="hand2",
            font=("Segoe UI Semibold", 10),
            padx=18,
            pady=10,
        ).pack(side="left", padx=(10, 0))

        status_panel = tk.Frame(self.content, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        status_panel.pack(fill="both", expand=True)
        tk.Label(
            status_panel,
            text="Последен agent статус",
            bg=PANEL,
            fg=TEXT,
            font=("Segoe UI Semibold", 15),
        ).pack(anchor="w", padx=20, pady=(18, 8))

        status_text = tk.Text(
            status_panel,
            bg="#07100f",
            fg="#d8fff1",
            insertbackground="#d8fff1",
            relief="flat",
            wrap="word",
            font=("Consolas", 10),
            padx=14,
            pady=12,
            height=14,
        )
        status_text.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        status_text.insert("end", self._format_agent_status())
        status_text.configure(state="disabled")

    def _format_agent_status(self) -> str:
        status_path = next((path for path in AGENT_STATUS_CANDIDATES if path.exists()), None)
        if status_path is None:
            return (
                "Няма намерен agent status файл.\n\n"
                "Инсталирайте агента или стартирайте wga_agent.py --status, за да се създаде първи статус."
            )
        try:
            data = json.loads(status_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            return f"Status файлът не може да бъде прочетен:\n{status_path}\n\n{exc}"

        hardware = data.get("hardware_info", {}) if isinstance(data.get("hardware_info"), dict) else {}
        system_info = hardware.get("computer_system", {}) if isinstance(hardware.get("computer_system"), dict) else {}
        memory = hardware.get("memory", {}) if isinstance(hardware.get("memory"), dict) else {}
        lines = [
            f"Файл: {status_path}",
            f"Време: {data.get('timestamp', 'N/A')}",
            f"Компютър: {data.get('hostname', 'N/A')}",
            f"Платформа: {data.get('platform', 'N/A')} {data.get('platform_release', '')}",
            f"Модел: {system_info.get('manufacturer', 'N/A')} {system_info.get('model', '')}",
            f"IP адреси: {', '.join(data.get('local_ips', [])) or 'N/A'}",
            f"Онлайн: {'Да' if data.get('online') else 'Не'}",
            f"Uptime: {data.get('system_uptime', 'N/A')}",
            f"RAM: {memory.get('used_gb', 'N/A')} / {memory.get('total_gb', 'N/A')} GB",
        ]
        return "\n".join(lines)

    def _run_agent_installer(self) -> None:
        if platform.system() != "Windows":
            messagebox.showerror("WGA Agent", "Инсталацията на агента е предназначена за Windows.", parent=self.root)
            return

        if AGENT_EXE.exists():
            executable = str(AGENT_EXE)
            parameters = "--install"
            fallback_command = [executable, "--install"]
        elif AGENT_SCRIPT.exists():
            executable = sys.executable
            parameters = f'"{AGENT_SCRIPT}" --install'
            fallback_command = [executable, str(AGENT_SCRIPT), "--install"]
        else:
            messagebox.showerror("WGA Agent", "wga_agent.py не е намерен в папката на приложението.", parent=self.root)
            return

        try:
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                executable,
                parameters,
                str(PROJECT_ROOT),
                1,
            )
            if result <= 32:
                raise OSError(f"ShellExecuteW failed with code {result}")
        except Exception as exc:
            try:
                subprocess.Popen(fallback_command, cwd=str(PROJECT_ROOT))
            except OSError:
                messagebox.showerror("WGA Agent", f"Неуспешно стартиране на инсталатора:\n{exc}", parent=self.root)
