from __future__ import annotations

import importlib
import logging
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, Signal, Slot

from services.paths import loading_asset, project_root


STARTUP_LOG = project_root() / "logs" / "wga-startup.log"


@dataclass(frozen=True)
class StartupStep:
    key: str
    label: str
    progress: int
    required: bool = True


STARTUP_STEPS = (
    StartupStep("configuration", "Зареждане на основната конфигурация", 12),
    StartupStep("language", "Проверка на езиковите настройки", 24, required=False),
    StartupStep("health", "Събиране на системната информация", 42, required=False),
    StartupStep("components", "Проверка на състоянието на компонентите", 60, required=False),
    StartupStep("software", "Проверка на наличния софтуер", 78, required=False),
    StartupStep("updates", "Проверка за актуализация", 88, required=False),
    StartupStep("assets", "Зареждане на графичните елементи", 92),
    StartupStep("interface", "Инициализиране на интерфейса", 96),
    StartupStep("launcher", "Подготовка на WGA началния панел", 98),
    StartupStep("ready", "Готово", 100),
)


class StartupWorker(QObject):
    step_started = Signal(str, str)
    step_completed = Signal(str, str, int)
    step_failed = Signal(str, str, str, bool)
    warning = Signal(str)
    finished = Signal(bool)

    def __init__(self, steps: tuple[StartupStep, ...] = STARTUP_STEPS, parent=None) -> None:
        super().__init__(parent)
        self.steps = steps
        self._logger = self._build_logger()
        self.preloaded_state: dict[str, object] = {}
        self._version_info: dict[str, object] | None = None

    @Slot()
    def run(self) -> None:
        has_required_error = False
        for step in self.steps[:7]:
            self.step_started.emit(step.key, step.label)
            try:
                detail = self._run_step(step.key)
            except Exception as exc:
                self._logger.exception("Startup step failed: %s", step.key)
                self.step_failed.emit(step.key, step.label, str(exc), step.required)
                if step.required:
                    has_required_error = True
                    break
                self.warning.emit(f"{step.label}: {exc}")
            else:
                self.step_completed.emit(step.key, detail, step.progress)
        self.finished.emit(not has_required_error)

    def _run_step(self, key: str) -> str:
        if key == "configuration":
            legacy_app = importlib.import_module("legacy_app")
            self._version_info = legacy_app.load_version_info()
            legacy_app.load_secure_store()
            legacy_app.load_settings()
            config_file = project_root() / "tasks.json"
            if config_file.exists():
                config_file.read_text(encoding="utf-8")
            return "Основната конфигурация и защитените настройки са заредени."
        if key == "language":
            legacy_app = importlib.import_module("legacy_app")
            try:
                self.preloaded_state["language_status"] = legacy_app.get_language_status()
                return "Езиковите настройки са прочетени."
            except Exception as exc:
                self.preloaded_state["language_status_error"] = str(exc)
                raise
        if key == "health":
            legacy_app = importlib.import_module("legacy_app")
            try:
                self.preloaded_state["health_items"] = legacy_app.collect_health_items()
                return "Системната информация е събрана."
            except Exception as exc:
                self.preloaded_state["health_error"] = str(exc)
                raise
        if key == "components":
            try:
                probe = self._build_legacy_probe()
                self.preloaded_state["component_status_rows"] = probe._dashboard_component_rows()
                return "Компонентите са проверени."
            except Exception as exc:
                self.preloaded_state["component_status_error"] = str(exc)
                raise
        if key == "software":
            try:
                probe = self._build_legacy_probe()
                tasks = probe._auto_install_tasks()
                status_map: dict[str, tuple[bool, str]] = {}
                for task in tasks:
                    status_map[task["id"]] = probe._safe_task_install_state(task)
                self.preloaded_state["program_selector_tasks"] = tasks
                self.preloaded_state["program_selector_status"] = status_map
                return f"Проверени са {len(tasks)} софтуерни задачи."
            except Exception as exc:
                self.preloaded_state["program_selector_error"] = str(exc)
                raise
        if key == "updates":
            legacy_app = importlib.import_module("legacy_app")
            try:
                version_info = self._version_info or legacy_app.load_version_info()
                result = legacy_app.check_for_updates(
                    str(version_info.get("version", "")),
                    str(version_info.get("version_info_url", "")),
                )
                self.preloaded_state["update_result"] = result
                if result.status == "up_to_date":
                    return "Програмата е актуална."
                if result.status == "update_available":
                    return f"Налична е нова версия: v{result.latest_version}."
                if self._looks_like_no_internet(result.error):
                    return "В момента няма връзка с интернет."
                return result.error or "Update проверката приключи с предупреждение."
            except Exception as exc:
                self.preloaded_state["update_error"] = str(exc)
                raise
        if key == "assets":
            required_assets = (
                "background_clean.png",
                "planet_overlay.png",
                "gold_waves_1.png",
                "gold_waves_2.png",
                "gold_light_line.png",
            )
            missing = [name for name in required_assets if not loading_asset(name).exists()]
            if missing:
                raise FileNotFoundError("Липсват loading assets: " + ", ".join(missing))
            return "Графичните слоеве за loading екрана са намерени."
        return "Готово."

    def _build_legacy_probe(self) -> object:
        legacy_app = importlib.import_module("legacy_app")
        probe = legacy_app.MainMenuUI.__new__(legacy_app.MainMenuUI)
        probe.root = None
        probe.settings = legacy_app.load_settings()
        probe.secure_store = legacy_app.load_secure_store()
        probe.launch_info = legacy_app.get_launch_location_info()
        probe.resource_status = legacy_app.check_resource_status(project_root())
        probe.version_info = self._version_info or legacy_app.load_version_info()
        probe.office_inventory_cache = {}
        probe.office_online_cache = {}
        probe.office_maintenance_cache = {}
        probe.adobe_reader_status_cache = None
        probe.language_status_cache = self.preloaded_state.get("language_status")
        probe.nexus_admin_status_cache = None
        probe.program_selector_tasks_cache = []
        probe.program_selector_status_cache = {}
        probe.component_status_cache = None
        probe.latest_health_items = list(self.preloaded_state.get("health_items", [])) if isinstance(self.preloaded_state.get("health_items"), list) else []
        return probe

    def _looks_like_no_internet(self, error: str) -> bool:
        text = error.lower()
        return any(
            marker in text
            for marker in (
                "timed out",
                "getaddrinfo",
                "name or service not known",
                "no address associated",
                "network is unreachable",
                "temporary failure",
                "няма връзка",
                "11001",
            )
        )

    def _build_logger(self) -> logging.Logger:
        STARTUP_LOG.parent.mkdir(parents=True, exist_ok=True)
        logger = logging.getLogger("wga.startup")
        logger.setLevel(logging.INFO)
        if not any(isinstance(handler, logging.FileHandler) and Path(handler.baseFilename) == STARTUP_LOG for handler in logger.handlers):
            handler = logging.FileHandler(STARTUP_LOG, encoding="utf-8")
            handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
            logger.addHandler(handler)
        return logger
