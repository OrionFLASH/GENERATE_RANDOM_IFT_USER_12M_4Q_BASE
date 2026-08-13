"""
Тесты разрешения путей и устойчивого создания каталогов логов.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

# Корень проекта и src в PYTHONPATH
_ROOT = Path(__file__).resolve().parents[2]
_SRC = _ROOT / "src"
if str(_SRC) not in sys.path:
    sys.path.insert(0, str(_SRC))

import main as app  # noqa: E402


class TestResolvePath(unittest.TestCase):
    """Проверки _resolve_path."""

    def test_relative_path_uses_project_root(self) -> None:
        """Относительный путь привязывается к корню проекта, не к CWD."""
        resolved = app._resolve_path("log")
        expected = (_ROOT / "log").resolve()
        self.assertEqual(resolved, expected)
        self.assertTrue(resolved.is_absolute())

    def test_absolute_path_unchanged(self) -> None:
        """Абсолютный путь не меняется."""
        with tempfile.TemporaryDirectory() as tmp:
            abs_path = Path(tmp) / "custom_log"
            resolved = app._resolve_path(str(abs_path))
            self.assertEqual(resolved, abs_path)

    def test_relative_not_cwd_dependent(self) -> None:
        """Смена CWD не влияет на результат _resolve_path."""
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = os.getcwd()
            try:
                os.chdir(tmp)
                resolved = app._resolve_path("OUT")
            finally:
                os.chdir(old_cwd)
            self.assertEqual(resolved, (_ROOT / "OUT").resolve())


class TestEnsureWritableDir(unittest.TestCase):
    """Проверки _ensure_writable_dir."""

    def test_creates_preferred_dir(self) -> None:
        """Успешное создание предпочтительного каталога."""
        with tempfile.TemporaryDirectory() as tmp:
            preferred = Path(tmp) / "log"
            result = app._ensure_writable_dir(preferred, fallback_subdir="log")
            self.assertEqual(result, preferred)
            self.assertTrue(preferred.is_dir())

    def test_fallback_on_permission_error(self) -> None:
        """При PermissionError используется каталог в home."""
        with tempfile.TemporaryDirectory() as tmp:
            preferred = Path(tmp) / "blocked_log"
            fake_home = Path(tmp) / "home"
            fake_home.mkdir()
            fallback = fake_home / f".{app._APP_NAME_FALLBACK}" / "log"

            real_mkdir = Path.mkdir

            def selective_mkdir(self: Path, *args, **kwargs):  # noqa: ANN001
                if Path(self) == preferred or str(self) == str(preferred):
                    raise PermissionError("simulated")
                return real_mkdir(self, *args, **kwargs)

            with mock.patch.object(Path, "mkdir", selective_mkdir):
                with mock.patch.object(Path, "home", return_value=fake_home):
                    result = app._ensure_writable_dir(preferred, fallback_subdir="log")

            self.assertEqual(result, fallback)
            self.assertTrue(fallback.is_dir())


class TestProjectLoggerPaths(unittest.TestCase):
    """Инициализация ProjectLogger не падает на относительном log."""

    def test_logger_creates_under_project_or_fallback(self) -> None:
        """get_logger с относительным log_dir создаёт файлы логов."""
        with tempfile.TemporaryDirectory() as tmp:
            log_dir = Path(tmp) / "test_log"
            logger = app.get_logger(log_dir=str(log_dir), log_level="DEBUG")
            self.assertIsNotNone(logger)
            logger.info("Тестовое сообщение")
            # Файлы должны появиться в указанном каталоге
            files = list(log_dir.glob("*.log"))
            self.assertGreaterEqual(len(files), 1)

    def test_logger_from_foreign_cwd(self) -> None:
        """Запуск с чужого CWD (как в IDE) не даёт PermissionError на 'log'."""
        with tempfile.TemporaryDirectory() as tmp:
            old_cwd = os.getcwd()
            try:
                os.chdir(tmp)
                # Относительный "log" должен резолвиться в корень проекта
                preferred = app._resolve_path("log")
                self.assertTrue(str(preferred).startswith(str(_ROOT)))
                # Создаём во временном каталоге, чтобы не трогать реальный log/
                logger = app.get_logger(
                    log_dir=str(Path(tmp) / "ide_log"),
                    log_level="INFO",
                )
                logger.info("ok from foreign cwd")
            finally:
                os.chdir(old_cwd)


class TestNormalizeConfigPaths(unittest.TestCase):
    """Нормализация путей в словаре конфига."""

    def test_normalize_makes_absolute(self) -> None:
        """log_dir / output_dir / input_file становятся абсолютными."""
        cfg = {
            "logging": {"log_dir": "log", "log_level": "DEBUG"},
            "output": {"output_dir": "OUT"},
            "loaders": {
                "ORG": {"input_file": "IN/org.csv"},
                "EMPLOYEE": {"source_file": "IN/emp.csv"},
            },
        }
        result = app._normalize_config_paths(cfg)
        self.assertTrue(Path(result["logging"]["log_dir"]).is_absolute())
        self.assertTrue(Path(result["output"]["output_dir"]).is_absolute())
        self.assertTrue(Path(result["loaders"]["ORG"]["input_file"]).is_absolute())
        self.assertTrue(Path(result["loaders"]["EMPLOYEE"]["source_file"]).is_absolute())
        self.assertTrue(result["logging"]["log_dir"].endswith(str(Path("log"))))


if __name__ == "__main__":
    unittest.main()
