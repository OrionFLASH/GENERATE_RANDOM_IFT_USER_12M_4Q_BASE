"""
Модуль для настройки системы логирования проекта.

Обеспечивает логирование с двумя уровнями детализации:
- INFO - основные события выполнения
- DEBUG - отладочная и диагностическая информация
"""

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Optional


class ProjectLogger:
    """
    Класс для управления логированием проекта.
    
    Создает логи с форматом: Уровень_(тема)_годмесяцдень_час.log
    DEBUG логи имеют формат: дата время - [уровень] - сообщение [class: <имя класса> | def: <имя функции>]
    """
    
    def __init__(self, log_dir: str = "log", log_level: str = "DEBUG") -> None:
        """
        Инициализация логгера.
        
        Args:
            log_dir: Директория для хранения логов
            log_level: Уровень логирования (INFO или DEBUG)
        """
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        self.log_level = getattr(logging, log_level.upper(), logging.DEBUG)
        
        # Настройка форматирования для DEBUG
        debug_formatter = logging.Formatter(
            '%(asctime)s - [%(levelname)s] - %(message)s [class: %(name)s | def: %(funcName)s]',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Настройка форматирования для INFO
        info_formatter = logging.Formatter(
            '%(asctime)s - [%(levelname)s] - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # Создание логгера
        self.logger = logging.getLogger('project_logger')
        self.logger.setLevel(self.log_level)
        
        # Очистка существующих обработчиков
        self.logger.handlers.clear()
        
        # Создание файловых обработчиков
        self._setup_file_handlers(debug_formatter, info_formatter)
    
    def _setup_file_handlers(self, debug_formatter: logging.Formatter, info_formatter: logging.Formatter) -> None:
        """
        Настройка файловых обработчиков для логирования.
        
        Args:
            debug_formatter: Форматтер для DEBUG логов
            info_formatter: Форматтер для INFO логов
        """
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H")
        
        # Обработчик для DEBUG
        debug_file = self.log_dir / f"DEBUG_project_{timestamp}.log"
        debug_handler = logging.FileHandler(debug_file, encoding='utf-8')
        debug_handler.setLevel(logging.DEBUG)
        debug_handler.setFormatter(debug_formatter)
        self.logger.addHandler(debug_handler)
        
        # Обработчик для INFO
        info_file = self.log_dir / f"INFO_project_{timestamp}.log"
        info_handler = logging.FileHandler(info_file, encoding='utf-8')
        info_handler.setLevel(logging.INFO)
        info_handler.setFormatter(info_formatter)
        self.logger.addHandler(info_handler)
    
    def get_logger(self) -> logging.Logger:
        """
        Получение настроенного логгера.
        
        Returns:
            Настроенный объект логгера
        """
        return self.logger


def get_logger(log_dir: Optional[str] = None, log_level: Optional[str] = None) -> logging.Logger:
    """
    Функция для получения глобального логгера проекта.
    
    Args:
        log_dir: Директория для логов (по умолчанию из переменных окружения)
        log_level: Уровень логирования (по умолчанию из переменных окружения)
    
    Returns:
        Настроенный объект логгера
    """
    if log_dir is None:
        log_dir = os.getenv("LOG_DIR", "log")
    if log_level is None:
        log_level = os.getenv("LOG_LEVEL", "DEBUG")
    
    logger_instance = ProjectLogger(log_dir=log_dir, log_level=log_level)
    return logger_instance.get_logger()

