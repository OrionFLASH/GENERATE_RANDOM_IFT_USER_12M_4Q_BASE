"""
Основной модуль проекта GENERATE_RANDOM_IFT_USER_12M_4Q_BASE.

Точка входа в приложение.
"""

import os
from typing import Optional

from src.logger import get_logger


def main() -> None:
    """
    Главная функция приложения.
    
    Инициализирует логгер и выполняет основную логику программы.
    """
    # Инициализация логгера
    log_dir = os.getenv("LOG_DIR", "log")
    log_level = os.getenv("LOG_LEVEL", "DEBUG")
    logger = get_logger(log_dir=log_dir, log_level=log_level)
    
    logger.info("Запуск приложения GENERATE_RANDOM_IFT_USER_12M_4Q_BASE")
    logger.debug("Инициализация основных компонентов приложения")
    
    # Здесь будет основная логика приложения
    logger.info("Приложение успешно запущено")
    logger.debug("Основные компоненты инициализированы")


if __name__ == "__main__":
    main()

