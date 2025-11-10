"""
Основной модуль проекта GENERATE_RANDOM_IFT_USER_12M_4Q_BASE.

Точка входа в приложение.
Весь код проекта находится в этом файле, разделенный на модули.
Использует библиотеки из стандартной установки Anaconda (pandas, openpyxl).
"""

# ============================================================================
# ПАРАМЕТРЫ КОНФИГУРАЦИИ
# ============================================================================

# Настройки приложения
APP_NAME = "GENERATE_RANDOM_IFT_USER_12M_4Q_BASE"  # Имя приложения
DEBUG = True  # Режим отладки

# Параметры логирования
LOG_DIR = "log"  # Директория для хранения логов
LOG_LEVEL = "DEBUG"  # Уровень логирования: INFO или DEBUG

# Общие параметры выходных файлов
OUTPUT_DIR = "OUT"  # Директория для выходных Excel файлов
OUTPUT_FILE_BASE = "result_base"  # Базовое имя выходного Excel файла

# ============================================================================
# КОНФИГУРАЦИЯ ЗАГРУЗЧИКОВ ДАННЫХ
# ============================================================================
# Структура конфигурации для каждого загрузчика данных
# Каждый загрузчик соответствует отдельному листу в Excel файле

LOADER_CONFIG = {
    'ORG': {
        # Входной файл
        'input_file': "IN/SVD_KB_DM_GAMIFICATION_ORG_UNIT_V20 - 2025.08.28.csv",
        
        # Настройки выходного листа
        'sheet_name': 'ORG',  # Имя листа в Excel
        'max_column_width': 100,  # Максимальная ширина колонки
        
        # Фильтры для исключения строк
        'filters': {
            'tb_code_exclude': {'99', '100', '101', '102'},  # Коды ТБ для исключения
            'gosb_code_exclude': {'0', '9038', '9040'}  # Коды ГОСБ для исключения
        },
        
        # Маппинг колонок: исходное_имя -> результирующее_имя
        'column_mapping': {
            'TB_CODE': 'Код ТБ',
            'TB_FULL_NAME': 'Полное ТБ',
            'TB_SHORT_NAME': 'Короткое ТБ',
            'GOSB_CODE': 'Код ГОСБ',
            'GOSB_NAME': 'Полное ГОСБ',
            'GOSB_SHORT_NAME': 'Короткое ГОСБ',
            'ORG_UNIT_CODE': 'Код подразделения'
        }
    }
    
    # Здесь в будущем можно добавить конфигурации для других листов:
    # 'USERS': {
    #     'input_file': "IN/users.csv",
    #     'sheet_name': 'USERS',
    #     'max_column_width': 100,
    #     'filters': {...},
    #     'column_mapping': {...}
    # },
    # 'METRICS': {
    #     'input_file': "IN/metrics.csv",
    #     'sheet_name': 'METRICS',
    #     'max_column_width': 100,
    #     'filters': {...},
    #     'column_mapping': {...}
    # }
}


# ============================================================================
# МОДУЛЬ ЛОГИРОВАНИЯ
# ============================================================================

import logging
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict


class ProjectLogger:
    """
    Класс для управления логированием проекта.
    
    Создает логи с форматом: Уровень_(тема)_годмесяцдень_часминута.log
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
        
        # Создание консольного обработчика для INFO
        self._setup_console_handler(info_formatter)
    
    def _setup_file_handlers(self, debug_formatter: logging.Formatter, info_formatter: logging.Formatter) -> None:
        """
        Настройка файловых обработчиков для логирования.
        
        Args:
            debug_formatter: Форматтер для DEBUG логов
            info_formatter: Форматтер для INFO логов
        """
        now = datetime.now()
        timestamp = now.strftime("%Y%m%d_%H%M")
        
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
    
    def _setup_console_handler(self, info_formatter: logging.Formatter) -> None:
        """
        Настройка консольного обработчика для вывода сообщений уровня INFO.
        
        Args:
            info_formatter: Форматтер для INFO логов
        """
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(info_formatter)
        self.logger.addHandler(console_handler)
    
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
        log_dir: Директория для логов (по умолчанию из параметров конфигурации)
        log_level: Уровень логирования (по умолчанию из параметров конфигурации)
    
    Returns:
        Настроенный объект логгера
    """
    if log_dir is None:
        log_dir = LOG_DIR
    if log_level is None:
        log_level = LOG_LEVEL
    
    logger_instance = ProjectLogger(log_dir=log_dir, log_level=log_level)
    return logger_instance.get_logger()


# ============================================================================
# МОДУЛЬ ЗАГРУЗКИ ОРГАНИЗАЦИОННЫХ ЕДИНИЦ
# ============================================================================

import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


class OrgUnitsLoader:
    """
    Класс для загрузки и обработки организационных единиц.
    
    Загружает данные из CSV, применяет фильтры и сохраняет в Excel.
    Использует библиотеки из стандартной установки Anaconda (pandas, openpyxl).
    """
    
    def __init__(
        self,
        config: Dict,
        output_file_base: str,
        output_dir: str = "OUT",
        logger: Optional[logging.Logger] = None
    ) -> None:
        """
        Инициализация загрузчика организационных единиц.
        
        Args:
            config: Словарь конфигурации из LOADER_CONFIG для данного загрузчика
            output_file_base: Базовое имя выходного Excel файла (без расширения)
            output_dir: Директория для выходных файлов
            logger: Логгер для записи событий
        """
        self.config = config
        self.input_file = Path(config['input_file'])
        self.output_file_base = output_file_base
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)  # Создаем папку, если её нет
        self.logger = logger or logging.getLogger(__name__)
        
        # Маппинг колонок: исходное_имя -> результирующее_имя
        self.column_mapping = config['column_mapping']
        
        # Фильтры для исключения строк
        filters = config.get('filters', {})
        self.tb_code_exclude = filters.get('tb_code_exclude', set())
        self.gosb_code_exclude = filters.get('gosb_code_exclude', set())
        
        # Настройки Excel
        self.sheet_name = config['sheet_name']
        self.max_column_width = config.get('max_column_width', 100)
    
    def load_csv(self) -> pd.DataFrame:
        """
        Загрузка данных из CSV файла.
        
        Returns:
            DataFrame с загруженными данными
            
        Raises:
            FileNotFoundError: Если файл не найден
            ValueError: Если файл пустой или неверный формат
        """
        if not self.input_file.exists():
            error_msg = f"Файл не найден: {self.input_file}"
            self.logger.error(error_msg)
            raise FileNotFoundError(error_msg)
        
        self.logger.info(f"Начало загрузки CSV файла: {self.input_file}")
        self.logger.debug(f"Загрузка данных из файла [class: OrgUnitsLoader | def: load_csv]")
        
        try:
            # Загрузка CSV с разделителем ';'
            df = pd.read_csv(
                self.input_file,
                sep=';',
                encoding='utf-8',
                dtype=str  # Загружаем все как строки для корректной фильтрации
            )
            
            self.logger.info(f"Загружено строк: {len(df)}")
            self.logger.debug(f"Загружено {len(df)} строк, колонок: {len(df.columns)} [class: OrgUnitsLoader | def: load_csv]")
            
            if df.empty:
                raise ValueError("CSV файл пуст")
            
            return df
            
        except Exception as e:
            error_msg = f"Ошибка при загрузке CSV: {str(e)}"
            self.logger.error(error_msg)
            raise
    
    def filter_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Фильтрация данных по заданным критериям.
        
        Исключает строки где:
        - TB_CODE в списке исключений (99, 100, 101, 102)
        - GOSB_CODE в списке исключений (0, 9038, 9040)
        
        Args:
            df: Исходный DataFrame
            
        Returns:
            Отфильтрованный DataFrame
        """
        initial_count = len(df)
        self.logger.debug(f"Начало фильтрации данных. Исходное количество строк: {initial_count} [class: OrgUnitsLoader | def: filter_data]")
        
        # Проверяем наличие необходимых колонок
        required_columns = ['TB_CODE', 'GOSB_CODE']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            error_msg = f"Отсутствуют необходимые колонки: {missing_columns}"
            self.logger.error(error_msg)
            raise ValueError(error_msg)
        
        # Фильтрация по TB_CODE
        df_filtered = df.copy()
        excluded_tb = 0
        if 'TB_CODE' in df.columns:
            count_before = len(df_filtered)
            df_filtered = df_filtered[~df_filtered['TB_CODE'].isin(self.tb_code_exclude)]
            excluded_tb = count_before - len(df_filtered)
            if excluded_tb > 0:
                self.logger.debug(f"Исключено строк по TB_CODE: {excluded_tb} [class: OrgUnitsLoader | def: filter_data]")
        
        # Фильтрация по GOSB_CODE
        excluded_gosb = 0
        if 'GOSB_CODE' in df_filtered.columns:
            count_before = len(df_filtered)
            df_filtered = df_filtered[~df_filtered['GOSB_CODE'].isin(self.gosb_code_exclude)]
            excluded_gosb = count_before - len(df_filtered)
            if excluded_gosb > 0:
                self.logger.debug(f"Исключено строк по GOSB_CODE: {excluded_gosb} [class: OrgUnitsLoader | def: filter_data]")
        
        final_count = len(df_filtered)
        excluded_total = initial_count - final_count
        
        self.logger.info(f"Отфильтровано строк: {excluded_total}, осталось: {final_count}")
        self.logger.debug(f"Фильтрация завершена. Исключено: {excluded_total}, осталось: {final_count} [class: OrgUnitsLoader | def: filter_data]")
        
        return df_filtered
    
    def select_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Выбор и переименование необходимых колонок.
        
        Args:
            df: Исходный DataFrame
            
        Returns:
            DataFrame с выбранными и переименованными колонками
        """
        self.logger.debug(f"Выбор колонок для экспорта [class: OrgUnitsLoader | def: select_columns]")
        
        # Проверяем наличие всех необходимых колонок
        missing_columns = [col for col in self.column_mapping.keys() if col not in df.columns]
        if missing_columns:
            error_msg = f"Отсутствуют необходимые колонки в исходных данных: {missing_columns}"
            self.logger.error(error_msg)
            raise ValueError(error_msg)
        
        # Выбираем только нужные колонки
        selected_df = df[list(self.column_mapping.keys())].copy()
        
        # Переименовываем колонки
        selected_df = selected_df.rename(columns=self.column_mapping)
        
        self.logger.debug(f"Выбрано и переименовано колонок: {len(selected_df.columns)} [class: OrgUnitsLoader | def: select_columns]")
        
        return selected_df
    
    def save_to_excel(self, df: pd.DataFrame) -> str:
        """
        Сохранение данных в Excel файл с настройками форматирования.
        
        Настройки:
        - Первая строка закреплена
        - Автофильтр включен
        - Ширина колонок по содержимому (максимум max_column_width)
        
        Args:
            df: DataFrame для сохранения
            
        Returns:
            Путь к созданному файлу
        """
        # Формируем имя файла с таймштампом
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        output_file = f"{self.output_file_base}_{timestamp}.xlsx"
        output_path = self.output_dir / output_file
        
        self.logger.info(f"Сохранение данных в Excel: {output_file}")
        self.logger.debug(f"Сохранение {len(df)} строк в файл {output_file} [class: OrgUnitsLoader | def: save_to_excel]")
        
        # Сохраняем в Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=self.sheet_name, index=False)
            
            # Получаем объект листа для форматирования
            worksheet = writer.sheets[self.sheet_name]
            
            # Закрепляем первую строку
            worksheet.freeze_panes = 'A2'
            
            # Включаем автофильтр
            worksheet.auto_filter.ref = worksheet.dimensions
            
            # Настраиваем ширину колонок
            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                
                # Находим максимальную длину содержимого в колонке
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                
                # Устанавливаем ширину (минимум 10, максимум max_column_width)
                adjusted_width = min(max(max_length + 2, 10), self.max_column_width)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Делаем первую строку жирной (заголовки)
            header_font = Font(bold=True)
            for cell in worksheet[1]:
                cell.font = header_font
        
        self.logger.info(f"Файл успешно создан: {output_path.absolute()}")
        self.logger.debug(f"Excel файл создан: {output_path.absolute()}, строк: {len(df)} [class: OrgUnitsLoader | def: save_to_excel]")
        
        return str(output_path.absolute())
    
    def process(self) -> str:
        """
        Полный цикл обработки: загрузка, фильтрация, выбор колонок, сохранение.
        
        Returns:
            Путь к созданному Excel файлу
        """
        self.logger.info("Начало обработки организационных единиц")
        self.logger.debug("Запуск полного цикла обработки данных [class: OrgUnitsLoader | def: process]")
        
        # Загрузка
        df = self.load_csv()
        
        # Фильтрация
        df_filtered = self.filter_data(df)
        
        # Выбор колонок
        df_final = self.select_columns(df_filtered)
        
        # Сохранение
        output_file = self.save_to_excel(df_final)
        
        self.logger.info(f"Обработка завершена успешно. Файл: {output_file}")
        self.logger.debug(f"Обработка завершена. Создан файл: {output_file} [class: OrgUnitsLoader | def: process]")
        
        return output_file


def load_org_units(
    config: Dict,
    output_file_base: str = "result_base",
    output_dir: str = "OUT",
    logger: Optional[logging.Logger] = None
) -> str:
    """
    Функция для загрузки и обработки организационных единиц.
    
    Args:
        config: Словарь конфигурации из LOADER_CONFIG для загрузчика
        output_file_base: Базовое имя выходного Excel файла (по умолчанию "result_base")
        output_dir: Директория для выходных файлов (по умолчанию "OUT")
        logger: Логгер для записи событий (опционально)
    
    Returns:
        Путь к созданному Excel файлу
        
    Example:
        >>> output = load_org_units(
        ...     config=LOADER_CONFIG['ORG'],
        ...     output_file_base="result_base",
        ...     output_dir="OUT"
        ... )
        >>> print(f"Создан файл: {output}")
    """
    loader = OrgUnitsLoader(
        config=config,
        output_file_base=output_file_base,
        output_dir=output_dir,
        logger=logger
    )
    return loader.process()


# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main() -> None:
    """
    Главная функция приложения.
    
    Инициализирует логгер и выполняет основную логику программы.
    Все параметры задаются в начале файла в разделе ПАРАМЕТРЫ КОНФИГУРАЦИИ.
    """
    # Инициализация логгера
    logger = get_logger(log_dir=LOG_DIR, log_level=LOG_LEVEL)
    
    logger.info("Запуск приложения GENERATE_RANDOM_IFT_USER_12M_4Q_BASE")
    logger.debug("Инициализация основных компонентов приложения")
    
    # Загрузка и обработка данных для каждого загрузчика из конфигурации
    for loader_name, loader_config in LOADER_CONFIG.items():
        logger.info(f"Обработка загрузчика: {loader_name}")
        logger.debug(f"Конфигурация загрузчика {loader_name}: входной файл={loader_config['input_file']}, лист={loader_config['sheet_name']} [class: main | def: main]")
        
        # Проверка существования входного файла
        input_path = Path(loader_config['input_file'])
        if not input_path.exists():
            error_msg = f"Входной файл не найден для загрузчика {loader_name}: {input_path.absolute()}"
            logger.error(error_msg)
            raise FileNotFoundError(error_msg)
        
        logger.info(f"Обработка файла: {input_path}")
        logger.debug(f"Входной файл: {input_path.absolute()}, базовое имя выходного: {OUTPUT_FILE_BASE}, директория выходных файлов: {OUTPUT_DIR} [class: main | def: main]")
        
        # Загрузка и обработка данных
        try:
            output_file = load_org_units(
                config=loader_config,
                output_file_base=OUTPUT_FILE_BASE,
                output_dir=OUTPUT_DIR,
                logger=logger
            )
            logger.info(f"Обработка загрузчика {loader_name} завершена успешно. Результат: {output_file}")
        except Exception as e:
            error_msg = f"Ошибка при обработке загрузчика {loader_name}: {str(e)}"
            logger.error(error_msg)
            raise


if __name__ == "__main__":
    main()
