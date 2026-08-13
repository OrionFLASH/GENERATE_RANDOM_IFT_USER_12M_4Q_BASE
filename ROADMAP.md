# ROADMAP

## Текущий статус

- [v] Вынос параметров в `config/` (common / full / employee_append)
- [v] Режим `employee_append`: ORG → дополнение EMPLOEE → CSV/Excel (+ STATS, ORG, USERS)
- [v] Генерация табельных: 11 нулей + 12-й 0/1 + случайный хвост без дублей и подряд
- [v] `org_distribution.per_tb` / `per_tb_gosb`: настраиваемые min/max людей на ТБ и ТБ+ГОСБ
- [v] Документация, коммит и push
- [v] Исправление `PermissionError` при создании `log/`: пути от корня проекта, устойчивый mkdir

## Режимы

| mode | Файл параметров | Действие |
|------|-----------------|----------|
| `full` | `config/full.json` | Полный пайплайн IFT |
| `employee_append` | `config/employee_append.json` | Дополнение списка сотрудников |

## Исправление PermissionError (log/)

- [v] Хелпер `_resolve_path`: относительные пути → от корня проекта
- [v] Нормализация `LOG_DIR` / `OUTPUT_DIR` / `input_file` / `source_file` при загрузке конфига
- [v] `ProjectLogger`: `mkdir(parents=True)` + fallback при отказе в доступе
- [v] Тесты в `src/Tests`
- [v] Обновление README
