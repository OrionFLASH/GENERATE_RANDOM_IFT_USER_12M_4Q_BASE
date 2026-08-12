# ROADMAP

## Текущий статус

- [v] Вынос параметров в `config/` (common / full / employee_append)
- [v] Режим `employee_append`: ORG → дополнение EMPLOEE → CSV/Excel (+ STATS, ORG, USERS)
- [v] Генерация табельных: 11 нулей + 12-й 0/1 + случайный хвост без дублей и подряд
- [v] `org_distribution.per_tb` / `per_tb_gosb`: настраиваемые min/max людей на ТБ и ТБ+ГОСБ
- [v] Документация, коммит и push

## Режимы

| mode | Файл параметров | Действие |
|------|-----------------|----------|
| `full` | `config/full.json` | Полный пайплайн IFT |
| `employee_append` | `config/employee_append.json` | Дополнение списка сотрудников |
