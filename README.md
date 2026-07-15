# PF BP-PY-ZY — проверки выгрузок макроса

Python-порт логики отчётов **Microsoft Access** (`q_*_Joined_1_Checks`, `q_*_Joined_3_BillToByINN`) для сверки привязок **BP / PY / ZY** по клиентам SAP (SOrg 3801–3806).

Репозиторий: [github.com/k1tit/final-check](https://github.com/k1tit/final-check)

---

## Что делает проект

1. Читает **нулевые выгрузки макроса** (Excel) по Sales Org: Base, BP, PY, ZY.
2. Применяет ту же цепочку, что в Access:
   - дедупликация Base;
   - исключения (`Exception.xlsx`);
   - join BP / PY / ZY;
   - фильтры (CGrp, OrBlk, маски в имени, тестовые ИНН);
   - проверки MD (комментарии аналитика);
   - Bill-to по ИНН.
3. Формирует **итоговые Excel-отчёты** с листами:
   - **«380X Несоответствия»** — найденные ошибки;
   - **«380X Привязка Bill-to по ИНН»**;
   - **Exception** — список исключений.
4. Сохраняет результат в `data/result/`.

Отчёты собираются **парами SOrg**:

| Пара       | SOrg        |
|------------|-------------|
| 3801_3803  | 3801, 3803  |
| 3802_3804  | 3802, 3804  |
| 3805_3806  | 3805, 3806  |

---

## Структура данных

Папка `data` **не входит в git** — её нужно положить вручную рядом со скриптами или указать путь в конфиге.

```
data/
  1 Нулевые файлы выгрузки макроса + файл исключений/
    3801/   *Base*.xlsx, *BP*.xlsx, *PY*.xlsx, *ZY*.xlsx
    3802/
    ...
    3806/
  Exception.xlsx          # исключения (SO + Customer)
  result/                 # сюда пишутся отчёты
    3801_3803/
    3802_3804/
    ...
  staging.duckdb          # временная БД на время прогона (создаётся и очищается автоматически)
```

В каждой папке SOrg скрипт берёт **самый новый** файл по дате, если совпадений несколько.

---

## Файлы проекта

| Файл | Назначение |
|------|------------|
| **`new_access_pf_checks.py`** | Основной скрипт: логика **1:1 с Access**, отчёт **парами**. |
| **`build_checks.py`** | Общая библиотека: чтение Excel, пути, оформление листов, исключения. |
| **`staging_db.py`** | Загрузка xlsx в **DuckDB** на время прогона. |
| **`parallel_io.py`** | Параллельная обработка SOrg. |
| **`check_data_files.py`** | Диагностика: какие xlsx читаются, где `BadZipFile`. |
| **`run_checks.cmd`** | Быстрый запуск на Windows. |

---

## Установка

```powershell
git clone https://github.com/k1tit/final-check.git
cd final-check
python -m pip install -r requirements.txt
```

**Windows:** для проблемных xlsx (часто 3804 BP) нужны **Microsoft Excel** и **pywin32** — чтение через COM, если pandas не открывает файл.

---

## Настройка путей

Файл `runtime_paths.json` в корне проекта:

```json
{
  "data_dir": "data"
}
```

Можно указать абсолютный путь к папке `data` на другой машине:

```json
{
  "data_dir": "C:\\Users\\user\\final-check\\data"
}
```

Из `data_dir` автоматически выводятся:

- `base_dir` → `data/1 Нулевые файлы выгрузки макроса + файл исключений`
- `output_dir` → `data/result`
- `exception_file` → `data/Exception.xlsx`

Переопределение через переменные окружения: `REPORTS_DATA_DIR`, `REPORTS_BASE_DIR`, `REPORTS_OUTPUT_DIR`.

---

## Запуск

```powershell
python new_access_pf_checks.py
```

или двойной клик по **`run_checks.cmd`** — откроется **консольное меню**:

```
=== PF BP-PY-ZY — выбор режима ===
  1  Все 3 пары (3801_3803, 3802_3804, 3805_3806) — 3 файла
  5  Одна пара — 1 файл (выбор: 3801_3803 / 3802_3804 / 3805_3806)
  2  Все SOrg по отдельности
  3  Выбранные SOrg по отдельности
  4  Своя группа SOrg в одном файле
  0  Выход
```

Далее — настройки с пояснениями «зачем» и «плюсы»:

**DuckDB staging** — временная БД для xlsx: меньше повторного чтения, обход BadZipFile (3804 BP), один раз загрузить все 6 SOrg. Минус: долгий первый этап. Рекомендация: **Y**.

**Параллельная обработка** — обе SOrg в паре одновременно (~×2 быстрее). Минус: больше RAM. Рекомендация: **Y**.

**Workers** — сколько SOrg параллельно в задаче: **0** = авто (для пары обычно 2), **1** = экономия RAM, **2+** = для больших групп SOrg.

**Без меню** (для скриптов/автоматизации):

```powershell
python new_access_pf_checks.py --no-menu --mode pairs
python new_access_pf_checks.py --no-menu --mode one_pair --folders 3802_3804
python new_access_pf_checks.py --no-menu --mode single
python new_access_pf_checks.py --no-menu --mode custom_group --folders 3801,3803
```

Флаги:

| Флаг | Описание |
|------|----------|
| `--no-menu` | Не показывать меню |
| `--mode pairs\|one_pair\|single\|custom_single\|custom_group` | Режим работы |
| `--folders 3802_3804` | Имя пары (one_pair) или фильтр SOrg |
| `--no-staging` | Не использовать DuckDB, читать xlsx напрямую |
| `--no-parallel` | Последовательная обработка |
| `--workers N` | Число параллельных SOrg в паре |

---

## Как устроен прогон (DuckDB staging)

По умолчанию:

1. **Старт** — все нужные SOrg (3801…3806) читаются из xlsx и заливаются во временную БД `data/staging.duckdb` (таблицы `so_3804_base`, `so_3804_bp`, …).
2. **Проверки** — данные берутся из DuckDB (xlsx повторно не читаются).
3. **Конец** — staging-таблицы удаляются (`finally`), файл БД остаётся пустым.

Если pandas выдаёт `BadZipFile`, а Excel файл открывает — на Windows срабатывает **fallback через Excel COM**.

Отключить staging: `--no-staging`.

---

## Результат

Пример для пары `3802_3804`:

```
data/result/3802_3804/
  Check PF BP-PY-ZY 3802_3804 10.06.2026.xlsx
```

Внутри книги:

- `3802 Несоответствия`, `3804 Несоответствия`
- `3802 Привязка Bill-to по ИНН`, `3804 Привязка Bill-to по ИНН`
- `Exception`

### Оформление листов Excel

- **Несоответствия** — цветовые группы колонок (Customer/BP/PY/ZY/Checks), автоширина столбцов.
- **Bill-to** — olive / aqua / жёлтый по группам колонок.
- **Exception** — жёлтая шапка.

---

## Диагностика

```powershell
python check_data_files.py
```

Покажет, какие файлы в папках SOrg читаются как xlsx, а какие битые/не того формата.

Типичные проблемы:

| Симптом | Что проверить |
|---------|----------------|
| `Нет каталога: ...base_dir` | `runtime_paths.json` → неверный `data_dir` |
| `BadZipFile` на 3804 BP | Файл на диске; открыть в Excel; перевыгрузить макросом; staging + COM |
| Пустые листы при ненулевом логе | Обновить код (`git pull`) — фильтрация по `_folder` |
| `ModuleNotFoundError: duckdb` | `python -m pip install -r requirements.txt` тем же Python, что запускает скрипт |

---

## Стек

Python, pandas, numpy, openpyxl, xlrd, duckdb, pywin32 (Windows).

---

## Лицензия и данные

Код в репозитории. Выгрузки макроса и отчёты — локальные данные пользователя, в git не коммитятся (см. `.gitignore`).
