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

## Основные скрипты

| Файл | Назначение |
|------|------------|
| **`new_access_pf_checks.py`** | Основной CLI: логика **1:1 с Access**, отчёт **парами**. Рекомендуется для сверки с эталоном Access. |
| **`build_checks.py`** | Общий движок + расширенные режимы (`single`, `custom_group`) и веб-интерфейс. |
| **`run_checks_web.py`** | Веб-UI (Flask) для запуска проверок из браузера. |
| **`staging_db.py`** | Загрузка xlsx в **DuckDB** на время прогона. |
| **`parallel_io.py`** | Параллельная обработка SOrg. |
| **`check_data_files.py`** | Диагностика: какие xlsx читаются, где `BadZipFile`. |

---

## Установка

```powershell
git clone https://github.com/k1tit/final-check.git
cd final-check
python -m pip install -r requirements.txt
```

Для веб-интерфейса дополнительно:

```powershell
pip install flask waitress
# или: install_web_deps.cmd
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

### CLI (Access-логика, пары)

```powershell
python new_access_pf_checks.py
```

Флаги:

| Флаг | Описание |
|------|----------|
| `--no-staging` | Не использовать DuckDB, читать xlsx напрямую |
| `--no-parallel` | Последовательная обработка |
| `--workers N` | Число параллельных SOrg в паре |

### CLI (build_checks)

```powershell
python build_checks.py --mode pairs
python build_checks.py --mode single
python build_checks.py --mode custom_group --folders 3801,3803
```

### Веб-интерфейс

```powershell
python run_checks_web.py
# или: run_checks_web.cmd / Запуск_отчётов_веб.vbs
```

Откроется `http://127.0.0.1:8765/` — выбор режима, папки data, запуск с прогрессом.

### Служба Windows (опционально)

```powershell
install_windows_service.cmd
```

---

## Как устроен прогон (DuckDB staging)

По умолчанию (`new_access_pf_checks.py` и `build_checks.py`):

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

Дополнительно `build_checks.py` может сохранять:

- `{пара}_ErrorsOnly.xlsx`
- `{пара}_BillToByINN.xlsx`

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

Python, pandas, numpy, openpyxl, xlrd, duckdb, pywin32 (Windows), Flask, Waitress, asyncio, Git.

---

## Лицензия и данные

Код в репозитории. Выгрузки макроса и отчёты — локальные данные пользователя, в git не коммитятся (см. `.gitignore`).
