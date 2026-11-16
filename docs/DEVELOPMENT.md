# Разработка

## Окружение

* Python 3.9+ (в CI — 3.11). Для оконной версии нужен `tkinter` (в стандартной
  сборке Python для Windows он есть; на Debian: `apt install python3-tk`).
* Зависимости: `xlrd` (чтение .xls), `openpyxl` (чтение/запись .xlsx).

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements-dev.txt
```

## Тесты

```bash
python -m pytest                   # всё
python -m pytest tests/test_samples.py -v
python -m pytest -k barcode        # по имени
```

332 теста. Они работают **на реальных файлах-образцах из корня репозитория**:

| Файл тестов | Что проверяет |
|---|---|
| `tests/test_textutil.py` | разбор чисел (`1 006,75`), штрихкодов, артикулов, ставок |
| `tests/test_detect.py` | классификацию всех встречающихся заголовков, поиск шапки, итогов |
| `tests/test_readers.py` | формат по магическим байтам, cp1251, повреждённый xls, HTML/CSV, xlsx в .xls |
| `tests/test_samples.py` | по каждому образцу: число позиций, наличие полей, штрихкоды, артикулы, отсутствие расхождений |
| `tests/test_validate.py` | арифметические проверки, режимы НДС, достройку значений, сверку итогов |
| `tests/test_pipeline.py` | сквозной прогон, структура книги, режимы, битые файлы, `--selftest`, столбцы «Сводки» с поставщиком и датой |
| `tests/synthetic.py` | генератор 13 раскладок одного документа (корпус общности) |
| `tests/test_generality.py` | одни и те же данные из любой раскладки: позиции, все 10 полей, арифметика |
| `tests/test_analysis.py` | покрытие полей, доверие, «требует проверки», отчёт `--explain` |
| `tests/test_meta.py` | поставщик/УНП/тип/номер/дата документа, `parse_date`, имя файла, таблица `EXPECTED` по всем 10 образцам |
| `tests/test_cli.py` | `--json` (массив `documents`) и `--diagnose` (хороший, отсутствующий, битый файл) |
| `tests/test_gui_smoke.py` | логика окна на заглушке tkinter: один файл, автоочистка, предпросмотр таблицы |

Важные числа-«замки» (при их изменении нужно осознанно обновить тесты):
301 позиция, 0 ошибок, 0 предупреждений, 140 рассчитанных значений.

Быстрая проверка без pytest:

```bash
python -m konvektor.cli --selftest .          # код возврата 0 = всё сошлось
python -m konvektor.cli --selftest . --json
```

## Сборка .exe

```bash
pip install -r requirements-dev.txt

pyinstaller --noconfirm --onefile --noconsole --name Konvektor \
    --icon assets/konvektor.ico --add-data "assets/konvektor_icon.png;assets" app_gui.py

pyinstaller --noconfirm --onefile --console --name Konvektor-cli \
    --icon assets/konvektor.ico app_cli.py

dist\Konvektor.exe            # оконная версия
dist\Konvektor-cli.exe --selftest .
dist\Konvektor-cli.exe --diagnose "файл.XLS"   # проверка режима разбора
```

На Linux PyInstaller соберёт бинарник, но **не** `.exe` — Windows-сборка делается
в GitHub Actions на `windows-latest` (кросс-компиляции из Linux нет).

### Почему PyInstaller, а не .NET

Приложение целиком на Python: разбор Excel — самая содержательная часть, а
`xlrd`/`openpyxl` дают полный контроль над «грязными» файлами. PyInstaller
заворачивает всё в один `.exe` (~15 МБ), установка не требуется. CI собирает
именно то, что лежит в репозитории, — расхождения между исходниками и `.exe` нет.

## CI/CD (GitHub Actions)

`.github/workflows/build.yml` запускается **автоматически** на:

* любой `push` (в том числе в рабочую ветку) — то есть после каждого изменения;
* `pull_request`;
* ручной запуск (`workflow_dispatch`);
* тег `v*` — дополнительно создаётся релиз с готовыми файлами.

Работы:

1. **test** (Ubuntu): `pip install -r requirements-dev.txt`, `pytest`,
   `python -m konvektor.cli --selftest .`.
2. **build-windows** (Windows): две сборки PyInstaller, проверка
   `Konvektor-cli.exe --version` и `--selftest .` на образцах, артефакт
   `Konvektor-windows` (оба `.exe` + ZIP), релиз на теге.

Забрать результат: **Actions** → нужный запуск → блок *Artifacts* →
`Konvektor-windows`.

### Как выпустить версию

```bash
# обновить __version__ в konvektor/__init__.py и CHANGELOG.md
git tag v1.0.0
git push origin v1.0.0
```

После этого в *Releases* появятся `Konvektor.exe`, `Konvektor-cli.exe`,
`Konvektor-windows.zip`.

## Отладка

* `KONVEKTOR_DEBUG=1 Konvektor-cli.exe файл.xls -o out.xlsx` — печатать traceback
  по нечитаемым файлам.
* `--json` — компактный машинный итог (позиции, ошибки, предупреждения).
* Лист «Сводка» в результате показывает, каким ридером и какие листы прочитаны,
  и все «Замечания по чтению файлов».
* Быстро посмотреть конкретный файл:

  ```bash
  python - <<'PY'
  from konvektor.readers import read_source
  from konvektor.extract import extract_file
  from konvektor.validate import validate_items
  sf = read_source("файл.xls"); print(sf.reader, [s.name for s in sf.sheets])
  r = extract_file(sf); print(len(r.items), r.warnings)
  rep = validate_items(r.items)
  for i in rep.issues: print(i.severity, i.row, i.field, i.message)
  PY
  ```

## Типовые проблемы

| Симптом | Причина / решение |
|---|---|
| `ModuleNotFoundError: tkinter` | Нет пакета Tk (Linux). `apt install python3-tk` или используйте CLI |
| `CompDocError: Workbook corruption` | «Кривой» .xls из 1С. Открывать с `ignore_workbook_corruption=True` (уже сделано в `_read_xls`) |
| Кракозябры в тексте | В файле нет записи CODEPAGE. `_read_xls` перечитывает с `cp1251` |
| `pyinstaller: command not found` | `pip install -r requirements-dev.txt` |
| Артефакт не появился | Сборка упала на шаге тестов — смотрите лог работы `test` |
| Excel не пересчитал ИТОГО | Формулы пересчитываются при открытии; в предпросмотре некоторых программ их нет — это нормально |
