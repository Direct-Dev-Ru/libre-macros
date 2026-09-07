# Установка Python-макросов для AlterOffice ACell

Руководство по установке макросов из репозитория **libre-macros** в **AlterOffice** (форк офисного пакета на базе LibreOffice). Основной редактор таблиц — **ACell** (аналог Calc).

Официальная документация AlterOffice по Python: [раздел Python в базе знаний ALMI Partner](https://support.almi-partner.ru/help/ru/22-python).

| Макрос | Файл | Документация |
|--------|------|--------------|
| Сбор книг | `collect_workbooks.py` | [04_DATA_COLLECTION.md](04_DATA_COLLECTION.md) |
| Оркестратор | `collect_pack` (тот же файл) | [18_ORCHESTRATOR.md](18_ORCHESTRATOR.md) |
| Визард параметров | `param_wizard.py` | [06_PARAM_WIZARD.md](06_PARAM_WIZARD.md) — пресеты в `…/libre-macros/collect_workbooks/merge_param_presets.json` |
| Выбор файла-источника | `pick_source_file.py` | [08_PICK_SOURCE.md](08_PICK_SOURCE.md) |
| Список файлов в папке | `file_list_macro.py` | см. [README.md](../README.md) |
| Контекстное меню ячейки | `context_menu.py` | §5.1 |

---

## 1. Требования

### AlterOffice и PyUNO

AlterOffice поддерживает макросы на **Python** через компонент **PyUNO** (доступ к API офисного пакета). Подробнее: [Быстрый старт](https://support.almi-partner.ru/help/ru/22-python/72-).

| Платформа | Проверка / установка |
|-----------|----------------------|
| **AlterOS (Linux)** | `rpm -qa \| grep libobasis-python` — пакет должен быть установлен; при отсутствии установите компонент Python из состава AlterOffice |
| **Windows** | При установке AlterOffice компоненты для макросов ставятся вместе с пакетом |

### Тест готовности системы

1. Откройте **AText** или **ACell**.
2. **Сервис → Выполнить макрос** (или **Сервис → Макросы → Выполнить макрос…**).
3. В «Библиотека» выберите **Макросы AlterOffice → HelloWorld**, в «Имя макроса» — **HelloWorldPython**.
4. Нажмите **Выполнить**.

Если появилось тестовое сообщение — PyUNO работает.

### Прочее

- Клон или копия репозитория с каталогом `macro-lib/`.
- **AlterOffice 2026+**: встроенная IDE **Pyzo** для редактирования макросов (**Сервис → Макросы → Редактировать макросы…**).

### Безопасность макросов

**Сервис → Параметры → AlterOffice → Безопасность → Безопасность макросов** — для своих скриптов допустим средний уровень с подтверждением или добавление каталога макросов в доверенные.

---

## 2. Куда копируются макросы

Пользовательские Python-скрипты размещаются в **`Scripts/python`** внутри профиля AlterOffice. Пути из [документации AlterOffice](https://support.almi-partner.ru/help/ru/22-python/72-) (в пользовательских установках могут отличаться):

### Пользовательский каталог (рекомендуется)

| ОС | Путь |
|----|------|
| **GNU/Linux** | `/home/USER/.config/alteroffice/5/user/Scripts/python/` |
| **Windows** | `%AppData%\AlterOffice3\4\user\Scripts\python\` |

### Общий каталог (макросы для всех пользователей)

| ОС | Версия | Путь |
|----|--------|------|
| Linux | AlterOffice 3.4 | `/opt/alteroffice3.4/share/Scripts/python/` |
| Linux | AlterOffice 2025 | `/opt/alteroffice2025.3/misc/Scripts/python/` |
| Linux | AlterOffice 2026 | `/opt/alteroffice2026.0/misc/Scripts/python/` |
| Windows | 3.4 | `%ProgramFiles%\AlterOffice\share\Scripts\python\` |
| Windows | 2025 / 2026 | `%ProgramFiles%\AlterOffice\misc\Scripts\python\` |

Если каталога `python` нет — создайте его вручную, **с соблюдением регистра** имён папок.

### Как узнать свой каталог

1. **ACell → Сервис → Параметры → AlterOffice → Пути** — строка «Макросы».
2. Либо откройте профиль пользователя и найдите `user/Scripts/python`.

> **Совместимость с LibreOffice:** скрипт `install.sh` по умолчанию копирует пакет **и** в LibreOffice (`~/.config/libreoffice/4/user/Scripts/python`), **и** в AlterOffice (`~/.config/alteroffice/5/user/Scripts/python`). Пути можно переопределить через `LO_MACROS_DIR` / `AO_MACROS_DIR` (см. §3).

---

## 3. Установка макросов libre-macros

### Установка через MacroInstaller.ods (графический установщик)

В каталоге `installer/` репозитория есть **книга Calc с встроенным макросом установки** — удобный способ поставить макросы без терминала.

| Файл | Назначение |
|------|------------|
| `installer/MacroInstaller.ods` | Установка из **папки на диске** (путь по умолчанию — ячейка A2) |
| `installer/MacroInstaller_bundled.ods` | **Самодостаточный** вариант: все `.py` уже внутри книги, папку выбирать не нужно |

> Для конечных пользователей (только bundled, без терминала и репозитория): [INSTALL_GUIDE_USER.md](INSTALL_GUIDE_USER.md).

#### Подготовка файлов (для разработчиков)

Из каталога `installer/`:

```bash
# обычный установщик (нужна папка macro-lib на диске)
python3 build_ods.py

# самодостаточный установщик (архив macro-lib внутри ODS)
python3 build_ods_bundled.py

# задать версию: записывается в macro-lib/version.txt и в ячейку B2 листа
python3 build_ods.py --version 3.9.3
python3 build_ods_bundled.py --version 3.9.3
```

Параметры сборки:

| Параметр | Описание |
|----------|----------|
| `--version VER` | Записать версию в `macro-lib/version.txt` и отобразить на листе (B2) |
| `--template FILE` | Шаблон ODS (по умолчанию `template.ods`) |
| `--output FILE` | Имя выходного файла |
| `--macro-lib-dir DIR` | Только для `build_ods_bundled.py` — каталог с исходниками макросов |
| `--no-protect` | Только для `build_ods_bundled.py` — не ставить пароль на изменение ODS |
| `--password PWD` | Пароль защиты от изменений (по умолчанию см. ниже) |
| `--no-docs` | Не включать `docs.zip` в ODS |
| `--skip-docx-build` | Не вызывать `docs/build_docx.py` перед упаковкой (только уже собранные `.docx`) |

После сборки `MacroInstaller_bundled.ods` автоматически получает **защиту от изменений** (как в LibreOffice: «открывать только для чтения» + пароль на запись). В `settings.xml` выставляются `LoadReadonly` и `ModifyPasswordInfo` (PBKDF2). Патч текстовый — без пересборки XML через ElementTree (иначе Calc игнорирует настройки).

Пароль при сборке (по убыванию приоритета): `--password` → переменная среды `MACRO_INSTALLER_PASSWORD` → `installer/installer_password.txt` → `macros_libre_123456`. Итоговый пароль сохраняется в `installer/installer_password.txt` (файл в `.gitignore`; образец — `installer_password.txt.example`).

Без `--version` версия читается из кода: `MACRO_VERSION` в `macro-lib/pythonpath/libre_macros_lib.py`, затем `collect_workbooks.py`, `functions_pp.py`; результат записывается в `macro-lib/version.txt`.

Перед упаковкой bundled-установщик во **всех включаемых `.py`** синхронизирует `MACRO_VERSION = "…"` с выбранной версией.

При сборке с документацией (`docs.zip` по умолчанию включён) скрипт **автоматически** вызывает `docs/build_docx.py --all` (нужен `pip install -r docs/requirements-build.txt` или системный `pandoc`). В архив попадают `.md`, `.docx`, `.pdf` и иллюстрации; при установке распаковываются в `…/alter_office_macros/docs/` (рядом с `test/`).

Ручная пересборка Word-копий:

```bash
pip install -r docs/requirements-build.txt
python3 docs/build_docx.py --all    # → docs/docx/*.docx
```

#### Запуск установщика в ACell / LibreOffice Calc

1. Откройте `MacroInstaller.ods` или `MacroInstaller_bundled.ods`.
2. Разрешите выполнение макросов (**Сервис → Параметры → Безопасность макросов**).
3. Нажмите кнопку **«Запуск диалога установки»** на листе  
   *(или **Сервис → Макросы → Выполнить макрос…** → документ → `install_macros` / `install_macros` из `installer_bundled.py`)*.
4. В диалоге отметьте нужные `.py` и нажмите **«Установить выбранные»**.

Макросы копируются в профиль пользователя: `…/user/Scripts/python/` (вместе с `pythonpath/`, если есть).

#### Отличия двух вариантов

**MacroInstaller.ods**

- После запуска открывается **выбор папки** с исходниками.
- Начальный каталог — из ячейки **A2** (поддерживается `%HOME%\Downloads\libre-macros` и аналоги).
- Подходит, если у вас уже есть клон репозитория или распакованный `macro-lib`.
- В диалоге установки всегда предлагаются **`functions_pp.py`** и **`functions_final.py`** (шаблоны без `g_exportedScripts`).

**MacroInstaller_bundled.ods**

- В **A2** указано `_Встроенный_Пакет_` — макросы берутся из **встроенного ZIP** внутри ODS.
- Диалог выбора папки **не показывается** (если архив на месте).
- В архив входят все рабочие `.py` из `macro-lib/`, включая **`functions_pp.py`** и **`functions_final.py`**.
- Документ должен быть **сохранён на диск** (`file://…`); для несохранённой копии сработает запасной вариант с выбором папки.
- Удобно передать один файл коллеге без репозитория.

> **Важно:** пересобирайте ODS скриптами `build_ods*.py` после изменений в `macro-lib`. При сохранении книги из Calc встроенные скрипты и архив должны оставаться в файле; если макрос пропал — возьмите свежесобранный `MacroInstaller*.ods` из репозитория.

---

### Linux (AlterOS и аналоги)

Из корня репозитория:

```bash
# по умолчанию ставит и в LibreOffice, и в AlterOffice:
#   ~/.config/libreoffice/4/user/Scripts/python
#   ~/.config/alteroffice/5/user/Scripts/python
# при необходимости:
#   export LO_MACROS_DIR=...
#   export AO_MACROS_DIR=...

# все макросы
./macro-lib/install.sh

# набор для сбора книг
./macro-lib/install.sh collect_workbooks param_wizard pick_source_file

# один макрос
./macro-lib/install.sh collect_workbooks
```

Скрипт `install.sh`:

- копирует `.py` из `macro-lib/` (и `pythonpath/`) в **оба** каталога назначения — LibreOffice и AlterOffice;
- пути: `LO_MACROS_DIR` (LO) и `AO_MACROS_DIR` (AO), см. значения по умолчанию выше;
- при установке `collect_workbooks` / `param_wizard` автоматически добавляет `functions_pp.py` и `functions_final.py`;
- при существующем файле спрашивает о перезаписи;
- не копирует служебные `generate_*.py`.

После установки **перезапустите ACell** — список макросов обновится.

### Windows

#### Вариант A — копирование вручную

1. Откройте `macro-lib\` в репозитории.
2. Скопируйте нужные `.py` в:

   ```text
   %AppData%\AlterOffice3\4\user\Scripts\python\
   ```

   (создайте `Scripts\python`, если папок нет).

3. Перезапустите ACell.

#### Вариант B — Git Bash

```bash
cd /d/Projects/libre-macros
LO_MACROS_DIR="$APPDATA/AlterOffice3/4/user/Scripts/python" ./macro-lib/install.sh
```

#### Вариант C — PowerShell

```powershell
$dest = "$env:APPDATA\AlterOffice3\4\user\Scripts\python"
New-Item -ItemType Directory -Force -Path $dest
$src  = "D:\Projects\libre-macros\macro-lib"
Copy-Item "$src\collect_workbooks.py", "$src\param_wizard.py", `
          "$src\pick_source_file.py", "$src\file_list_macro.py" $dest
```

### Импорт макроса в документ ODF

Если макрос должен ехать **вместе с книгой** (а не только в профиле пользователя), его встраивают в архив документа: `Scripts/python/…` + правки `META-INF/manifest.xml`. Подробно: [Импорт python макроса в документ](https://support.almi-partner.ru/help/ru/22-python/141-python).

---

## 4. Запуск макросов

### Через диалог

1. Откройте книгу в **ACell** (для `collect_workbooks` — с листом `Параметры_Объединения`).
2. **Сервис → Макросы → Выполнить макрос…**
3. **Мои макросы → Python →** `<имя_файла_без_.py>` **→** `<функция>`
4. **Выполнить**

| Файл | Точка входа | Назначение |
|------|-------------|------------|
| `collect_workbooks` | `collect_workbooks` | Сбор данных по параметрам |
| `param_wizard` | `set_merge_param` | Визард параметров, пресеты, «Файлы-Источники» |
| `pick_source_file` | `select_source_file` | Выбор файла в строку «Файлы-Источники» |
| `file_list_macro` | `analyze_files_dialog` | Отчёт о файлах и папках |
| `context_menu` | `setContextMenu` | Установка пунктов в контекстное меню ячейки (§5.1) |

### Кнопка на листе (элементы управления формы)

Подходит для строки «Файлы-Источники» и листа параметров. См. [Взаимодействие с формами](https://support.almi-partner.ru/help/ru/22-python/158-alteroffice):

1. **Вид → Панели инструментов → Элементы управления формы**
2. Нарисуйте кнопку → правый щелчок → **Элемент управления**
3. Вкладка **События** → **При нажатии** → **Макрос…**
4. Выберите Python-макрос, например `pick_source_file.select_source_file`

---

## 5. Контекстное меню ячейки ACell

Быстрый запуск макросов из контекстного меню ячейки (правый щелчок по ячейке).

### 5.1. Макрос `context_menu.py` (рекомендуется для libre-macros)

В репозитории есть установщик контекстного меню: он сканирует соседние `.py` в `Scripts/python`, находит в них описание `CONTEXT_CELL_MENU` и записывает выбранные пункты в профиль пользователя.

#### Установка файлов

`context_menu.py` должен лежать **рядом** с остальными макросами (тот же каталог `Scripts/python`):

```bash
# по умолчанию — и LibreOffice, и AlterOffice
./macro-lib/install.sh

# или явно набор с context_menu
./macro-lib/install.sh context_menu collect_workbooks param_wizard pick_source_file
```

Скрипт `install.sh` копирует `context_menu.py` вместе с остальными файлами из `macro-lib/`.

#### Однократная регистрация пунктов меню

1. Откройте любую книгу в **ACell**.
2. **Сервис → Макросы → Выполнить макрос…**
3. **Мои макросы → Python → `context_menu` → `setContextMenu`**
4. В диалоге отметьте макросы для установки (по умолчанию отмечены все найденные).
5. Нажмите **Установить** — в отчёте будут **полные пути** к изменённым файлам `cell.xml`.
6. **Перезапустите ACell** — пункты появятся в контекстном меню ячейки.

Повторный запуск **не дублирует** уже установленные пункты: макрос ищет существующие ветки подменю (например, «Питон → Объединение») и добавляет в них новые команды.

#### Куда записывается меню

Файл контекстного меню ячейки таблицы:

| ОС / продукт | Путь (профиль пользователя) |
|--------------|----------------------------|
| **AlterOffice (Linux)** | `~/.config/alteroffice/5/user/config/soffice.cfg/modules/scalc/popupmenu/cell.xml` |
| **LibreOffice (Linux)** | `~/.config/libreoffice/4/user/config/soffice.cfg/modules/scalc/popupmenu/cell.xml` |

> Важно: путь содержит каталог **`user`** (`…/5/user/config/…`, не `…/5/config/…`). Установщик определяет его автоматически из каталога `Scripts/python`.

При установке сканируются также профили flatpak и других вариантов LibreOffice / AlterOffice, если они найдены на компьютере.

#### Какие макросы попадают в список

Установщик **читает исходный код** соседних `.py` (без `import`) и ищет константу `CONTEXT_CELL_MENU`. Сейчас в комплекте libre-macros:

| Файл | Пункт меню |
|------|------------|
| `collect_workbooks.py` | Сбор книг |
| `param_wizard.py` | Визард параметров |
| `pick_source_file.py` | Выбор файла-источника |

Структура подменю задаётся в каждом файле, например:

```python
CONTEXT_CELL_MENU_KEY = "collect_workbooks"
CONTEXT_CELL_MENU = (
    "submenu#Питон",
    "submenu#Объединение",
    "run_function#collect_workbooks;module#collect_workbooks.py;display_name#Сбор книг;order#10",
)
```

Уровни `submenu#…` — вложенные подменю; `run_function#…` — вызываемая функция и подпись пункта.

**Порядок пунктов** в одном подменю задаётся атрибутом `order#N` (меньше число — выше в меню). Пункты из разных `.py` с одной цепочкой `submenu#…` сортируются вместе по `order`. Без `order` пункт ставится в конец (как `10000`). Пример для `PythonMacros`:

| order | Пункт |
|------:|-------|
| 10 | Обработка книг |
| 20 | Мастер параметров |
| 30 | Вкл/выкл параметр |
| 40 | Параметр вверх |
| 50 | Параметр вниз |
| 60 | Выбрать источник |

#### Добавление своего макроса в меню

1. Положите `.py` в тот же `Scripts/python`.
2. Добавьте в файл константы `CONTEXT_CELL_MENU_KEY` и `CONTEXT_CELL_MENU` (формат выше).
3. Снова выполните `context_menu.setContextMenu` и отметьте новый макрос.

Отдельный `setContextMenu` в каждом файле **не нужен** — точка входа одна: `context_menu.setContextMenu`.

#### Типичные проблемы

| Симптом | Решение |
|---------|---------|
| «Не найдено CONTEXT_CELL_MENU» | Установите макросы в `Scripts/python`; проверьте синтаксис константы в `.py` |
| Пункты не видны после установки | Перезапустите ACell; проверьте путь `…/user/config/…/cell.xml` в отчёте макроса |
| Дублирующиеся ветки меню | Удалите старые пункты вручную из `cell.xml` или через настройку интерфейса (§5.2) |
| Пункт «уже есть», но в меню нет | Старый пункт мог быть в другой ветке (например, `Python`); повторите установку — макрос перенесёт в целевую ветку |

### 5.2. Настройка интерфейса вручную

Альтернатива без `context_menu.py` — добавить пункт меню через диалог настройки:

1. **ACell → Вид → Панели инструментов → Настройка…**
2. Вкладка **Контекстные меню** (или **Меню** / **Команды** — зависит от версии AlterOffice).
3. В списке контекстных меню выберите категорию для **таблицы / ячейки** (например, «Таблица», «Ячейки», «Spreadsheet»).
4. Создайте **новую команду** или выберите существующую группу.
5. **Добавить → Макрос… → Мои макросы → Python →** нужная функция.
6. Задайте подпись пункта, например:
   - «Сбор книг» → `collect_workbooks.collect_workbooks`
   - «Параметр…» → `param_wizard.set_merge_param`
   - «Выбрать источник…» → `pick_source_file.select_source_file`
7. **OK** → сохраните настройку (**Сохранить в: AlterOffice** или **ACell**, если предлагается выбор области).

Настройка сохраняется в профиле пользователя и действует для всех книг.

### 5.3. Программная регистрация (Basic + автозапуск)

AlterOffice позволяет **программно** добавить пункт в контекстное меню через библиотеку Basic. Пример из базы знаний: [Макрос для копирования значения ячеек в ACell](https://support.almi-partner.ru/help/ru/23-basic/278-acell) — там макрос `AddMenuContextItem` добавляет команду «Копировать значение ячеек».

**Общая схема:**

1. Импортируйте библиотеку Basic (**Сервис → Макросы → Управление диалогами… → Библиотеки → Импорт…** → `script.xlb`). Импорт библиотек: [Импорт макросов в AlterOffice](https://support.almi-partner.ru/help/ru/23/277).
2. **Один раз** выполните макрос регистрации (`AddMenuContextItem` или аналог).
3. При необходимости привяжите регистрацию к событию **Запуск приложения**: **Вид → Панели инструментов → Настройка… → События → Сохранить в: AlterOffice → Запуск приложения → Макрос…**

**Вызов Python-макроса из Basic** (если пункт меню должен запускать `.py` из `Scripts/python`):

```basic
Sub RunCollectWorkbooks
    Dim oScriptProvider As Object
    Dim oScript As Object
    oScriptProvider = ThisComponent.getScriptProvider()
    oScript = oScriptProvider.getScript( _
        "vnd.sun.star.script:collect_workbooks.py$collect_workbooks?language=Python&location=user")
    oScript.invoke(Array(), Array(), Array())
End Sub
```

Аналогично для `param_wizard.py$set_merge_param` и `pick_source_file.py$select_source_file`. Подробнее: [Вызов Python из Basic](https://support.almi-partner.ru/help/ru/22-python/73-python-basic).

> Для libre-macros предпочтительнее §5.1 (`context_menu.py`). Basic-обёртка — для особых сценариев организации.

### 5.4. Что выбрать

| Способ | Плюсы | Минусы |
|--------|-------|--------|
| **`context_menu.py` (§5.1)** | Один макрос, диалог выбора, общие подменю, несколько `.py` | Нужен перезапуск ACell; правка `cell.xml` в профиле |
| Настройка интерфейса (§5.2) | Не нужен код, Python напрямую | Ручная настройка на каждой рабочей станции |
| Basic + меню (§5.3) | Автозапуск, единообразие в организации | Нужна Basic-библиотека, сложнее сопровождение |
| Кнопка на листе (§4) | Привязка к конкретной книге-шаблону | Только в этой книге |

---

## 6. Обновление и версии

```bash
# перезапись одного макроса (LO + AO)
./macro-lib/install.sh collect_workbooks

# с обновлением MACRO_VERSION в collect_workbooks.py
./macro-lib/install.sh -v 3.9.2 collect_workbooks
```

Версия основного макроса — `MACRO_VERSION` в `libre_macros_lib.py` (копия в `collect_workbooks.py`, `version.txt`). При запуске макроса на листе параметров синхронизируется ячейка «Версия».

---

## 7. Принципы Python-макросов в AlterOffice

Краткая выжимка из [Переход с Basic на Python-UNO](https://support.almi-partner.ru/help/ru/22-python/83-basic-python-uno) и [Функции для работы с документами](https://support.almi-partner.ru/help/ru/22-python/74-).

### Минимальная структура файла

```python
# -*- coding: utf-8 -*-
import uno
import unohelper

def my_macro(*args):
    doc = XSCRIPTCONTEXT.getDocument()
    cell = doc.Sheets[0].getCellByPosition(0, 0)
    cell.setString("Hello from ACell")

g_exportedScripts = (my_macro,)
```

### Многострочные `def` и AlterOffice 2026

В **AlterOffice 2026** (см. `test/pythonscript.py`, `getModuleByUrl`) перед
`compile()` скрипт из `Scripts/python/` **фильтруется построчно**: остаются
`import`/`from` и тела `def`/`class`. Модульные присваивания, многострочные
`def`/`from … import (`, декораторы `@…` и «голые» константы **ломают** загрузку
(`SyntaxError` или `NameError` → «Could not compile file and find macro»).

Обычный `import соседний.py` читает файл **с диска** (без фильтра). Фильтр —
только при запуске макроса из меню / кнопки через Script Framework.

**Полный чеклист разработки:** [.cursor/rules/alteroffice-2026-pythonscript.mdc](../.cursor/rules/alteroffice-2026-pythonscript.mdc).

Кратко:

1. сигнатуры `def` и `from x import a, b` — **в одну строку**;
2. константы / XML / счётчики / мутабельный state — в **`pythonpath/`**, в скрипте только import;
3. defaults аргументов — литералы или имена из import, не из вырезанных присваиваний;
4. новые макросы — тонкая обёртка `def entry(*args): return _lib.entry(*args)` + логика в pythonpath;
5. `CONTEXT_CELL_MENU` можно оставить в entry (сканер читает AST с диска);
6. проверка: `PYTHONPATH=macro-lib/pythonpath python3 macro-lib/aoffice_compat_check.py --tree macro-lib --warn-assigns`;
7. свёртка def: `python3 macro-lib/aoffice_flatten_defs.py --in-place --tree macro-lib`  
   (оба шага уже в `install.sh`).

После правок на Windows AO при странных ошибках очистите `%TEMP%\AlterOfficeScripts\` и переустановите макросы.

### Точки входа: `g_exportedScripts`

В конце файла объявляется кортеж **callable** — только эти функции видны в диалоге «Выполнить макрос». Вспомогательные функции **не** включают в кортеж:

```python
def helper():
    pass

def public_entry(*args):
    helper()

g_exportedScripts = (public_entry,)   # helper в списке не появится
```

### Контекст выполнения

| Объект | Назначение |
|--------|------------|
| `XSCRIPTCONTEXT.getDocument()` | Текущий документ (активное окно ACell) |
| `XSCRIPTCONTEXT.getComponentContext()` | Контекст UNO для `createInstanceWithContext` |
| `XSCRIPTCONTEXT.getDesktop()` | Рабочий стол (все открытые документы) |

**Важно:** `XSCRIPTCONTEXT` определён на уровне **файла макроса** и **недоступен** в модулях, импортированных из скрипта. Передавайте `doc` / `ctx` аргументами или читайте контекст в точке входа.

### Пути к файлам

Для `loadComponentFromURL`, `storeToURL` и т.п. используйте **URL**, не системный путь:

```python
url = uno.systemPathToFileUrl("/home/USER/data/book.ods")
```

На Windows в URL — прямые слэши `/`.

### Импорт соседних модулей

Рядом со скриптом можно создать каталог **`pythonpath/`** — AlterOffice добавляет его в `sys.path` перед запуском. Макросы libre-macros (`param_wizard`) импортируют `collect_workbooks` из того же `Scripts/python/`.

### Диалоги и UI

В Basic есть `CreateUnoDialog` и `MsgBox`; в Python их **нет** — создавайте через сервисы (`com.sun.star.awt.UnoControlDialog`, `Toolkit.createMessageBox` и т.д.). Примеры — в `param_wizard.py`, `file_list_macro.py`.

### Долгие операции и строка состояния

Синхронное обновление `StatusIndicator` в длинном цикле **блокирует интерфейс**. Для прогресса используйте пакетные операции и периодическую отдачу управления UI (см. параметры `MERGE_UI_YIELD_EVERY_ROWS` в `collect_workbooks`) или потоки — [пример в документации](https://support.almi-partner.ru/help/ru/22-python/74-).

### Редактирование

- **AlterOffice 2026+:** IDE **Pyzo** в составе пакета.
- Любой текстовый редактор; кодировка UTF-8: `# -*- coding: utf-8 -*-` в начале файла.
- **Синтаксис Python строгий** — копируйте примеры аккуратно ([Быстрый старт](https://support.almi-partner.ru/help/ru/22-python/72-)).

---

## 8. Ограничения и типичные ошибки

| Ограничение | Пояснение |
|-------------|-----------|
| Запуск только из ACell | Макросы с `import uno` **не работают** при `python3 script.py` в терминале — нужен PyUNO внутри AlterOffice |
| `g_exportedScripts` обязателен | Без кортежа функция не появится в списке макросов |
| Последовательности UNO | Передавайте **кортежи**, не списки Python ([документация](https://support.almi-partner.ru/help/ru/22-python/83-basic-python-uno)) |
| Параметры `out` | Методы с `[out]` возвращают кортеж `(result, out1, out2, …)` |
| Структуры UNO | Все поля структуры задаются сразу, порядок как в IDL |
| Импорт `collect_workbooks` | `param_wizard` подхватывает каталог постобработки из соседнего файла — установите оба макроса |
| Безопасность | При блокировке макросов — настройки доверия (§1) |

---

## 9. Устранение неполадок

| Симптом | Что проверить |
|---------|----------------|
| В списке макросов нет Python | PyUNO установлен (`libobasis-python` на AlterOS); перезапуск ACell |
| Тест HelloWorldPython не работает | Переустановка компонента скриптов AlterOffice |
| Макрос не виден | Файл в `Scripts/python`, расширение `.py`, есть `g_exportedScripts` |
| «Basic runtime error» / `import uno` | Запуск **из ACell**, не из системного Python |
| `collect_workbooks` не находит параметры | Лист `Параметры_Объединения*`; при нескольких — активируйте нужный |
| Пустой список постобработки в визарде | Рядом установлен `collect_workbooks.py` |
| Пункт контекстного меню не появился | Выполните `context_menu.setContextMenu` (§5.1), перезапустите ACell; проверьте путь `…/user/config/…/cell.xml` |
| `No module named 'context_menu'` | Установите `context_menu.py` в `Scripts/python` рядом с макросами |
| Пути к файлам на Windows | `\` или `/`; макросы libre-macros нормализуют пути |

---

## 10. Рекомендуемый набор для сбора книг

```bash
./macro-lib/install.sh collect_workbooks param_wizard pick_source_file context_menu
```

После установки один раз выполните **`context_menu.setContextMenu`** (§5.1) и перезапустите ACell.

Вместе с `collect_workbooks` автоматически копируются **`functions_pp.py`** и **`functions_final.py`** — шаблоны пользовательских функций постобработки и финальной обработки (`functions_pp.py#…`, `functions_final.py#…` в колонке B). Bundled-установщик (`build_ods_bundled.py`) и ODS-установщик (`build_ods.py`) включают их в список установки всегда.

Опционально: `file_list_macro.py` (инвентаризация каталогов).

---

## 11. Полезные ссылки (ALMI Partner)

| Тема | URL |
|------|-----|
| Обзор Python | https://support.almi-partner.ru/help/ru/22-python |
| Быстрый старт, каталоги, Pyzo | https://support.almi-partner.ru/help/ru/22-python/72- |
| Вызов Python из Basic | https://support.almi-partner.ru/help/ru/22-python/73-python-basic |
| Работа с документами, URL, StatusBar | https://support.almi-partner.ru/help/ru/22-python/74- |
| Basic → Python-UNO | https://support.almi-partner.ru/help/ru/22-python/83-basic-python-uno |
| Макрос внутри ODF-документа | https://support.almi-partner.ru/help/ru/22-python/141-python |
| Формы и события элементов | https://support.almi-partner.ru/help/ru/22-python/158-alteroffice |
| Контекстное меню (пример Basic) | https://support.almi-partner.ru/help/ru/23-basic/278-acell |
| Импорт библиотек Basic | https://support.almi-partner.ru/help/ru/23/277 |
| ACell: листы и ячейки (API) | https://support.almi-partner.ru/help/ru/37-acell/126-acell |
