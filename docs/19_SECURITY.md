# Безопасность: гейты и доверенные пути

Практическая сводка усилений безопасности макросов libre-macros  
(состояние **`macro-lib/version.txt` → 3.10.711**).

Полный аудит угроз: [SECURITY_AUDIT.md](SECURITY_AUDIT.md).  
Санация Python: [16_CODE_SANITIZE.md](16_CODE_SANITIZE.md).

---

## 1. Модель в двух фразах

1. **Книга параметров = конфигурация** с правами пользователя LibreOffice / AlterOffice.
2. Опасные действия либо **отрезаны гейтом** (pre-shell, плагины, корни источников), либо **фильтруются** (AST-санация B/C) — это не песочница ОС.

Не открывайте чужие `.ods` / `.xltx` с макрос-параметрами, если не доверяете автору.

---

## 2. Что усилили (хронология)

| Версия | Изменение |
|--------|-----------|
| ≤ 3.10.698 | Pre-shell: `shell=False`, confirm, гейт `MERGE_ALLOW_PRE_SCRIPT` ↔ `Merge_Allow_Pre_Scripts` |
| 3.10.701+ | Гейт «функция_плагин» (`MERGE_ALLOW_PLUGINS` / `Merge_Allow_Plugins`) |
| 3.10.709 | **Удалён** весь LDAP/AD (`ad_list_users`, `ldap3_bundled`, `ad-linux/`, …) |
| 3.10.710 | Санация B/C: **fail-closed**; builtins blocklist = AST (`getattr`, …) |
| 3.10.711 | **Доверенные корни** для «Файлы-Источники» (strict по умолчанию) |

Ранее: discard при закрытии источников (C-2), consent/terms, JSON-only в C.

---

## 3. Предварительный_скрипт

- JSON в параметрах: `argv`, опционально `cwd` / `timeout` / `env` / `clean_env`.
- Запуск **только** если:
  1. OS env `MERGE_ALLOW_PRE_SCRIPT` задана и не пуста;
  2. в **глобальных** настройках есть `Merge_Allow_Pre_Scripts` с **encrypt** и значением `lm1:…`;
  3. расшифрованный plaintext **совпадает** с env.
- Иначе сбор останавливается **до** диалога «Внимание!!!».
- Не сохраняйте секрет допуска в шаблоне книги.

Визард: параметр «Предварительный_скрипт», субдиалог в колонке C.

---

## 4. функция_плагин

- Перед сборкой цепочек PP/final: env ↔ encrypted global.
- Имена в JSON (`allow_env` / `allow_global`) или defaults: `MERGE_ALLOW_PLUGINS` / `Merge_Allow_Plugins`.
- При сохранении для канонических имён допуска — **принудительный** encrypt.

---

## 5. Доверенные корни «Файлы-Источники»

### 5.1. Зачем

Чужой шаблон не должен читать `/etc`, чужие home, весь `C:\` через abs-путь или `/*/*.xlsx`.

### 5.2. Режим

| Режим | Как задать | Поведение |
|-------|------------|-----------|
| **strict** (default) | cfg / не задано | Путь и glob только под разрешёнными корнями |
| **off** | env `MERGE_SOURCE_ROOTS_MODE=off` **или** глобальная `Merge_Source_Roots_Mode=off` | Как раньше: любой путь |

Приоритет режима: **env → глобальная →** `SOURCE_ROOTS_DEFAULT_MODE` в cfg.

### 5.3. Корни по умолчанию

Ручной файл (правьте при необходимости):

`macro-lib/pythonpath/libre_macros_source_roots_cfg.py`

| ОС | Defaults |
|----|----------|
| Linux | `/home/{user}`, `/home/TN/{user}` |
| Windows | `C:\Users\{user}`, `S:\{user}`, весь диск `H:\` |
| Обе | Каталог текущей книги параметров (`SOURCE_ROOTS_INCLUDE_BOOK_DIR`) |
| Обе | `SOURCE_ROOTS_EXTRA_FIXED` — доп. фиксированные пути в cfg |

`{user}` — имя пользователя ОС (`USER` / `USERNAME` / `getpass`).

### 5.4. Дополнительные корни (не из книги!)

| Источник | Имя |
|----------|-----|
| OS env | `MERGE_ALLOWED_SOURCE_ROOTS` |
| Глобальные переменные профиля | `Merge_Allowed_Source_Roots` |

Разделители списка: перевод строки, `;`, и `:` **только** если дальше абсолютный путь (`/…` или `X:\…`), чтобы не ломать `C:\Users\…`.

Примеры:

```text
# Linux env
MERGE_ALLOWED_SOURCE_ROOTS=/mnt/reports;/data/share

# Windows env
MERGE_ALLOWED_SOURCE_ROOTS=D:\Trusted;\\fileserver\dept
```

Глобальная переменная — тот же текст в значении (визард глобальных настроек).  
**Не** кладите whitelist в ячейку «Файлы-Источники»: книга тогда расширит сама себе доступ.

### 5.5. Glob

Перед разворотом маски берётся **литеральный префикс** (до первого `*` / `?` / `[` / `**`).  
Если префикс не под корнем — **отказ без обхода ФС**.  
После разворота каждый файл снова проверяется.

| Пример (user=`ivan`) | Итог |
|----------------------|------|
| `/home/ivan/data/a.xlsx` | OK |
| `/home/TN/ivan/*.csv` | OK |
| `H:\share\**\*.xlsx` (Windows) | OK |
| `/etc/passwd` | запрет |
| `/home/other/x.xlsx` | запрет |
| `/*/*.xlsx` | запрет (префикс `/`) |
| `C:\Users\petr\…` | запрет |
| `эта_книга` | всегда OK |

### 5.6. Remote URL

`smb://`, `http://` … в strict по умолчанию **запрещены**  
(`SOURCE_ROOTS_ALLOW_REMOTE = False` в cfg).

### 5.7. Код

| Модуль | Роль |
|--------|------|
| `libre_macros_source_roots_cfg.py` | Константы, шаблоны, имена env/global |
| `libre_macros_source_roots_lib.py` | Сбор корней, проверка пути/glob, сообщения об ошибках |
| `collect_workbooks.expand_source_files` / `parse_collect_settings` | Включение гейта в сбор |

---

## 6. Санация Python (B/C, лямбды)

- Перед `eval`/`exec`: `lm_sanitize_code` (запрет import, сети, FS write, `getattr`/`__subclasses__`, …).
- Если sanitize **недоступен** или код не прошёл — compile возвращает `None` (**fail-closed**).
- Builtins без `__import__` / `open` / `eval` / `getattr` / … (`LM_SANITIZE_BUILTIN_BLOCKLIST`).

Это **не** запрещает легитимную работу с ячейками через API макроса / UNO.

Подробности: [16_CODE_SANITIZE.md](16_CODE_SANITIZE.md).

**Ещё открыто:** абсолютный `file#func` вне каталога макроса (H-2 в аудите).

---

## 7. Чеклист администратора

1. Переустановить макросы после обновления версии.
2. Pre-shell / плагины: настроить env + encrypted global **только** на нужных ПК.
3. Шары вне home/H:: добавить в `MERGE_ALLOWED_SOURCE_ROOTS` или `Merge_Allowed_Source_Roots`.
4. Временно отладить чужие пути: `MERGE_SOURCE_ROOTS_MODE=off` (не оставлять в проде без причины).
5. Не распространять шаблоны книг с секретами допуска и с путями вне политики.
6. При отказе «вне доверенных корней» — смотреть текст ошибки (список корней + как добавить).

---

## 7.1. Runtime-карта: `__ENV_*` / `__OS_*`

При каждом сборе в карту переменных (сегмент «Глобальные_переменные») копируется
**весь** `os.environ` как `__ENV_<имя>` и набор `__OS_*` (пользователь, home, платформа…).

| Риск | Митигация |
|------|-----------|
| Секреты из env попадают в `<<Переменные.__ENV_…>>` | Не подставлять `__ENV_*` в публичные отчёты; не включать полный дамп карты в журнал |
| Утечка через шаблон пути/формулы | Осознанно выбирать имена; для чувствительных сценариев не класть секреты в env процесса LO |
| Коллизия имён с пользовательскими глобальными | Не создавайте глобальные с префиксами `__ENV_` / `__OS_` — seed среды перезапишет их |

Подробности имён: [04_DATA_COLLECTION.md](04_DATA_COLLECTION.md) §6.0.1.  
Код: `libre_macros_runtime_env_lib.py`.

---

## 8. Связанные документы

| Документ | Тема |
|----------|------|
| [SECURITY_AUDIT.md](SECURITY_AUDIT.md) | Полный отчёт, матрица угроз |
| [16_CODE_SANITIZE.md](16_CODE_SANITIZE.md) | AST / builtins / file# |
| [04_DATA_COLLECTION.md](04_DATA_COLLECTION.md) | Сбор, «Файлы-Источники» |
| [15_PARAM_REFERENCE.md](15_PARAM_REFERENCE.md) | Справочник параметров |
| [06_PARAM_WIZARD.md](06_PARAM_WIZARD.md) | Визард, глобальные переменные |
