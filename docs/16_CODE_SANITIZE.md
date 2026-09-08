# Санация пользовательского Python-кода

Библиотека: `macro-lib/pythonpath/libre_macros_sanitize_lib.py`  
Статус: подключена в `collect_workbooks._merge_pp_compile_postprocess_code` (inline lambda/def)
и в фильтрах имён листов (`libre_macros_sheet_filter_lib.compile_sheet_name_filter` —
«Листы_не_удалять», «удаление_листов» / «скрытие_листов»).

Связанные находки аудита: [SECURITY_AUDIT.md](SECURITY_AUDIT.md) — H-5 (бывш. C-1/H-4), H-2.

---

## Зачем

В постобработке / финале макрос выполняет пользовательский Python двумя путями:

1. **Inline** — колонка B (`lambda` / `def`) или поле `"code"` в JSON колонки C.
2. **file#func** — ссылка вида `functions_pp.py#pp_range_zagotovka` (файл на диске → вырезание `def` → тот же eval/exec).

Текущая «песочница» в макросе — blocklist builtins (`__import__`, `open`, `eval`, …). Её можно обойти через `getattr` / `__subclasses__` и т.п. (см. C-1).

Библиотека санации даёт **статическую проверку AST и путей** до компиляции. Это не полная песочница ОС, а фильтр опасных конструкций и ограничение `file#` каталогом макроса.

---

## Что проверяется

| Объект | API | Что делает |
|--------|-----|------------|
| Текст кода | `lm_sanitize_code` | AST + запрет import/exec, сети, base64/шифроконтента, записи в ФС, dunder |
| Ссылка `path#func` | `lm_sanitize_file_ref` | Разбор, имя функции, путь только под `base_dir`, по умолчанию без abs |
| Файл + функция | `lm_sanitize_load_file_function` | file_ref → чтение → вырезание `def` → `lm_sanitize_code` |
| Builtins | `lm_sanitize_filtered_builtins` | Словарь builtins без blocklist (для будущей подстановки в eval) |

Код **не выполняется**. Результат — `ok` / список `issues` (error / warning).

---

## Политики

| Политика | Константа | Поведение |
|----------|-----------|-----------|
| **standard** (по умолчанию) | `LM_SANITIZE_POLICY_STANDARD` | Запрет import, `eval`/`exec`/`open`/`getattr`/…, dunder вроде `__subclasses__`; эвристики по тексту — warning |
| **strict** | `LM_SANITIZE_POLICY_STRICT` | То же + эвристики `__subclasses__` / `__builtins__` в тексте → error |
| **permissive** | `LM_SANITIZE_POLICY_PERMISSIVE` | `getattr`/`hasattr`/… → warning; прочие опасные вызовы → error |

Рекомендация для будущей интеграции с B/C и file#: **standard**.

---

## Запреты (standard / strict)

**Узлы AST:** `import`, `from … import` (и `exec`/`print`-statement на Py2, если есть в `ast`).

**Песочница / introspection:** `__import__`, `eval`, `exec`, `compile`, `getattr`, `setattr`, dunder вроде `__subclasses__`, …

**Сеть / сокеты (любые протоколы):** имена/вызовы `socket`, `ssl`, `urllib*`, `requests`, `aiohttp`, `ftplib`, `smtplib`, `asyncio` сетевые API, `create_connection`, `urlopen`, `HTTP(S)Connection`, `WebSocket`, константы `AF_INET` / `SOCK_*` (текстовые маркеры), …

**Base64 и шифроконтент:**
- API: `b64encode`/`b64decode`, `binascii`, `Fernet`, `AES`, `encrypt`/`decrypt`, модули `base64`, `cryptography`, `Crypto`, …
- литералы: длинные строки (≥40) похожие на base64/base64url или hex-blob (≥48) → код `ENCODED_BLOB`

**Запись в ФС:** `open`, `.write` / `writelines`, `mkdir`/`makedirs`, `unlink`/`rmtree`, `shutil.*`, `tempfile`, `write_text`/`write_bytes`, UNO `storeAsURL` / `storeToURL`, …

Коды issue: `NETWORK`, `CRYPTO`, `FS_WRITE`, `ENCODED_BLOB`, плюс `HEUR_*` для текстовых маркеров.

Разрешены «безобидные» dunder: `__name__`, `__doc__`, `__qualname__`.

Обычная работа с ячейками Calc (`getCellByPosition`, `setString`, …) этими правилами не запрещается.
---

## file#func — пути

По умолчанию (`allow_absolute=False`):

- абсолютный путь → ошибка `ABS_PATH` / `PATH_DENIED`;
- сегменты `..`, выход за `base_dir` → отказ;
- расширение должно быть `.py`;
- имя функции — идентификатор (ASCII или кириллица).

Это закрывает рекомендацию аудита **H-2** (сейчас макрос принимает любой abs path).

Опционально: `allow_absolute=True` + `allowed_roots=[...]` для контролируемых корней.

---

## Примеры вызова (вне макроса)

```python
import libre_macros_sanitize_lib as san

# Inline из колонки B
r = san.lm_sanitize_code(
    "lambda doc, sheet, dr, hr: sheet.getCellByPosition(0, 0)",
    policy=san.LM_SANITIZE_POLICY_STANDARD,
)
assert r.ok

# Обход песочницы — не пройдёт
bad = san.lm_sanitize_code("lambda: ().__class__.__bases__[0].__subclasses__()")
assert not bad.ok

# Сеть / base64 / запись на диск — не пройдут
assert not san.lm_sanitize_code("def f():\n    return socket.socket()").ok
assert not san.lm_sanitize_code("def f():\n    return b64decode('YQ==')").ok
assert not san.lm_sanitize_code(
    "def f():\n    open('/tmp/x','w').write('x')"
).ok
blob = "A" * 48
assert not san.lm_sanitize_code("def f():\n    return %r" % blob).ok
print(san.lm_sanitize_format_issues(bad.issues))

# file#
ref = san.lm_sanitize_file_ref(
    "functions_pp.py#pp_range_zagotovka",
    base_dir="/path/to/Scripts/python",
)
assert ref.ok and ref.abs_path

# Файл + санация тела def
loaded = san.lm_sanitize_load_file_function(
    "functions_pp.py#pp_range_zagotovka",
    base_dir="/path/to/Scripts/python",
)
assert loaded["ok"]
```

Исключения: `LmSanitizeError` через `lm_sanitize_code_or_raise` / `lm_sanitize_file_ref_or_raise`.

---

## Примеры «плохих» функций в macro-lib

В конце `functions_pp.py` и `functions_final.py` лежат заготовки **только для теста/примера** санации (не подключать в лист параметров):

| Функция | Что должно баниться |
|---------|---------------------|
| `pp_sanitize_test_network_socket` | `NETWORK` (socket) |
| `pp_sanitize_test_network_urlopen` | `NETWORK` (urlopen) |
| `pp_sanitize_test_base64_blob` | `CRYPTO` / `ENCODED_BLOB` |
| `pp_sanitize_test_fs_write` | `FS_WRITE` |
| `pp_sanitize_test_sandbox_escape` | `BANNED_ATTR` (`__subclasses__`) |

Проверка без макроса:

```bash
PYTHONPATH=macro-lib/pythonpath python3 -c "
import libre_macros_sanitize_lib as san
r = san.lm_sanitize_load_file_function(
    'functions_pp.py#pp_sanitize_test_fs_write', base_dir='macro-lib')
assert not r['ok']
print(san.lm_sanitize_format_issues(r['issues']))
"
```

---

## Планируемая интеграция (ещё не сделано)

Точки в `collect_workbooks.py`:

1. `_merge_pp_compile_postprocess_code` — перед `eval`/`exec` вызвать `lm_sanitize_code`; при `not ok` — лог и `None`.
2. `_merge_pp_resolve_script_path` / `_merge_pp_load_function_from_file_ref` — заменить/обернуть `lm_sanitize_file_ref` + `lm_sanitize_load_file_function` (запрет abs вне каталога макроса).
3. Опционально: собрать builtins через `lm_sanitize_filtered_builtins` вместо локального `_merge_pp_postprocess_eval_builtins`.

После подключения: обновить [07_USER_FUNCTIONS.md](07_USER_FUNCTIONS.md), `functions_pp.py` (справка про blocklist) и строку статуса в [SECURITY_AUDIT.md](SECURITY_AUDIT.md).

**Не путать** с `_lm_safe_eval_*` / `_merge_safe_eval_*` — там whitelist символов для коротких арифметических выражений `{row_id±N}`, не пользовательский Python постобработки.

---

## Ограничения

- Статический анализ **не заменяет** изоляцию процесса: доверенный код с разрешённым API макроса (`merge_pp_*`, UNO) по-прежнему может менять книгу и писать в лог.
- Запись в ячейки книги через UNO **разрешена**; запрещена запись в **файловую систему** (`open`/`write`/shutil/tempfile/`storeAsURL`).
- Обход через динамические строки (`getattr(obj, "__" + "class__")`) ловится лишь частично (эвристики / запрет `getattr`).
- Детектор base64/hex-литералов может дать ложное срабатывание на длинные «похожие» строки (порог ≥40 / ≥48 символов).
- Автотесты по правилам проекта не добавляются; проверка — вручную / `python3 -c` импорт и вызовы API.

---

## Версия

`MACRO_VERSION` в модуле синхронизируется с остальными файлами `pythonpath/` при релизах.
