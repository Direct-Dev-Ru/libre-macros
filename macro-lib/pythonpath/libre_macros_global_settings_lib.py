# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.717"
"""
Глобальные настройки collect_workbooks (JSON рядом с пресетами).

Файл: ~/.config/libre-macros/collect_workbooks/merge_global_settings.json
"""

import hashlib
import json
import os
import re
import shutil
import sys

try:
    unicode
except NameError:
    unicode = str

GLOBAL_SETTINGS_FILE = u"merge_global_settings.json"
GLOBAL_SETTINGS_SUBDIR = u"collect_workbooks"
# v2: дефолт префикса листа параметров collect_params (миграция со старого дефолта).
# v3: settings_mirror_path + global_variables.
# v4: global_variables_key_file + encrypt у переменных.
GLOBAL_SETTINGS_JSON_VER = 4

DEFAULT_DATE_FORMAT = u"DD.MM.YYYY"
DEFAULT_NUMBER_FORMAT = u"# ##0,00"
DEFAULT_FONT_NAME = u"PT Sans"
DEFAULT_FONT_SIZE = 11.0
# Новый дефолт; старый префикс всегда распознаётся (обратная совместимость).
DEFAULT_PARAM_SHEET_PREFIX = u"collect_params"
LEGACY_PARAM_SHEET_PREFIX = u"Параметры_Объединения"
CUSTOM_NUMBER_FORMAT_LABEL = u"— свой формат —"
# Сегменты ключа runtime-карты для seed из глобальных настроек.
GLOBAL_VARIABLES_FILE_KEY = u"Глобальные_переменные"
GLOBAL_VARIABLES_SHEET_KEY = u"Глобальные_переменные"

DATE_FORMAT_CHOICES = (
    u"DD.MM.YYYY",
    u"DD.MM.YYYY NNN",
    u"DD.MM.YYYY DDD",
    u"DD.MM.YY",
    u"DD.MM.YYYY HH:MM",
    u"DD.MM.YYYY HH:MM:SS",
    u"YYYY-MM-DD",
    u"YY-MM-DD",
    u"YYYY-MM-DD HH:MM",
    u"YYYY-MM-DD HH:MM:SS",
    u"DD/MM/YYYY",
    u"DD/MM/YY",
    u"MM/DD/YYYY",
    u"MM/DD/YY",
    u"DD.MMM.YYYY",
    u"D MMMM YYYY",
    u"ДД.ММ.ГГГГ",
    u"ДД.ММ.ГГ",
    u"ДД.ММ.ГГГГ ЧЧ:ММ",
    u"ДД.ММ.ГГГГ ЧЧ:ММ:СС",
    u"ГГГГ-ММ-ДД",
)

NUMBER_FORMAT_CHOICES = (
    u"# ##0,00",
    u"# ##0",
    u"0,00",
    u"0",
    u"0%",
    u"0,00%",
    CUSTOM_NUMBER_FORMAT_LABEL,
)

# Excel-стиль (#,##0.00 / 0.00) — не предлагать в combo: в Calc ru/RU запятая = десятичный.
# Ручной ввод всё ещё допускается; UNO нормализует при применении.


def is_calc_friendly_number_format(fmt):
    """True, если пресет годится для Calc/ods (без Excel-тысячных/десятичной точки)."""
    s = unicode(fmt or u"").strip()
    if s == u"" or s == CUSTOM_NUMBER_FORMAT_LABEL:
        return True
    low = s.casefold()
    if low == u"general":
        return False
    if u"#,##" in s:
        return False
    # 0.00 / #0.00 — Excel decimal point (не E-научный и не дата)
    if re.match(r"^[#0\s,]*\.\d+%?$", s):
        return False
    return True


def calc_number_format_choices(extra=None):
    """Пресеты числовых форматов для визардов (без Excel-стиля)."""
    out = []
    seen = set()
    for item in NUMBER_FORMAT_CHOICES:
        s = unicode(item or u"").strip()
        if s == u"":
            continue
        if not is_calc_friendly_number_format(s) and s != CUSTOM_NUMBER_FORMAT_LABEL:
            continue
        k = s.casefold()
        if k in seen:
            continue
        seen.add(k)
        out.append(s)
    for item in extra or ():
        s = unicode(item or u"").strip()
        if s == u"":
            continue
        k = s.casefold()
        if k in seen:
            continue
        seen.add(k)
        out.append(s)
    return tuple(out)


def _config_base():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    return os.path.join(base, "libre-macros")


def config_root_dir():
    """Корень профиля libre-macros (~/.config/libre-macros)."""
    return _config_base()


def global_settings_dir():
    path = os.path.join(_config_base(), GLOBAL_SETTINGS_SUBDIR)
    try:
        os.makedirs(path, exist_ok=True)
    except TypeError:
        if not os.path.isdir(path):
            try:
                os.makedirs(path)
            except OSError:
                pass
    except OSError:
        pass
    return path


def global_settings_path():
    return os.path.join(global_settings_dir(), GLOBAL_SETTINGS_FILE)


def normalize_settings_mirror_path(raw):
    """Абсолютный путь зеркала; пусто → выключено."""
    s = unicode(raw or u"").strip()
    if s == u"":
        return u""
    try:
        expanded = os.path.expanduser(os.path.expandvars(s))
        return unicode(os.path.normpath(os.path.abspath(expanded)))
    except Exception:
        return s


def normalize_global_variables_key_file(raw):
    """Путь к файлу ключа шифрования глобальных переменных; пусто → выключено."""
    return normalize_settings_mirror_path(raw)


def global_variables_key_password_from_file(path):
    """
    Ключ шифрования = SHA-256 содержимого файла (hex).
    Возвращает (password_hex, error_or_None).
    """
    s = normalize_global_variables_key_file(path)
    if s == u"":
        return u"", u"Не задан файл ключа шифрования."
    if not os.path.isfile(s):
        return u"", u"Файл ключа не найден:\n%s" % s
    try:
        with open(s, "rb") as fh:
            data = fh.read()
    except Exception as err:
        return u"", u"Не удалось прочитать файл ключа:\n%s" % err
    digest = hashlib.sha256(data).hexdigest()
    return unicode(digest), None


def encrypt_global_variable_value(plain, key_file_path):
    """Зашифровать значение (lm1:); (ciphertext, error_or_None)."""
    password, err = global_variables_key_password_from_file(key_file_path)
    if err:
        return u"", err
    try:
        from libre_macros_normalize_lib import op_encrypt

        return unicode(op_encrypt(plain, key=password)), None
    except Exception as exc:
        return u"", u"Ошибка шифрования: %s" % exc


def decrypt_global_variable_value(stored, key_file_path):
    """Расшифровать lm1:-значение; без префикса — как есть. (text, error_or_None)."""
    s = unicode(stored if stored is not None else u"")
    try:
        from libre_macros_normalize_lib import CRYPTO_PREFIX, _has_tag_prefix
    except Exception:
        return s, None
    if not _has_tag_prefix(s, CRYPTO_PREFIX):
        return s, None
    password, err = global_variables_key_password_from_file(key_file_path)
    if err:
        return s, err
    try:
        from libre_macros_normalize_lib import op_decrypt

        return unicode(op_decrypt(s, key=password)), None
    except Exception as exc:
        return s, u"Ошибка расшифровки: %s" % exc


def prepare_global_variables_for_store(items, key_file_path):
    """
    Нормализация + шифрование значений с encrypt=True.
    Merge_Allow_Pre_Scripts / Merge_Allow_Plugins — всегда encrypt=True.
    Возвращает (list_or_None, error_or_None).
    """
    try:
        from libre_macros_allow_gate_lib import is_force_encrypt_allow_global_name
    except Exception:
        try:
            from libre_macros_pre_shell_cfg import is_pre_shell_allow_global_name as is_force_encrypt_allow_global_name
        except Exception:
            is_force_encrypt_allow_global_name = None
    key_path = normalize_global_variables_key_file(key_file_path)
    need_key = False
    for item in items or ():
        if not isinstance(item, dict):
            continue
        name = unicode(item.get("name") or u"").strip()
        encrypt = bool(item.get("encrypt"))
        if is_force_encrypt_allow_global_name is not None and is_force_encrypt_allow_global_name(name):
            encrypt = True
        if encrypt:
            need_key = True
            break
    if need_key and key_path == u"":
        return None, u"Для шифрования переменных укажите файл ключа (SHA-256 содержимого)."
    out = []
    for item in items or ():
        if not isinstance(item, dict):
            continue
        name = unicode(item.get("name") or u"").strip()
        if name == u"" or u"~" in name:
            continue
        encrypt = bool(item.get("encrypt"))
        if is_force_encrypt_allow_global_name is not None and is_force_encrypt_allow_global_name(name):
            encrypt = True
        value = unicode(item.get("value") if item.get("value") is not None else u"")
        if encrypt:
            # Уже lm1: — не шифровать повторно (иначе битый blob).
            try:
                from libre_macros_normalize_lib import CRYPTO_PREFIX, _has_tag_prefix

                already = _has_tag_prefix(value, CRYPTO_PREFIX)
            except Exception:
                already = False
            if not already:
                cipher, err = encrypt_global_variable_value(value, key_path)
                if err:
                    return None, u"Переменная «%s»: %s" % (name, err)
                value = cipher
        out.append({u"name": name, u"value": value, u"encrypt": encrypt})
    return normalize_global_variables(out), None


def global_variables_for_edit(settings=None):
    """
    Список для UI: расшифрованные значения при encrypt=True (если ключ доступен).
    Элемент: {name, value, encrypt}.
    """
    gs = normalize_global_settings(settings or load_global_settings())
    key_path = unicode(gs.get("global_variables_key_file") or u"")
    out = []
    for item in gs.get("global_variables") or []:
        name = unicode(item.get("name") or u"").strip()
        if name == u"":
            continue
        encrypt = bool(item.get("encrypt"))
        value = unicode(item.get("value") if item.get("value") is not None else u"")
        if encrypt:
            plain, _err = decrypt_global_variable_value(value, key_path)
            value = plain
        out.append({u"name": name, u"value": value, u"encrypt": encrypt})
    return out


def resolved_global_variables(settings=None):
    """
    Runtime: [{name, value}] с расшифровкой encrypt-значений.
    При ошибке ключа encrypted-значение пропускается (не кладётся в карту).
    """
    gs = normalize_global_settings(settings or load_global_settings())
    key_path = unicode(gs.get("global_variables_key_file") or u"")
    out = []
    for item in gs.get("global_variables") or []:
        name = unicode(item.get("name") or u"").strip()
        if name == u"":
            continue
        encrypt = bool(item.get("encrypt"))
        value = unicode(item.get("value") if item.get("value") is not None else u"")
        if encrypt:
            plain, err = decrypt_global_variable_value(value, key_path)
            if err:
                print(u"global_variables: %s: %s" % (name, err))
                continue
            value = plain
        out.append({u"name": name, u"value": value})
    return out


def _normalize_global_variable_entry(item):
    if not isinstance(item, dict):
        return None
    name = unicode(item.get("name") or u"").strip()
    if name == u"" or u"~" in name:
        return None
    value = unicode(item.get("value") if item.get("value") is not None else u"")
    encrypt = bool(item.get("encrypt"))
    try:
        from libre_macros_allow_gate_lib import is_force_encrypt_allow_global_name

        if is_force_encrypt_allow_global_name(name):
            encrypt = True
    except Exception:
        try:
            from libre_macros_pre_shell_cfg import is_pre_shell_allow_global_name

            if is_pre_shell_allow_global_name(name):
                encrypt = True
        except Exception:
            pass
    return {u"name": name, u"value": value, u"encrypt": encrypt}


def normalize_global_variables(items):
    """Список [{name, value, encrypt}, …]; дубликаты имён — последнее по casefold."""
    out = []
    seen = {}
    for item in items or ():
        norm = _normalize_global_variable_entry(item)
        if norm is None:
            continue
        try:
            key = norm[u"name"].casefold()
        except Exception:
            key = norm[u"name"].lower()
        if key in seen:
            out[seen[key]] = norm
        else:
            seen[key] = len(out)
            out.append(norm)
    return out


def global_variables_from_settings(settings=None):
    """Сырой список из JSON (значения могут быть lm1: при encrypt)."""
    gs = normalize_global_settings(settings or load_global_settings())
    return list(gs.get("global_variables") or [])


def global_variables_key_file_from_settings(settings=None):
    gs = normalize_global_settings(settings or load_global_settings())
    return unicode(gs.get("global_variables_key_file") or u"")


def save_global_variables(items, settings=None, key_file=None, update_key_file=False):
    """
    Записать global_variables в JSON (остальные поля сохраняются).
    При update_key_file=True пишется и global_variables_key_file (в т.ч. пустая строка).
    По умолчанию путь ключа не трогается.
    Возвращает None при успехе или строку ошибки.
    """
    gs = dict(normalize_global_settings(settings or load_global_settings()))
    gs[u"global_variables"] = normalize_global_variables(items)
    if update_key_file:
        gs[u"global_variables_key_file"] = normalize_global_variables_key_file(key_file)
    return save_global_settings(gs)


def default_global_settings():
    return {
        "version": GLOBAL_SETTINGS_JSON_VER,
        "default_date_format": DEFAULT_DATE_FORMAT,
        "number_format": DEFAULT_NUMBER_FORMAT,
        "font_name": DEFAULT_FONT_NAME,
        "font_size": DEFAULT_FONT_SIZE,
        "param_sheet_prefix": DEFAULT_PARAM_SHEET_PREFIX,
        "suppress_param_action_messages": False,
        "merge_flags": {},
        "settings_mirror_path": u"",
        "global_variables": [],
        "global_variables_key_file": u"",
    }


def normalize_param_sheet_prefix(raw, default=None):
    """
    Префикс имени листа параметров (по умолч. collect_params).
    Пустое / недопустимое → default.
    """
    if default is None:
        default = DEFAULT_PARAM_SHEET_PREFIX
    s = unicode(raw or u"").strip()
    if s == u"":
        return unicode(default)
    # Calc: имя листа без []:*/?\ и не длиннее 31
    bad = u"[]:*/?\\"
    cleaned = []
    for ch in s:
        if ch in bad or ord(ch) < 32:
            continue
        cleaned.append(ch)
    s = u"".join(cleaned).strip()
    if s == u"":
        return unicode(default)
    if len(s) > 31:
        s = s[:31]
    return s


def param_sheet_name_prefixes(current=None):
    """
    Префиксы для распознавания листов параметров.

    Всегда: текущий (настройка/cfg), collect_params, Параметры_Объединения.
    """
    out = []
    try:
        import libre_macros_collect_cfg as ccfg

        cfg_p = unicode(getattr(ccfg, "MERGE_PARAM_SHEET_NAME", u"") or u"").strip()
    except Exception:
        cfg_p = u""
    for p in (
        current,
        cfg_p,
        DEFAULT_PARAM_SHEET_PREFIX,
        LEGACY_PARAM_SHEET_PREFIX,
    ):
        s = unicode(p or u"").strip()
        if s == u"":
            continue
        if s not in out:
            out.append(s)
    return out


def is_param_sheet_name(sheet_name, current_prefix=None):
    """True, если имя листа — лист параметров (текущий или legacy-префикс)."""
    s = unicode(sheet_name or u"").strip()
    if s == u"":
        return False
    try:
        key = s.casefold()
    except Exception:
        key = s.lower()
    for pref in param_sheet_name_prefixes(current_prefix):
        try:
            pk = pref.casefold()
        except Exception:
            pk = pref.lower()
        if key == pk or key.startswith(pk + u"_") or key.startswith(pk):
            return True
    return False


def normalize_global_settings(obj):
    defaults = default_global_settings()
    if obj is None or not isinstance(obj, dict):
        return dict(defaults)
    out = dict(defaults)
    raw_ver = 1
    try:
        raw_ver = int(obj.get("version") or 1)
    except (TypeError, ValueError):
        raw_ver = 1
    dt = unicode(obj.get("default_date_format", u"") or u"").strip()
    if dt != u"":
        out["default_date_format"] = dt
    num = unicode(obj.get("number_format", u"") or u"").strip()
    if num != u"":
        out["number_format"] = num
    font = unicode(obj.get("font_name", u"") or u"").strip()
    if font != u"":
        out["font_name"] = font
    raw_size = obj.get("font_size", None)
    if raw_size is not None and unicode(raw_size).strip() != u"":
        try:
            size = float(unicode(raw_size).replace(u",", u"."))
            if size >= 1.0:
                out["font_size"] = size
        except (TypeError, ValueError):
            pass
    if "param_sheet_prefix" in obj:
        prefix = normalize_param_sheet_prefix(
            obj.get("param_sheet_prefix"), DEFAULT_PARAM_SHEET_PREFIX
        )
        # v1→v2: сохранённый старый дефолт → collect_params (листы Параметры_Объединения* по-прежнему читаются).
        if raw_ver < 2 and prefix == LEGACY_PARAM_SHEET_PREFIX:
            prefix = DEFAULT_PARAM_SHEET_PREFIX
        out["param_sheet_prefix"] = prefix
    if "suppress_param_action_messages" in obj:
        out["suppress_param_action_messages"] = bool(obj.get("suppress_param_action_messages"))
    raw_flags = obj.get("merge_flags")
    if isinstance(raw_flags, dict):
        flags = {}
        for key, val in raw_flags.items():
            k = unicode(key or u"").strip()
            if k == u"":
                continue
            flags[k] = bool(val)
        out["merge_flags"] = flags
    if "settings_mirror_path" in obj:
        out["settings_mirror_path"] = normalize_settings_mirror_path(
            obj.get("settings_mirror_path")
        )
    if "global_variables" in obj:
        out["global_variables"] = normalize_global_variables(obj.get("global_variables"))
    if "global_variables_key_file" in obj:
        out["global_variables_key_file"] = normalize_global_variables_key_file(
            obj.get("global_variables_key_file")
        )
    out["version"] = GLOBAL_SETTINGS_JSON_VER
    return out


def merge_flags_from_settings(settings=None):
    """Словарь merge_flags из глобальных настроек (только явно заданные ключи)."""
    gs = normalize_global_settings(settings or load_global_settings())
    raw = gs.get("merge_flags")
    if not isinstance(raw, dict):
        return {}
    out = {}
    for key, val in raw.items():
        k = unicode(key or u"").strip()
        if k == u"":
            continue
        out[k] = bool(val)
    return out


def load_global_settings():
    path = global_settings_path()
    if not os.path.isfile(path):
        return default_global_settings()
    try:
        with open(path, "rb") as f:
            raw = f.read()
        if not raw:
            return default_global_settings()
        data = json.loads(raw.decode("utf-8"))
    except Exception as err:
        print(u"global_settings: read %s: %s" % (path, err))
        return default_global_settings()
    file_ver = 1
    if isinstance(data, dict):
        try:
            file_ver = int(data.get("version") or 1)
        except (TypeError, ValueError):
            file_ver = 1
    out = normalize_global_settings(data)
    # Один раз переписать JSON после миграции префикса / version.
    if file_ver < GLOBAL_SETTINGS_JSON_VER:
        try:
            _save_global_settings_file(out, notify=False)
        except Exception:
            pass
    return out


def _save_global_settings_file(payload, notify=True):
    """Запись JSON без повторной нормализации/миграции (для load→migrate)."""
    path = global_settings_path()
    global_settings_dir()
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    with open(path, "wb") as f:
        f.write(text.encode("utf-8"))
    if notify:
        try:
            notify_config_changed(settings=payload)
        except Exception:
            pass


def save_global_settings(settings):
    payload = normalize_global_settings(settings)
    try:
        _save_global_settings_file(payload, notify=True)
    except Exception as err:
        return u"Не удалось сохранить глобальные настройки:\n%s" % err
    try:
        apply_runtime_font_settings(payload)
    except Exception:
        pass
    try:
        apply_runtime_param_sheet_prefix(payload)
    except Exception:
        pass
    return None


def _paths_equal(a, b):
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except Exception:
        return unicode(a or u"") == unicode(b or u"")


def _is_path_inside(child, parent):
    """True, если child совпадает с parent или лежит внутри parent."""
    try:
        child_abs = os.path.abspath(child)
        parent_abs = os.path.abspath(parent)
    except Exception:
        return False
    if _paths_equal(child_abs, parent_abs):
        return True
    try:
        common = os.path.commonpath([child_abs, parent_abs])
        return _paths_equal(common, parent_abs)
    except Exception:
        try:
            prefix = parent_abs.rstrip(os.sep) + os.sep
            return child_abs.startswith(prefix) or _paths_equal(child_abs, parent_abs)
        except Exception:
            return False


def sync_settings_mirror(mirror_path=None, settings=None):
    """
    Скопировать содержимое ~/.config/libre-macros → папку зеркала.

    Пустой путь — no-op. Mirror внутри config root — пропуск (защита от рекурсии).
    Возвращает None при успехе/пропуске или строку ошибки (не бросает наружу).
    """
    try:
        if mirror_path is None:
            gs = normalize_global_settings(settings or load_global_settings())
            mirror_path = gs.get("settings_mirror_path") or u""
        dest = normalize_settings_mirror_path(mirror_path)
        if dest == u"":
            return None
        src = config_root_dir()
        if not os.path.isdir(src):
            return None
        if _is_path_inside(dest, src):
            return (
                u"Зеркало настроек: путь не должен совпадать с профилем "
                u"или лежать внутри него:\n%s" % dest
            )
        try:
            os.makedirs(dest)
        except OSError:
            if not os.path.isdir(dest):
                return u"Зеркало настроек: не удалось создать папку:\n%s" % dest
        # Копируем элементы верхнего уровня (файлы и каталоги).
        try:
            names = os.listdir(src)
        except OSError as err:
            return u"Зеркало настроек: не удалось прочитать профиль:\n%s" % err
        for name in names:
            if name in (u".", u".."):
                continue
            src_item = os.path.join(src, name)
            dst_item = os.path.join(dest, name)
            try:
                if os.path.isdir(src_item) and not os.path.islink(src_item):
                    if os.path.isdir(dst_item):
                        shutil.rmtree(dst_item)
                    elif os.path.lexists(dst_item):
                        os.remove(dst_item)
                    shutil.copytree(src_item, dst_item)
                elif os.path.isfile(src_item) or os.path.islink(src_item):
                    if os.path.isdir(dst_item) and not os.path.islink(dst_item):
                        shutil.rmtree(dst_item)
                    shutil.copy2(src_item, dst_item)
            except Exception as err:
                return u"Зеркало настроек: ошибка копирования «%s»:\n%s" % (name, err)
        return None
    except Exception as err:
        return u"Зеркало настроек: %s" % err


def notify_config_changed(settings=None, mirror_path=None):
    """
    После любой записи в ~/.config/libre-macros — синхронизировать зеркало.
    Ошибки не пробрасывает; возвращает строку ошибки или None.
    """
    try:
        return sync_settings_mirror(mirror_path=mirror_path, settings=settings)
    except Exception as err:
        try:
            print(u"notify_config_changed: %s" % err)
        except Exception:
            pass
        return unicode(err)


def resolve_number_format_choice(text):
    """Текст из combo → строка формата Calc (пустая метка «свой» → пусто)."""
    s = unicode(text or u"").strip()
    if s == u"" or s == CUSTOM_NUMBER_FORMAT_LABEL:
        return u""
    return s


def apply_runtime_param_sheet_prefix(settings=None):
    """
    Подставить param_sheet_prefix в MERGE_PARAM_SHEET_NAME (collect + wizard cfg).

    Возвращает актуальный префикс.
    """
    gs = normalize_global_settings(settings or load_global_settings())
    prefix = normalize_param_sheet_prefix(
        gs.get("param_sheet_prefix"), DEFAULT_PARAM_SHEET_PREFIX
    )
    try:
        import libre_macros_collect_cfg as ccfg

        ccfg.MERGE_PARAM_SHEET_NAME = prefix
    except Exception:
        pass
    try:
        import libre_macros_param_wizard_cfg as pwcfg

        pwcfg.MERGE_PARAM_SHEET_NAME = prefix
    except Exception:
        pass
    return prefix


def apply_runtime_font_settings(settings=None):
    """
    Подставить font_name / font_size в runtime libre_macros_lib и direct.

    Возвращает (font_name, font_size).
    """
    gs = normalize_global_settings(settings or load_global_settings())
    name = unicode(gs.get("font_name") or DEFAULT_FONT_NAME).strip() or DEFAULT_FONT_NAME
    try:
        size = float(gs.get("font_size") or DEFAULT_FONT_SIZE)
    except (TypeError, ValueError):
        size = float(DEFAULT_FONT_SIZE)
    if size < 1.0:
        size = float(DEFAULT_FONT_SIZE)
    try:
        import libre_macros_lib as lm

        lm._LM_PP_FONT_NAME = name
        lm._LM_PP_DEFAULT_FONT_SIZE = float(size)
    except Exception:
        pass
    try:
        import libre_macros_direct_lib as direct

        direct._DIRECT_PP_DEFAULT_FONT_NAME = name
        direct._DIRECT_PP_DEFAULT_FONT_SIZE = float(size)
    except Exception:
        pass
    return name, size


def direct_options_from_global_settings(settings=None):
    """Поля для direct_bind_options из глобальных настроек."""
    gs = normalize_global_settings(settings or load_global_settings())
    return {
        "default_date_format": gs.get("default_date_format", DEFAULT_DATE_FORMAT),
        "number_format": gs.get("number_format", DEFAULT_NUMBER_FORMAT),
        "font_name": gs.get("font_name", DEFAULT_FONT_NAME),
        "font_size": gs.get("font_size", DEFAULT_FONT_SIZE),
    }
