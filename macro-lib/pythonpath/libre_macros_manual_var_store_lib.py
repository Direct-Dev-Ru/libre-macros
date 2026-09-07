# -*- coding: utf-8 -*-
"""
Персистентные последние значения «Ручной_ввод_в_карту_переменных».

Каталог: ~/.config/libre-macros/collect_workbooks/ (Windows: %APPDATA%/libre-macros/…).
Файл на лист параметров: <safe_sheet_name>.json
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.693"
import json
import os
import re
import sys

try:
    unicode
except NameError:
    unicode = str

STORE_JSON_VER = 1
_UNSAFE_FILE_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]+')


def _log(msg):
    try:
        print("[manual_var_store] %s" % msg)
    except Exception:
        pass


def store_dir():
    """Тот же каталог, что у профиля визарда / consent."""
    try:
        from libre_macros_wizard_profile_lib import wizard_profile_dir

        return wizard_profile_dir()
    except Exception:
        if sys.platform == "win32":
            base = os.environ.get("APPDATA", os.path.expanduser("~"))
        else:
            base = os.environ.get(
                "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
            )
        path = os.path.join(base, "libre-macros", "collect_workbooks")
        try:
            os.makedirs(path, exist_ok=True)
        except OSError:
            pass
        return path


def sanitize_param_sheet_filename(sheet_name):
    """Имя файла JSON по имени листа параметров."""
    s = unicode(sheet_name or u"").strip()
    if s == u"":
        s = u"_unnamed"
    s = _UNSAFE_FILE_RE.sub(u"_", s)
    s = s.strip(u" .")
    if s == u"":
        s = u"_unnamed"
    if len(s) > 120:
        s = s[:120]
    return s + u".json"


def store_path_for_sheet(sheet_name):
    return os.path.join(store_dir(), sanitize_param_sheet_filename(sheet_name))


def _parse_bool(raw, default=False):
    if raw is None:
        return bool(default)
    if isinstance(raw, bool):
        return raw
    s = unicode(raw).strip().casefold()
    if s in (u"1", u"true", u"yes", u"да", u"y", u"on"):
        return True
    if s in (u"0", u"false", u"no", u"нет", u"n", u"off", u""):
        return False
    return bool(default)


def _empty_payload(sheet_name=u""):
    return {
        u"version": STORE_JSON_VER,
        u"sheet": unicode(sheet_name or u""),
        u"values": {},
    }


def load_sheet_store(sheet_name):
    path = store_path_for_sheet(sheet_name)
    if not os.path.isfile(path):
        return _empty_payload(sheet_name)
    try:
        with open(path, "rb") as f:
            raw = f.read()
        data = json.loads(raw.decode("utf-8"))
    except Exception as err:
        _log(u"load %s: %s" % (path, err))
        return _empty_payload(sheet_name)
    if not isinstance(data, dict):
        return _empty_payload(sheet_name)
    values = data.get(u"values")
    if not isinstance(values, dict):
        values = {}
    out = _empty_payload(sheet_name)
    out[u"sheet"] = unicode(data.get(u"sheet") or sheet_name or u"")
    try:
        out[u"version"] = int(data.get(u"version") or STORE_JSON_VER)
    except Exception:
        out[u"version"] = STORE_JSON_VER
    cleaned = {}
    for key, item in values.items():
        nm = unicode(key or u"").strip()
        if nm == u"":
            continue
        if isinstance(item, dict):
            cleaned[nm] = {
                u"text": unicode(
                    item.get(u"text") if item.get(u"text") is not None else u""
                ),
                u"encrypt": _parse_bool(item.get(u"encrypt"), default=False),
            }
        else:
            cleaned[nm] = {u"text": unicode(item if item is not None else u""), u"encrypt": False}
    out[u"values"] = cleaned
    return out


def save_sheet_store(sheet_name, payload):
    path = store_path_for_sheet(sheet_name)
    data = payload if isinstance(payload, dict) else _empty_payload(sheet_name)
    values = data.get(u"values") if isinstance(data.get(u"values"), dict) else {}
    out_vals = {}
    for key, item in values.items():
        nm = unicode(key or u"").strip()
        if nm == u"":
            continue
        if isinstance(item, dict):
            out_vals[nm] = {
                u"text": unicode(
                    item.get(u"text") if item.get(u"text") is not None else u""
                ),
                u"encrypt": bool(item.get(u"encrypt")),
            }
        else:
            out_vals[nm] = {
                u"text": unicode(item if item is not None else u""),
                u"encrypt": False,
            }
    body = {
        u"version": STORE_JSON_VER,
        u"sheet": unicode(sheet_name or u""),
        u"values": out_vals,
    }
    text = json.dumps(body, ensure_ascii=False, indent=2)
    try:
        store_dir()
        with open(path, "wb") as f:
            f.write(text.encode("utf-8"))
        try:
            from libre_macros_global_settings_lib import notify_config_changed

            notify_config_changed()
        except Exception:
            pass
        return True
    except Exception as err:
        _log(u"save %s: %s" % (path, err))
        return False


def _encrypt_text(plain):
    plain = unicode(plain if plain is not None else u"")
    if plain == u"":
        return u""
    try:
        from libre_macros_normalize_lib import op_encrypt
        from libre_macros_source_extra_lib import resolve_encrypt_passwords_key

        return unicode(op_encrypt(plain, key=resolve_encrypt_passwords_key()) or u"")
    except Exception as err:
        _log(u"encrypt: %s" % err)
        return plain


def _decrypt_text(stored):
    stored = unicode(stored if stored is not None else u"")
    if stored == u"":
        return u""
    try:
        from libre_macros_source_extra_lib import _decrypt_source_extra_secret

        return unicode(_decrypt_source_extra_secret(stored) or u"")
    except Exception as err:
        _log(u"decrypt: %s" % err)
        return stored


def get_remembered_value(sheet_name, var_name):
    """
    Прочитать последнее значение переменной для листа.
    Возвращает unicode или None, если записи нет.
    """
    nm = unicode(var_name or u"").strip()
    if nm == u"" or unicode(sheet_name or u"").strip() == u"":
        return None
    payload = load_sheet_store(sheet_name)
    item = (payload.get(u"values") or {}).get(nm)
    if item is None:
        # casefold fallback
        want = nm.casefold()
        for k, v in (payload.get(u"values") or {}).items():
            if unicode(k).casefold() == want:
                item = v
                break
    if not isinstance(item, dict):
        return None
    text = unicode(item.get(u"text") if item.get(u"text") is not None else u"")
    if _parse_bool(item.get(u"encrypt"), default=False):
        return _decrypt_text(text)
    return text


def put_remembered_value(sheet_name, var_name, text, encrypt=False):
    """
    Записать последнее значение. encrypt=True → lm1: тем же ключом, что пароли источника.
    """
    sheet = unicode(sheet_name or u"").strip()
    nm = unicode(var_name or u"").strip()
    if sheet == u"" or nm == u"":
        return False
    payload = load_sheet_store(sheet)
    values = payload.get(u"values")
    if not isinstance(values, dict):
        values = {}
        payload[u"values"] = values
    raw = unicode(text if text is not None else u"")
    do_enc = bool(encrypt)
    stored = _encrypt_text(raw) if do_enc else raw
    values[nm] = {u"text": stored, u"encrypt": do_enc}
    payload[u"sheet"] = sheet
    return save_sheet_store(sheet, payload)


def block_remember_last(block):
    """Опция remember_last (по умолчанию True)."""
    if not isinstance(block, dict):
        return True
    if u"remember_last" in block:
        return _parse_bool(block.get(u"remember_last"), default=True)
    if u"remember" in block:
        return _parse_bool(block.get(u"remember"), default=True)
    if u"persist_last" in block:
        return _parse_bool(block.get(u"persist_last"), default=True)
    return True


def block_encrypt_last(block):
    """Опция encrypt_last (по умолчанию False; только при remember_last)."""
    if not isinstance(block, dict):
        return False
    if not block_remember_last(block):
        return False
    if u"encrypt_last" in block:
        return _parse_bool(block.get(u"encrypt_last"), default=False)
    if u"encrypt_remembered" in block:
        return _parse_bool(block.get(u"encrypt_remembered"), default=False)
    return False


def resolve_dialog_default(block, sheet_name=u""):
    """
    default для диалога: remembered (если remember_last) иначе block.default.
    """
    if not isinstance(block, dict):
        return u""
    name = unicode(block.get(u"name") or u"").strip()
    if block_remember_last(block) and sheet_name and name:
        remembered = get_remembered_value(sheet_name, name)
        if remembered is not None:
            return remembered
    default_raw = block.get(u"default")
    if default_raw is None:
        return u""
    return unicode(default_raw)


def save_dialog_answer(block, sheet_name, text):
    """После OK диалога — при remember_last записать в store."""
    if not isinstance(block, dict):
        return False
    if not block_remember_last(block):
        return False
    name = unicode(block.get(u"name") or u"").strip()
    sheet = unicode(sheet_name or u"").strip()
    if name == u"" or sheet == u"":
        return False
    return put_remembered_value(
        sheet, name, text, encrypt=block_encrypt_last(block)
    )
