# -*- coding: utf-8 -*-
"""
Гейт допуска: OS env ↔ зашифрованная глобальная переменная.

Используется для «Предварительный_скрипт» (фиксированные имена) и
«функция_плагин» (имена в JSON/визарде; пусто → имена по умолчанию).
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.708"
import os
import sys

try:
    unicode
except NameError:
    unicode = str

try:
    import uno
except Exception:
    uno = None

# Имена по умолчанию для «функция_плагин», если allow_env / allow_global не заданы.
PLUGIN_ALLOW_ENV_NAME = u"MERGE_ALLOW_PLUGINS"
PLUGIN_ALLOW_GLOBAL_NAME = u"Merge_Allow_Plugins"
PLUGIN_ALLOW_ERROR_DIALOG_TITLE = u"функция_плагин — невозможно выполнить"


def _log(msg):
    try:
        print(u"[allow_gate] %s" % msg)
    except Exception:
        pass


def fold_allow_global_name(name):
    """
    Нормализация имени глобальной допуска:
    регистр игнорируется; «_», «-» и прочие не-буквы/цифры выкидываются.
    """
    try:
        from libre_macros_pre_shell_cfg import fold_pre_shell_allow_name

        return fold_pre_shell_allow_name(name)
    except Exception:
        s = unicode(name or u"")
        chars = []
        i = 0
        while i < len(s):
            ch = s[i]
            i += 1
            try:
                if ch.isalnum():
                    chars.append(ch)
            except Exception:
                pass
        folded = u"".join(chars)
        try:
            return folded.casefold()
        except Exception:
            return folded.lower()


PLUGIN_ALLOW_GLOBAL_NAME_FOLD = fold_allow_global_name(PLUGIN_ALLOW_GLOBAL_NAME)


def is_plugin_allow_global_name(name):
    """True, если имя — канонический Merge_Allow_Plugins (гибкое написание)."""
    return fold_allow_global_name(name) == PLUGIN_ALLOW_GLOBAL_NAME_FOLD


def is_force_encrypt_allow_global_name(name):
    """Имена допуска pre-shell / plugin — всегда encrypt при сохранении."""
    try:
        from libre_macros_pre_shell_cfg import is_pre_shell_allow_global_name

        if is_pre_shell_allow_global_name(name):
            return True
    except Exception:
        pass
    return is_plugin_allow_global_name(name)


def get_env_value(env_name):
    """Значение OS env по точному имени (strip); пусто → не задано."""
    name = unicode(env_name or u"").strip()
    if name == u"":
        return u""
    try:
        raw = os.environ.get(name)
    except Exception:
        raw = None
    if raw is None:
        return u""
    return unicode(raw).strip()


def _name_casefold(name):
    s = unicode(name or u"").strip()
    try:
        return s.casefold()
    except Exception:
        return s.lower()


def lookup_global_from_runtime_map(global_name, variables_map, global_sheet_key=None, global_file_key=None):
    """
    Значение глобальной переменной только из сегмента «Глобальные_переменные».
    Имя сравнивается casefold. Ручной ввод / файлы игнорируются.
    Возвращает (found: bool, value: unicode).
    """
    want = _name_casefold(global_name)
    if want == u"":
        return False, u""
    sheet_key = unicode(global_sheet_key or u"").strip()
    file_key = unicode(global_file_key or u"").strip()
    if sheet_key == u"" or file_key == u"":
        try:
            from libre_macros_global_settings_lib import (
                GLOBAL_VARIABLES_FILE_KEY,
                GLOBAL_VARIABLES_SHEET_KEY,
            )

            if sheet_key == u"":
                sheet_key = GLOBAL_VARIABLES_SHEET_KEY
            if file_key == u"":
                file_key = GLOBAL_VARIABLES_FILE_KEY
        except Exception:
            sheet_key = sheet_key or u"Глобальные_переменные"
            file_key = file_key or u"Глобальные_переменные"
    vm = variables_map if isinstance(variables_map, dict) else {}
    found = False
    value = u""
    for key, raw in vm.items():
        parts = unicode(key or u"").split(u"~")
        if len(parts) < 3:
            continue
        name = parts[0]
        sheet = parts[1] if len(parts) > 1 else u""
        fkey = parts[2] if len(parts) > 2 else u""
        if sheet != sheet_key or fkey != file_key:
            continue
        if _name_casefold(name) != want:
            continue
        if isinstance(raw, (list, tuple)):
            text = unicode(raw[0] if len(raw) > 0 else u"")
        else:
            text = unicode(raw if raw is not None else u"")
        found = True
        value = text
    return found, value


def global_var_encrypt_ok(global_name, settings=None):
    """
    В JSON глобальных настроек переменная с таким именем (casefold)
    должна быть с encrypt=True и значением lm1:.
    Возвращает (ok, err_or_None).
    """
    want = _name_casefold(global_name)
    if want == u"":
        return False, u"Не задано имя глобальной переменной."
    try:
        from libre_macros_global_settings_lib import (
            global_variables_from_settings,
            load_global_settings,
        )
        from libre_macros_normalize_lib import CRYPTO_PREFIX, _has_tag_prefix
    except Exception as err:
        return False, u"Не удалось проверить шифрование глобальной переменной: %s" % err
    gs = settings
    if gs is None:
        try:
            gs = load_global_settings()
        except Exception as err:
            return False, u"Не удалось загрузить глобальные настройки: %s" % err
    items = global_variables_from_settings(gs) or []
    matched = None
    for item in items:
        if not isinstance(item, dict):
            continue
        if _name_casefold(item.get(u"name")) == want:
            matched = item
    display = unicode(global_name or u"").strip()
    if matched is None:
        return False, (
            u"Не задана глобальная переменная «%s» "
            u"(вкладка «Глобальные» → «Глобальные переменные»)."
            % display
        )
    if not bool(matched.get(u"encrypt")):
        return False, (
            u"Глобальная переменная «%s» должна храниться только в зашифрованном виде "
            u"(галочка «Шифровать», файл ключа)."
            % display
        )
    stored = unicode(matched.get(u"value") if matched.get(u"value") is not None else u"")
    if stored.strip() == u"":
        return False, u"Глобальная переменная «%s» пуста." % display
    if not _has_tag_prefix(stored, CRYPTO_PREFIX):
        return False, (
            u"Глобальная переменная «%s» должна быть сохранена как lm1: "
            u"(шифровать обязательно)."
            % display
        )
    return True, None


def verify_named_allow_gate( env_name, global_name, variables_map=None, subject=u"Операция", global_sheet_key=None, global_file_key=None):
    """
    Общий гейт:
      1) имя env и имя глобальной заданы (непустые);
      2) OS env с этим именем задана и не пуста;
      3) глобальная с encrypt=True и lm1:;
      4) в runtime-карте (только сегмент глобальных) значение совпадает с env.

    Возвращает None при успехе, иначе текст ошибки.
    """
    env_nm = unicode(env_name or u"").strip()
    glob_nm = unicode(global_name or u"").strip()
    subj = unicode(subject or u"Операция").strip() or u"Операция"
    if env_nm == u"":
        return u"%s запрещена: не задано имя переменной среды." % subj
    if glob_nm == u"":
        return u"%s запрещена: не задано имя глобальной переменной." % subj
    env_val = get_env_value(env_nm)
    if env_val == u"":
        return (
            u"%s запрещена: не задана или пуста переменная среды «%s»."
            % (subj, env_nm)
        )
    enc_ok, enc_err = global_var_encrypt_ok(glob_nm)
    if not enc_ok:
        return u"%s запрещена: %s" % (subj, enc_err)

    found, map_val = lookup_global_from_runtime_map(
        glob_nm, variables_map, global_sheet_key=global_sheet_key, global_file_key=global_file_key,
    )
    if not found:
        return (
            u"%s запрещена: в runtime-карте нет глобальной переменной «%s» "
            u"(нужна только из «Глобальные переменные», не ручной ввод и не файл)."
            % (subj, glob_nm)
        )
    if unicode(map_val).strip() == u"":
        return (
            u"%s запрещена: глобальная переменная «%s» пуста."
            % (subj, glob_nm)
        )
    if unicode(map_val) != env_val:
        return (
            u"%s запрещена: значение среды «%s» не совпадает "
            u"с глобальной переменной «%s»."
            % (subj, env_nm, glob_nm)
        )
    return None


def plugin_block_allow_names(block):
    """
    Извлечь (allow_env, allow_global) из блока функция_плагин.
    Пустое поле → имена по умолчанию MERGE_ALLOW_PLUGINS / Merge_Allow_Plugins.
    """
    if not isinstance(block, dict):
        return PLUGIN_ALLOW_ENV_NAME, PLUGIN_ALLOW_GLOBAL_NAME
    env_nm = unicode(
        block.get(u"allow_env")
        or block.get(u"env")
        or block.get(u"среда")
        or block.get(u"allow_env_name")
        or u""
    ).strip()
    glob_nm = unicode(
        block.get(u"allow_global")
        or block.get(u"global")
        or block.get(u"глобальная")
        or block.get(u"allow_global_name")
        or u""
    ).strip()
    if env_nm == u"":
        env_nm = PLUGIN_ALLOW_ENV_NAME
    if glob_nm == u"":
        glob_nm = PLUGIN_ALLOW_GLOBAL_NAME
    return env_nm, glob_nm


def verify_plugin_allow_gate(block, variables_map=None):
    """Гейт для одного JSON-блока функция_плагин (с именами по умолчанию)."""
    env_nm, glob_nm = plugin_block_allow_names(block)
    return verify_named_allow_gate(
        env_nm,
        glob_nm,
        variables_map=variables_map,
        subject=u"функция_плагин",
    )


def show_allow_gate_error_dialog(doc, text, title=None):
    """
    Диалог ошибки гейта: тот же вид, что у предварительного скрипта
    (крупный красный текст, алая полоса-титл).
    Возвращает True, если диалог показан.
    """
    try:
        from libre_macros_pre_shell_lib import show_pre_shell_error_dialog
    except Exception as err:
        _log(u"show_allow_gate_error_dialog import: %s" % err)
        return False
    caption = unicode(title or u"").strip()
    if caption == u"":
        body0 = unicode(text or u"")
        try:
            folded = body0.casefold()
        except Exception:
            folded = body0.lower()
        if u"функция_плагин" in folded:
            caption = PLUGIN_ALLOW_ERROR_DIALOG_TITLE
        else:
            caption = None
    return show_pre_shell_error_dialog(doc, text, title=caption)
