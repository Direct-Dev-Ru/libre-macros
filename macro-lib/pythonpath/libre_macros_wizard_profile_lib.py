# -*- coding: utf-8 -*-
"""
Персистент визарда: сохранённые лямбды (профиль рядом с пресетами параметров).

Файл: ~/.config/libre-macros/collect_workbooks/merge_wizard_profile.json
"""
from __future__ import print_function, unicode_literals

MACRO_VERSION = "3.10.722"
import json
import os
import sys

try:
    unicode
except NameError:
    unicode = str

WIZARD_PROFILE_FILE = u"merge_wizard_profile.json"
WIZARD_PROFILE_SUBDIR = u"collect_workbooks"
WIZARD_PROFILE_JSON_VER = 1

REPLACE_VALUES_DEFAULT_LAMBDAS = (
    {
        u"name": u"Строка не пустая",
        u"expr": u'lambda rows: str(rows("Столбец") or "").strip() != ""',
    },
    {
        u"name": u"Пустая ячейка",
        u"expr": u'lambda rows: str(rows("Столбец") or "").strip() == ""',
    },
    {
        u"name": u"Первое совпадение",
        u"expr": u"lambda rows: idx == 0",
    },
)


def _config_base():
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    return os.path.join(base, "libre-macros")


def wizard_profile_dir():
    path = os.path.join(_config_base(), WIZARD_PROFILE_SUBDIR)
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        pass
    return path


def wizard_profile_path():
    return os.path.join(wizard_profile_dir(), WIZARD_PROFILE_FILE)


def _normalize_lambda_entry(item):
    if not isinstance(item, dict):
        return None
    name = unicode(item.get("name") or u"").strip()
    expr = unicode(item.get("expr") or item.get("code") or u"").strip()
    if name == u"" or expr == u"":
        return None
    return {u"name": name, u"expr": expr}


def normalize_saved_lambdas(items):
    out = []
    seen = set()
    for item in items or ():
        norm = _normalize_lambda_entry(item)
        if norm is None:
            continue
        key = norm[u"name"].casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(norm)
    return out


def load_wizard_profile():
    path = wizard_profile_path()
    if not os.path.isfile(path):
        return {u"version": WIZARD_PROFILE_JSON_VER, u"saved_lambdas": []}
    try:
        with open(path, "rb") as f:
            raw = f.read()
        data = json.loads(raw.decode("utf-8"))
    except Exception:
        return {u"version": WIZARD_PROFILE_JSON_VER, u"saved_lambdas": []}
    if not isinstance(data, dict):
        return {u"version": WIZARD_PROFILE_JSON_VER, u"saved_lambdas": []}
    lambdas = normalize_saved_lambdas(data.get("saved_lambdas"))
    return {u"version": WIZARD_PROFILE_JSON_VER, u"saved_lambdas": lambdas}


def save_wizard_profile(profile):
    path = wizard_profile_path()
    prof = profile if isinstance(profile, dict) else {}
    payload = {
        u"version": WIZARD_PROFILE_JSON_VER,
        u"saved_lambdas": normalize_saved_lambdas(prof.get("saved_lambdas")),
    }
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    try:
        wizard_profile_dir()
        with open(path, "wb") as f:
            f.write(text.encode("utf-8"))
        try:
            from libre_macros_global_settings_lib import notify_config_changed

            notify_config_changed()
        except Exception:
            pass
        return True, None
    except Exception as err:
        return False, unicode(err)


def load_saved_lambdas():
    return list(load_wizard_profile().get("saved_lambdas") or [])


def ensure_replace_values_default_lambdas():
    """
    При первом открытии визарда замена_значений: если список пуст — записать примеры.
    Возвращает актуальный список [{name, expr}, …].
    """
    prof = load_wizard_profile()
    lambdas = list(prof.get("saved_lambdas") or [])
    if len(lambdas) > 0:
        return lambdas
    lambdas = normalize_saved_lambdas(list(REPLACE_VALUES_DEFAULT_LAMBDAS))
    prof[u"saved_lambdas"] = lambdas
    save_wizard_profile(prof)
    return lambdas


def save_named_lambda(name, expr, lambdas=None):
    """Добавить или перезаписать лямбду по имени. (lambdas, err)."""
    nm = unicode(name or u"").strip()
    ex = unicode(expr or u"").strip()
    if nm in (u"", u"—", u"-", u"--"):
        return lambdas, u"нужно имя"
    if ex == u"":
        return lambdas, u"пустое выражение"
    items = normalize_saved_lambdas(lambdas if lambdas is not None else load_saved_lambdas())
    key = nm.casefold()
    out = []
    replaced = False
    for item in items:
        if item[u"name"].casefold() == key:
            out.append({u"name": nm, u"expr": ex})
            replaced = True
        else:
            out.append(item)
    if not replaced:
        out.append({u"name": nm, u"expr": ex})
    ok, err = save_wizard_profile({u"saved_lambdas": out})
    if not ok:
        return items, err
    return out, None


def delete_saved_lambda(name, lambdas=None):
    """Удалить по имени. (lambdas, err)."""
    nm = unicode(name or u"").strip()
    if nm == u"":
        return lambdas, u"нужно имя"
    key = nm.casefold()
    items = normalize_saved_lambdas(lambdas if lambdas is not None else load_saved_lambdas())
    out = [x for x in items if x[u"name"].casefold() != key]
    if len(out) == len(items):
        return items, u"не найдено: «%s»" % nm
    ok, err = save_wizard_profile({u"saved_lambdas": out})
    if not ok:
        return items, err
    return out, None


def saved_lambda_expr_by_name(name, lambdas=None):
    nm = unicode(name or u"").strip()
    if nm == u"":
        return u""
    key = nm.casefold()
    for item in lambdas if lambdas is not None else load_saved_lambdas():
        if item[u"name"].casefold() == key:
            return unicode(item.get("expr") or u"")
    return u""


def saved_lambda_names(lambdas=None):
    return [
        unicode(x.get("name") or u"")
        for x in (lambdas if lambdas is not None else load_saved_lambdas())
        if unicode(x.get("name") or u"").strip()
    ]
