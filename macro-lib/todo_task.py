# -*- coding: utf-8 -*-
"""
Макросы управления строками таблицы задач (лист с колонками
«№ п/п» … «GUID Строки»).

Точки входа:
  todo_task_add    — добавить строку после последней задачи
  todo_task_delete — удалить текущую строку-задачу
  todo_task_edit   — диалог редактирования текущей задачи
  todo_task_copy   — копировать текущую задачу (без GUID/№, срок очищается)
  todo_task_colorize — раскрасить строки-задачи (диалог: лист / все по шаблону)
  todo_task_sort   — отсортировать задачи (диалог выбора критерия)
  todo_task_move   — переместить/скопировать выбранные задачи на другой лист Задачи…
  todo_task_merge  — слияние задач с листа свода сотрудников (только руководитель)
  todo_task_help   — справка по макросам серии (Writer / Atext)
  todo_task_settings — настройки серии: частные (по пути книги) / глобальные по умолчанию
  todo_task_refs   — редактирование справочников __Справочники_задачи
"""
MACRO_VERSION = "3.10.689"
from todo_task_lib import todo_task_add as _add
from todo_task_lib import todo_task_colorize as _colorize
from todo_task_lib import todo_task_copy as _copy
from todo_task_lib import todo_task_delete as _delete
from todo_task_lib import todo_task_edit as _edit
from todo_task_lib import todo_task_help as _help
from todo_task_lib import todo_task_merge as _merge
from todo_task_lib import todo_task_move as _move
from todo_task_lib import todo_task_refs as _refs
from todo_task_lib import todo_task_settings as _settings
from todo_task_lib import todo_task_sort as _sort


def todo_task_add(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _add(doc)


def todo_task_delete(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _delete(doc)


def todo_task_edit(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _edit(doc)


def todo_task_copy(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _copy(doc)


def todo_task_colorize(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _colorize(doc)


def todo_task_sort(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _sort(doc)


def todo_task_move(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _move(doc)


def todo_task_merge(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _merge(doc)


def todo_task_help(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _help(doc)


def todo_task_settings(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _settings(doc)


def todo_task_refs(*args):
    doc = None
    try:
        doc = XSCRIPTCONTEXT.getDocument()
    except Exception:
        pass
    return _refs(doc)


CONTEXT_CELL_MENU_KEY = "todo_task"
CONTEXT_CELL_MENU = (
    "submenu#PythonMacros",
    "separator#before",
    "run_function#todo_task_add;module#todo_task.py;display_name#Задача: добавить;order#200",
    "run_function#todo_task_edit;module#todo_task.py;display_name#Задача: редактировать;order#210",
    "run_function#todo_task_copy;module#todo_task.py;display_name#Задача: копировать;order#220",
    "run_function#todo_task_move;module#todo_task.py;display_name#Задача: переместить;order#225",
    "run_function#todo_task_delete;module#todo_task.py;display_name#Задача: удалить;order#230",
    "run_function#todo_task_colorize;module#todo_task.py;display_name#Задача: раскрасить;order#240",
    "run_function#todo_task_sort;module#todo_task.py;display_name#Задача: сортировать;order#245",
    "run_function#todo_task_merge;module#todo_task.py;display_name#Задача: слияние с сводом;order#246",
    "run_function#todo_task_refs;module#todo_task.py;display_name#Задача: справочники;order#248",
    "run_function#todo_task_help;module#todo_task.py;display_name#Задача: справка;order#249",
    "run_function#todo_task_settings;module#todo_task.py;display_name#Задача: настройки;order#250",
    "separator#after",
)

g_exportedScripts = (
    todo_task_add,
    todo_task_delete,
    todo_task_edit,
    todo_task_copy,
    todo_task_colorize,
    todo_task_sort,
    todo_task_move,
    todo_task_merge,
    todo_task_refs,
    todo_task_help,
    todo_task_settings,
)
