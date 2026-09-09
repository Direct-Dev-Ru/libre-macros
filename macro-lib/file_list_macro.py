# -*- coding: utf-8 -*-
MACRO_VERSION = "3.10.713"
import json
import os
import sys
from datetime import datetime
from fnmatch import fnmatch
from urllib.parse import unquote, urlparse

import uno
import unohelper
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.awt import XActionListener
from com.sun.star.awt.MessageBoxResults import CANCEL, NO, YES
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, QUERYBOX

# Константы в pythonpath: AlterOffice 2026 вырезает модульные присваивания в .py-скрипте.
from libre_macros_file_list_cfg import DEFAULT_MAX_DEPTH as _DEFAULT_MAX_DEPTH, DEFAULT_SIZE_UNIT as _DEFAULT_SIZE_UNIT, DEPTH_INDENT as _DEPTH_INDENT, RESULT_HEADER_BASE as _RESULT_HEADER_BASE, SETTINGS_APP_NAME as _SETTINGS_APP_NAME, SETTINGS_MODULE_NAME as _SETTINGS_MODULE_NAME, SIZE_UNIT_DIVISOR as _SIZE_UNIT_DIVISOR, SIZE_UNIT_OPTIONS as _SIZE_UNIT_OPTIONS, UI_YIELD_EVERY as FILE_LIST_UI_YIELD_EVERY, WRITE_CHUNK as FILE_LIST_WRITE_CHUNK
import libre_macros_file_list_cfg as _fl_state


def _file_list_awt_toolkit():
    try:
        return XSCRIPTCONTEXT.getDesktop().getToolkit()
    except Exception:
        pass
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        return ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Toolkit", ctx
        )
    except Exception:
        return None


def file_list_ui_yield(force=False, counter=None):
    """Обработать очередь событий Calc, чтобы интерфейс не зависал."""
    if not force:
        if counter is not None:
            if counter <= 0 or counter % FILE_LIST_UI_YIELD_EVERY != 0:
                return
        else:
            _fl_state.yield_counter = _fl_state.yield_counter + 1
            if _fl_state.yield_counter % FILE_LIST_UI_YIELD_EVERY != 0:
                return
    timeout_ms = 100
    try:
        toolkit = _file_list_awt_toolkit()
        if toolkit is not None:
            if hasattr(toolkit, "processEventsToIdle"):
                toolkit.processEventsToIdle(int(timeout_ms))
            elif hasattr(toolkit, "processEventsToIdleWithMode"):
                toolkit.processEventsToIdleWithMode(int(timeout_ms), 0)
    except Exception:
        pass
    try:
        desktop = XSCRIPTCONTEXT.getDesktop()
        frames = desktop.getFrames()
        if frames is not None and hasattr(frames, "createEnumeration"):
            enum = frames.createEnumeration()
            while enum.hasMoreElements():
                fr = enum.nextElement()
                try:
                    win = fr.getContainerWindow()
                    if win is None:
                        continue
                    tk = win.getToolkit()
                    if tk is not None and hasattr(tk, "processEventsToIdle"):
                        tk.processEventsToIdle(int(timeout_ms))
                        break
                except Exception:
                    pass
    except Exception:
        pass
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        helper = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", ctx
        )
        frame = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
        if helper is not None and frame is not None:
            helper.executeDispatch(frame, ".uno:Repaint", "_self", 0, ())
    except Exception:
        pass


def file_list_ui_status_start(text, max_value):
    file_list_ui_status_finish()
    try:
        frame = XSCRIPTCONTEXT.getDesktop().getCurrentFrame()
        if frame is None:
            return
        si = frame.createStatusIndicator()
        if si is not None:
            si.start(str(text), int(max_value))
            _fl_state.status_indicator = si
    except Exception:
        _fl_state.status_indicator = None


def file_list_ui_status_set(value, text=None):
    try:
        if _fl_state.status_indicator is not None:
            _fl_state.status_indicator.setValue(int(value))
            if text is not None:
                _fl_state.status_indicator.setText(str(text))
    except Exception:
        pass
    file_list_ui_yield(force=True)


def file_list_ui_status_finish():
    if _fl_state.status_indicator is None:
        return
    try:
        _fl_state.status_indicator.end()
    except Exception:
        pass
    _fl_state.status_indicator = None


def _truncate_status_text(text, max_len=72):
    s = str(text)
    if len(s) <= max_len:
        return s
    return s[: max_len - 1] + u"…"


def _uno_to_str(value):
    """Текст из UNO/pyuno в str (кириллица в путях и полях диалога)."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        for enc in ("utf-8", sys.getfilesystemencoding() or "utf-8", "cp1251"):
            if not enc:
                continue
            try:
                return value.decode(enc)
            except Exception:
                pass
        return value.decode("utf-8", errors="replace")
    return str(value)


def _control_get_text(control):
    if control is None:
        return ""
    text = ""
    try:
        text = control.getText()
    except Exception:
        pass
    try:
        model_text = control.Model.Text
        if model_text is not None and str(model_text).strip() != "":
            text = model_text
    except Exception:
        pass
    return _uno_to_str(text).strip()


def _control_set_text(control, text):
    if control is None:
        return
    value = _uno_to_str(text)
    try:
        control.setText(value)
    except Exception:
        pass
    try:
        control.Model.Text = value
    except Exception:
        pass


def _path_to_file_url(path):
    """Системный путь → file:// URL для FolderPicker."""
    sys_path = normalize_path(path)
    if sys_path == "" or not os.path.isdir(sys_path):
        return None
    try:
        return uno.systemPathToFileUrl(sys_path)
    except Exception:
        return None


def path_from_user_input(path_or_url):
    """
    Путь из поля диалога или FolderPicker: file:// URL или системный путь → абсолютный путь ОС.
    """
    text = _uno_to_str(path_or_url).strip()
    if text == "":
        return ""
    low = text.lower()
    if low.startswith("file:"):
        try:
            return os.path.normpath(uno.fileUrlToSystemPath(text))
        except Exception:
            pass
        try:
            parsed = urlparse(text)
            path = unquote(parsed.path or "")
            if (
                sys.platform == "win32"
                and len(path) > 2
                and path[0] == "/"
                and path[2] == ":"
            ):
                path = path[1:]
            if path != "":
                return os.path.normpath(path)
        except Exception:
            pass
    return normalize_path(text)


def _dialog_toolkit():
    return XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstanceWithContext(
        "com.sun.star.awt.Toolkit", XSCRIPTCONTEXT.getComponentContext()
    )


def _parent_window():
    return XSCRIPTCONTEXT.getDesktop().getCurrentFrame().getContainerWindow()


def _settings_dir():
    """Каталог настроек: ~/.config/... (Linux) или %APPDATA%\\... (Windows)."""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, _SETTINGS_APP_NAME, _SETTINGS_MODULE_NAME)


def _settings_file_path():
    return os.path.join(_settings_dir(), "settings.json")


def _default_dialog_settings():
    return {
        "path": "",
        "sheet_name": "",
        "max_depth": _DEFAULT_MAX_DEPTH,
        "masks": "",
        "size_unit": _DEFAULT_SIZE_UNIT,
    }


def normalize_size_unit(value):
    text = str(value or _DEFAULT_SIZE_UNIT).strip().lower()
    aliases = {
        "b": "b",
        "б": "b",
        "байт": "b",
        "bytes": "b",
        "kb": "kb",
        "кб": "kb",
        "k": "kb",
        "mb": "mb",
        "мб": "mb",
        "m": "mb",
        "gb": "gb",
        "гб": "gb",
        "g": "gb",
    }
    return aliases.get(text, _DEFAULT_SIZE_UNIT)


def size_unit_label(unit_key):
    key = normalize_size_unit(unit_key)
    for code, label in _SIZE_UNIT_OPTIONS:
        if code == key:
            return label
    return _SIZE_UNIT_OPTIONS[0][1]


def size_unit_from_label(label):
    text = str(label or "").strip()
    for code, item_label in _SIZE_UNIT_OPTIONS:
        if text == item_label or text.lower() == code:
            return code
    return normalize_size_unit(text)


def build_result_headers(size_unit):
    headers = list(_RESULT_HEADER_BASE)
    headers[5] = "Размер (%s)" % size_unit_label(size_unit)
    return headers


def convert_size_for_display(size_bytes, size_unit):
    unit = normalize_size_unit(size_unit)
    divisor = _SIZE_UNIT_DIVISOR.get(unit, 1)
    value = float(size_bytes) / float(divisor)
    if unit == "b":
        return value
    return round(value, 2)


def load_dialog_settings():
    settings = _default_dialog_settings()
    path = _settings_file_path()
    if not os.path.isfile(path):
        return settings
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, dict):
            for key in settings:
                if key in data:
                    settings[key] = data[key]
    except Exception as err:
        print("Не удалось прочитать настройки: %s" % err)
    try:
        settings["max_depth"] = int(settings.get("max_depth") or _DEFAULT_MAX_DEPTH)
    except (TypeError, ValueError):
        settings["max_depth"] = _DEFAULT_MAX_DEPTH
    if settings["max_depth"] < 1:
        settings["max_depth"] = _DEFAULT_MAX_DEPTH
    settings["size_unit"] = normalize_size_unit(settings.get("size_unit"))
    return settings


def save_dialog_settings(settings):
    try:
        folder = _settings_dir()
        os.makedirs(folder, exist_ok=True)
        payload = _default_dialog_settings()
        if isinstance(settings, dict):
            for key in payload:
                if key in settings:
                    payload[key] = settings[key]
        payload["max_depth"] = int(payload.get("max_depth") or _DEFAULT_MAX_DEPTH)
        payload["size_unit"] = normalize_size_unit(payload.get("size_unit"))
        with open(_settings_file_path(), "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
    except Exception as err:
        print("Не удалось сохранить настройки: %s" % err)


def _set_combo_items(control, items, select_text=None):
    if control is None:
        return
    try:
        seq = tuple(str(x) for x in items)
        control.Model.StringItemList = seq
        if select_text is not None and str(select_text) in seq:
            control.setText(str(select_text))
        elif len(seq) > 0:
            control.setText(seq[0])
        else:
            control.setText("")
    except Exception:
        pass


def create_custom_dialog():
    try:
        toolkit = _dialog_toolkit()

        dialog_model = XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", XSCRIPTCONTEXT.getComponentContext()
        )
        dialog_model.PositionX = 100
        dialog_model.PositionY = 100
        dialog_model.Width = 300
        dialog_model.Height = 200
        dialog_model.Title = "Пользовательский диалог"

        button_yes_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        button_yes_model.PositionX = 50
        button_yes_model.PositionY = 100
        button_yes_model.Width = 80
        button_yes_model.Height = 30
        button_yes_model.Label = "Да"
        button_yes_model.Name = "YesButton"

        button_no_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        button_no_model.PositionX = 140
        button_no_model.PositionY = 100
        button_no_model.Width = 80
        button_no_model.Height = 30
        button_no_model.Label = "Нет"
        button_no_model.Name = "NoButton"

        button_cancel_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        button_cancel_model.PositionX = 230
        button_cancel_model.PositionY = 100
        button_cancel_model.Width = 80
        button_cancel_model.Height = 30
        button_cancel_model.Label = "Отмена"
        button_cancel_model.Name = "CancelButton"

        dialog_model.insertByName("YesButton", button_yes_model)
        dialog_model.insertByName("NoButton", button_no_model)
        dialog_model.insertByName("CancelButton", button_cancel_model)

        dialog = XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", XSCRIPTCONTEXT.getComponentContext()
        )
        dialog.setModel(dialog_model)

        class ButtonHandler(unohelper.Base, XActionListener):
            def __init__(self):
                unohelper.Base.__init__(self)
                self.result = None

            def actionPerformed(self, event):
                self.result = event.Source.Model.Name
                dialog.endExecute()

        handler = ButtonHandler()
        dialog.getControl("YesButton").addActionListener(handler)
        dialog.getControl("NoButton").addActionListener(handler)
        dialog.getControl("CancelButton").addActionListener(handler)

        dialog.createPeer(toolkit, _parent_window())
        dialog.execute()
        return handler.result

    except Exception as e:
        print("Ошибка при создании диалогового окна: %s" % e)
        return None


def show_message(message, title="Сообщение"):
    try:
        toolkit = _dialog_toolkit()
        msgbox = toolkit.createMessageBox(
            _parent_window(), MESSAGEBOX, MSG_BUTTONS.BUTTONS_OK, title, message
        )
        msgbox.execute()
    except Exception as e:
        print("Ошибка при создании диалогового окна: %s" % e)


def show_message_2(message):
    show_message(message, "Отладка")


def show_folder_picker(initial_dir=None):
    try:
        ctx = XSCRIPTCONTEXT.getComponentContext()
        file_picker = ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.ui.dialogs.FolderPicker", ctx
        )
        start = path_from_user_input(initial_dir) if initial_dir else ""
        if start == "" or not os.path.isdir(start):
            start = os.path.expanduser("~")
        start_url = _path_to_file_url(start)
        if start_url is not None:
            file_picker.setDisplayDirectory(start_url)
        if file_picker.execute() == 1:
            return path_from_user_input(file_picker.getDirectory())
        return None
    except Exception as e:
        print("Ошибка при выборе папки: %s" % e)
        return None


def show_sheet_exists_dialog(sheet_name):
    """
    Лист уже есть: YES — перезаписать, NO — удалить и создать заново, CANCEL — отмена.
    """
    try:
        toolkit = _dialog_toolkit()
        text = (
            "Лист «%s» уже существует.\n\n"
            "Да — перезаписать содержимое\n"
            "Нет — удалить лист и создать заново\n"
            "Отмена — прервать операцию"
        ) % sheet_name
        msgbox = toolkit.createMessageBox(
            _parent_window(),
            QUERYBOX,
            MSG_BUTTONS.BUTTONS_YES_NO_CANCEL,
            "Лист существует",
            text,
        )
        answer = msgbox.execute()
        if answer == YES:
            return "overwrite"
        if answer == NO:
            return "recreate"
        return "cancel"
    except Exception as e:
        print("Ошибка диалога листа: %s" % e)
        return "cancel"


def normalize_path(path):
    """Абсолютный путь с учётом ОС (Windows / Linux)."""
    expanded = os.path.expanduser(os.path.expandvars(_uno_to_str(path).strip()))
    if expanded == "":
        return ""
    return os.path.normpath(os.path.abspath(expanded))


def display_path(path):
    """Относительный путь для вывода в таблицу (единый вид разделителя)."""
    if path in (".", ""):
        return "."
    return path.replace("\\", "/")


def parse_masks(mask_text):
    """
    Маски файлов: *.xlsx;*.pdf или через запятую. Пусто — все файлы.
    """
    text = str(mask_text or "").strip()
    if text == "":
        return []
    parts = []
    for chunk in text.replace(",", ";").split(";"):
        item = chunk.strip()
        if item != "" and item not in parts:
            parts.append(item)
    return parts


def file_matches_masks(file_name, masks):
    if not masks:
        return True
    for pattern in masks:
        if fnmatch(file_name, pattern):
            return True
    return False


def parent_relative(root_path, full_path):
    """Родитель относительно корневого пути."""
    rel = os.path.relpath(os.path.dirname(full_path), root_path)
    return display_path(rel)


def relative_entry_path(root_path, full_path):
    return display_path(os.path.relpath(full_path, root_path))


def format_depth_indent(text, depth):
    """Визуальный отступ в ячейке: глубина 1 — без отступа, далее +1 уровень."""
    level = max(0, int(depth) - 1)
    return (_DEPTH_INDENT * level) + str(text)


def sort_entries_hierarchical(entries):
    """
    Древовидный порядок: предки перед потомками по сегментам пути.
    На одном уровне — папки перед файлами с тем же путём-префиксом.
    """
    def sort_key(item):
        rel = str(item.get("rel_path", "")).replace("\\", "/")
        parts = tuple(p.lower() for p in rel.split("/") if p != "")
        kind_rank = 0 if item.get("kind") == "папка" else 1
        return (parts, kind_rank)

    sorted_entries = list(entries)
    sorted_entries.sort(key=sort_key)
    return sorted_entries


def _set_cell_number(cell, value):
    """Calc ожидает double для Value; Python int может передаться как hyper."""
    try:
        cell.Value = float(value)
    except Exception:
        cell.String = str(value)


def dir_tree_size(dir_path, folder_depth, max_depth, masks, op_counter=None):
    """
    Размер папки: сумма файлов внутри с учётом лимита глубины и масок.
    """
    total = 0
    try:
        names = os.listdir(dir_path)
    except OSError:
        return 0
    for name in names:
        if op_counter is not None:
            op_counter[0] = op_counter[0] + 1
            file_list_ui_yield(counter=op_counter[0])
        full = os.path.join(dir_path, name)
        if os.path.isfile(full):
            if file_matches_masks(name, masks):
                try:
                    total += os.path.getsize(full)
                except OSError:
                    pass
        elif os.path.isdir(full) and folder_depth < max_depth:
            total += dir_tree_size(
                full, folder_depth + 1, max_depth, masks, op_counter
            )
    return total


def collect_path_entries(root_path, max_depth=None, masks=None, show_progress=True):
    """
  Собрать файлы и папки до указанной глубины.

  max_depth=1 — только первый уровень внутри root_path (по умолчанию).
  """
    root = path_from_user_input(root_path)
    if root == "" or not os.path.isdir(root):
        return [], root

    if max_depth is None:
        max_depth = _DEFAULT_MAX_DEPTH
    try:
        max_depth = int(max_depth)
    except (TypeError, ValueError):
        max_depth = _DEFAULT_MAX_DEPTH
    if max_depth < 1:
        max_depth = 1

    mask_list = masks if masks is not None else []
    entries = []
    op_counter = [0]
    top_progress = [0]
    _fl_state.yield_counter = 0

    try:
        top_names = sorted(os.listdir(root))
    except OSError as err:
        print("Не удалось прочитать корень «%s»: %s" % (root, err))
        return [], root
    top_total = max(1, len(top_names))

    if show_progress:
        file_list_ui_status_start(u"Анализ каталога", top_total)

    def walk_dir(dir_path, depth_of_dir):
        try:
            names = sorted(os.listdir(dir_path))
        except OSError as err:
            print("Не удалось прочитать «%s»: %s" % (dir_path, err))
            return

        for name in names:
            full = os.path.join(dir_path, name)
            item_depth = depth_of_dir + 1
            if item_depth > max_depth:
                continue

            if depth_of_dir == 0:
                top_progress[0] = top_progress[0] + 1
                if show_progress:
                    if os.path.isdir(full):
                        status = u"1-й уровень (папка): %s" % _truncate_status_text(name)
                    else:
                        status = u"1-й уровень (файл): %s" % _truncate_status_text(name)
                    file_list_ui_status_set(top_progress[0], status)

            op_counter[0] = op_counter[0] + 1
            if show_progress:
                file_list_ui_yield(counter=op_counter[0])

            try:
                mtime = os.path.getmtime(full)
                mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
            except OSError:
                mtime_str = ""

            parent = parent_relative(root, full)
            rel_path = relative_entry_path(root, full)

            if os.path.isfile(full):
                if not file_matches_masks(name, mask_list):
                    continue
                try:
                    size = os.path.getsize(full)
                except OSError:
                    size = 0
                entries.append(
                    {
                        "kind": "файл",
                        "name": name,
                        "rel_path": rel_path,
                        "parent": parent,
                        "depth": item_depth,
                        "size": size,
                        "mtime": mtime_str,
                        "ext": os.path.splitext(name)[1],
                        "full_path": os.path.normpath(full),
                    }
                )
            elif os.path.isdir(full):
                size_counter = op_counter if show_progress else None
                size = dir_tree_size(full, item_depth, max_depth, mask_list, size_counter)
                entries.append(
                    {
                        "kind": "папка",
                        "name": name,
                        "rel_path": rel_path,
                        "parent": parent,
                        "depth": item_depth,
                        "size": size,
                        "mtime": mtime_str,
                        "ext": "",
                        "full_path": os.path.normpath(full),
                    }
                )
                walk_dir(full, item_depth)

    try:
        walk_dir(root, 0)
    finally:
        if show_progress:
            file_list_ui_status_finish()

    return sort_entries_hierarchical(entries), root


def get_sheet_by_name(doc, sheet_name):
    if doc is None or sheet_name is None:
        return None
    name = str(sheet_name).strip()
    if name == "":
        return None
    try:
        return doc.Sheets.getByName(name)
    except Exception:
        pass
    try:
        for sheet in doc.Sheets:
            if sheet.Name == name:
                return sheet
    except Exception:
        pass
    return None


def _sheet_names(doc):
    try:
        return [sheet.Name for sheet in doc.Sheets]
    except Exception:
        return []


def _insert_sheet(doc, sheet_name):
    """Создать лист и вернуть его по имени (надёжнее, чем объект из insertNewByName)."""
    name = str(sheet_name).strip()
    if name == "":
        return None
    try:
        index = int(doc.Sheets.getCount())
    except Exception:
        index = 0
    try:
        doc.Sheets.insertNewByName(name, index)
    except Exception as err:
        print("Не удалось создать лист «%s»: %s" % (name, err))
        return None
    return get_sheet_by_name(doc, name)


def _ensure_active_sheet(doc, sheet):
    if doc is None or sheet is None:
        return False
    try:
        controller = doc.getCurrentController()
        if controller is None:
            return False
        active = controller.getActiveSheet()
        if active is not None and active.Name == sheet.Name:
            return True
        controller.setActiveSheet(sheet)
        return True
    except Exception as err:
        print("Не удалось активировать лист: %s" % err)
        try:
            by_name = get_sheet_by_name(doc, sheet.Name)
            if by_name is not None:
                doc.getCurrentController().setActiveSheet(by_name)
                return True
        except Exception:
            pass
    return False


def _focus_sheet_top_left(doc, sheet):
    if not _ensure_active_sheet(doc, sheet):
        return
    try:
        controller = doc.getCurrentController()
        if controller is not None:
            controller.select(sheet.getCellRangeByPosition(0, 0, 0, 0))
    except Exception:
        pass


def clear_sheet_contents(sheet, max_cols=16, max_rows=20000):
    if sheet is None:
        return
    end_row = 0
    end_col = max_cols - 1
    try:
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        used = cursor.getRangeAddress()
        end_row = max(int(used.EndRow), 0)
        end_col = max(int(used.EndColumn), max_cols - 1)
    except Exception:
        end_row = 0
    try:
        range_addr = sheet.getCellRangeByPosition(0, 0, end_col, end_row)
        range_addr.clearContents(7)
    except Exception:
        for row in range(min(end_row + 1, 500)):
            for col in range(max_cols):
                cell = sheet.getCellByPosition(col, row)
                cell.String = ""
                try:
                    cell.Value = 0.0
                except Exception:
                    pass


def resolve_target_sheet(doc, sheet_name):
    """
    Получить лист для вывода. None/пусто — активный лист.
    Если лист есть — спросить: перезаписать / удалить и создать / отмена.
    """
    if sheet_name is None or str(sheet_name).strip() == "":
        return doc.getCurrentController().getActiveSheet(), "active"

    target_name = str(sheet_name).strip()
    if target_name not in _sheet_names(doc):
        new_sheet = _insert_sheet(doc, target_name)
        if new_sheet is None:
            show_message("Не удалось создать лист «%s»." % target_name)
            return None, "error"
        return new_sheet, "created"

    choice = show_sheet_exists_dialog(target_name)
    if choice == "cancel":
        return None, "cancel"
    if choice == "recreate":
        try:
            doc.Sheets.removeByName(target_name)
        except Exception as err:
            show_message("Не удалось удалить лист «%s»: %s" % (target_name, err))
            return None, "error"
        new_sheet = _insert_sheet(doc, target_name)
        if new_sheet is None:
            return None, "error"
        return new_sheet, "recreated"

    sheet = get_sheet_by_name(doc, target_name)
    if sheet is None:
        show_message("Лист «%s» не найден." % target_name)
        return None, "error"
    clear_sheet_contents(sheet)
    return sheet, "overwrite"


def _build_sheet_data_rows(entries, size_unit):
    """Строки данных для setDataArray (без заголовка)."""
    unit = normalize_size_unit(size_unit)
    rows = []
    for item in entries:
        depth = int(item["depth"])
        rows.append(
            [
                _uno_to_str(item["kind"]),
                format_depth_indent(item["name"], depth),
                format_depth_indent(item["rel_path"], depth),
                _uno_to_str(item["parent"]),
                float(depth),
                float(convert_size_for_display(item["size"], unit)),
                _uno_to_str(item["mtime"]),
                _uno_to_str(item["ext"]),
                _uno_to_str(item.get("full_path", "")),
            ]
        )
    return rows


def _sheet_write_rows_fallback(sheet, start_row, rows):
    """Поштучная запись, если setDataArray недоступен (моки, старый LO)."""
    for ri, row in enumerate(rows):
        r = start_row + ri
        for col, value in enumerate(row):
            cell = sheet.getCellByPosition(col, r)
            if col in (4, 5) and isinstance(value, (int, float)):
                _set_cell_number(cell, value)
            else:
                cell.String = _uno_to_str(value)


def _sheet_set_data_array(sheet, start_row, rows):
    """
    Записать блок строк одним setDataArray (как getDataArray в collect_workbooks).
    Возвращает True при успехе.
    """
    if not rows:
        return True
    nrows = len(rows)
    ncols = len(rows[0])
    try:
        data = tuple(tuple(row) for row in rows)
        end_row = start_row + nrows - 1
        rng = sheet.getCellRangeByPosition(0, start_row, ncols - 1, end_row)
        rng.setDataArray(data)
        return True
    except Exception as err:
        print("setDataArray: %s — поштучная запись" % err)
        _sheet_write_rows_fallback(sheet, start_row, rows)
        return False


def _format_header_row(sheet, ncols):
    try:
        sheet.getCellRangeByPosition(0, 0, ncols - 1, 0).CharWeight = 150
    except Exception:
        col = 0
        while col < ncols:
            try:
                sheet.getCellByPosition(col, 0).CharWeight = 150
            except Exception:
                pass
            col += 1


def write_entries_to_sheet(results_sheet, entries, size_unit=_DEFAULT_SIZE_UNIT, clear_first=True, show_progress=False):
    if clear_first:
        clear_sheet_contents(results_sheet)
    headers = build_result_headers(size_unit)
    ncols = len(headers)

    _sheet_set_data_array(results_sheet, 0, [list(headers)])
    _format_header_row(results_sheet, ncols)

    data_rows = _build_sheet_data_rows(entries, size_unit)
    total = len(data_rows)
    chunk_count = max(1, (total + FILE_LIST_WRITE_CHUNK - 1) // FILE_LIST_WRITE_CHUNK)

    if show_progress and total > 0:
        file_list_ui_status_start(u"Запись на лист", chunk_count)

    written = 0
    row_offset = 1
    chunk_idx = 0
    try:
        i = 0
        while i < total:
            chunk = data_rows[i : i + FILE_LIST_WRITE_CHUNK]
            _sheet_set_data_array(results_sheet, row_offset, chunk)
            row_offset += len(chunk)
            written += len(chunk)
            i += len(chunk)
            chunk_idx += 1
            if show_progress:
                file_list_ui_status_set(
                    chunk_idx,
                    u"Записано %d / %d строк" % (written, total),
                )
            file_list_ui_yield(force=True)
    finally:
        if show_progress:
            file_list_ui_status_finish()

    try:
        for col in range(ncols):
            results_sheet.getColumns().getByIndex(col).OptimalWidth = True
    except Exception:
        pass

    return written


def freeze_sheet_first_row(doc, sheet):
    """Закрепить первую строку листа (заголовки таблицы)."""
    if doc is None or sheet is None:
        return
    try:
        _ensure_active_sheet(doc, sheet)
        controller = doc.getCurrentController()
        if controller is None:
            return
        controller.freezeAtPosition(0, 1)
    except Exception as err:
        print("Не удалось закрепить первую строку: %s" % err)


def analyze_files(path, sheet_name=None, max_depth=None, masks=None, size_unit=_DEFAULT_SIZE_UNIT):
    """
    Собирает данные о файлах и папках и выводит на лист.

    max_depth — глубина обхода (1 = первый уровень внутри path).
    masks — строка или список масок (*.xlsx;*.pdf).
    size_unit — b | kb | mb | gb (б, Кб, Мб, Гб).
    """
    try:
        root = path_from_user_input(path)
        if root == "" or not os.path.isdir(root):
            show_message(
                "Указанный путь «%s» не существует или не является папкой." % path
            )
            return False

        if isinstance(masks, str):
            mask_list = parse_masks(masks)
        elif masks:
            mask_list = list(masks)
        else:
            mask_list = []

        entries, root = collect_path_entries(
            root, max_depth=max_depth, masks=mask_list, show_progress=True
        )
        if len(entries) == 0:
            if mask_list:
                show_message(
                    "В «%s» (глубина %s) нет файлов по маскам: %s"
                    % (root, max_depth or _DEFAULT_MAX_DEPTH, ", ".join(mask_list))
                )
            else:
                show_message(
                    "В «%s» (глубина %s) не найдено файлов и папок."
                    % (root, max_depth or _DEFAULT_MAX_DEPTH)
                )
            return False

        doc = XSCRIPTCONTEXT.getDocument()
        results_sheet, mode = resolve_target_sheet(doc, sheet_name)
        if results_sheet is None:
            return False

        _ensure_active_sheet(doc, results_sheet)
        written = write_entries_to_sheet(
            results_sheet,
            entries,
            size_unit=normalize_size_unit(size_unit),
            clear_first=(mode != "created"),
            show_progress=True,
        )
        freeze_sheet_first_row(doc, results_sheet)
        _focus_sheet_top_left(doc, results_sheet)
        print(
            "Записано %s элементов на лист «%s» (режим: %s)"
            % (written, results_sheet.Name, mode)
        )
        return True

    except Exception as e:
        error_msg = "Ошибка при анализе файлов: %s" % e
        show_message(error_msg)
        print(error_msg)
        import traceback

        print(traceback.format_exc())
        return False


def _add_label(dialog_model, name, text, x, y, w, h):
    model = dialog_model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    model.PositionX = x
    model.PositionY = y
    model.Width = w
    model.Height = h
    model.Label = text
    model.Name = name
    dialog_model.insertByName(name, model)
    return model


def _add_edit(dialog_model, name, x, y, w, h, text=""):
    model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
    model.PositionX = x
    model.PositionY = y
    model.Width = w
    model.Height = h
    model.Name = name
    model.Text = text
    dialog_model.insertByName(name, model)
    return model


def _add_button(dialog_model, name, label, x, y, w, h):
    model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    model.PositionX = x
    model.PositionY = y
    model.Width = w
    model.Height = h
    model.Label = label
    model.Name = name
    dialog_model.insertByName(name, model)
    return model


def _add_combo(dialog_model, name, x, y, w, h):
    model = dialog_model.createInstance("com.sun.star.awt.UnoControlComboBoxModel")
    model.PositionX = x
    model.PositionY = y
    model.Width = w
    model.Height = h
    model.Name = name
    model.Dropdown = True
    dialog_model.insertByName(name, model)
    return model


def create_file_analyzer_dialog():
    try:
        saved = load_dialog_settings()
        toolkit = _dialog_toolkit()
        dialog_model = XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", XSCRIPTCONTEXT.getComponentContext()
        )

        dlg_w = 520
        dlg_h = 350
        m = 10
        field_w = dlg_w - m * 2 - 100 - 8
        dialog_model.PositionX = 100
        dialog_model.PositionY = 100
        dialog_model.Width = dlg_w
        dialog_model.Height = dlg_h
        dialog_model.Title = "Сбор данных о файлах и папках"

        y = m
        _add_label(dialog_model, "LabelPath", "Путь к папке:", m, y, 140, 14)
        y += 18
        _add_edit(
            dialog_model,
            "PathField",
            m,
            y,
            field_w,
            22,
            path_from_user_input(saved.get("path", "")),
        )
        _add_button(dialog_model, "BrowseButton", "Обзор...", m + field_w + 8, y - 2, 100, 26)

        y += 34
        _add_label(
            dialog_model,
            "LabelSheet",
            "Имя листа (пусто = активный; если есть — спросим действие):",
            m,
            y,
            dlg_w - m * 2,
            14,
        )
        y += 18
        _add_edit(
            dialog_model,
            "SheetField",
            m,
            y,
            dlg_w - m * 2,
            22,
            str(saved.get("sheet_name", "")),
        )

        y += 34
        _add_label(
            dialog_model,
            "LabelDepth",
            "Глубина (1 = первый уровень):",
            m,
            y,
            180,
            14,
        )
        _add_label(
            dialog_model,
            "LabelSizeUnit",
            "Единицы размера:",
            m + 250,
            y,
            140,
            14,
        )
        y += 18
        _add_edit(
            dialog_model,
            "DepthField",
            m,
            y,
            80,
            22,
            str(saved.get("max_depth", _DEFAULT_MAX_DEPTH)),
        )
        _add_combo(dialog_model, "SizeUnitCombo", m + 250, y, 120, 22)

        y += 34
        _add_label(
            dialog_model,
            "LabelMasks",
            "Маски файлов (пусто = все; пример: *.xlsx;*.pdf):",
            m,
            y,
            dlg_w - m * 2,
            14,
        )
        y += 18
        _add_edit(
            dialog_model,
            "MaskField",
            m,
            y,
            dlg_w - m * 2,
            22,
            str(saved.get("masks", "")),
        )

        btn_y = dlg_h - 42
        _add_button(dialog_model, "RunButton", "Запуск", m, btn_y, 100, 28)
        _add_button(dialog_model, "ClearButton", "Очистить", m + 110, btn_y, 100, 28)
        _add_button(dialog_model, "CancelButton", "Отмена", m + 220, btn_y, 100, 28)

        dialog = XSCRIPTCONTEXT.getComponentContext().getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", XSCRIPTCONTEXT.getComponentContext()
        )
        dialog.setModel(dialog_model)

        class DialogHandler(unohelper.Base, XActionListener):
            def __init__(self, dialog_ref):
                unohelper.Base.__init__(self)
                self.dialog = dialog_ref
                self.result = None
                self.path_field = None
                self.sheet_field = None
                self.depth_field = None
                self.mask_field = None
                self.size_unit_combo = None

            def actionPerformed(self, event):
                button_name = event.Source.Model.Name

                if button_name == "BrowseButton":
                    current = _control_get_text(self.path_field)
                    selected_path = show_folder_picker(current if current else None)
                    if selected_path:
                        _control_set_text(self.path_field, selected_path)

                elif button_name == "ClearButton":
                    _control_set_text(self.path_field, "")
                    _control_set_text(self.sheet_field, "")
                    _control_set_text(self.depth_field, str(_DEFAULT_MAX_DEPTH))
                    _control_set_text(self.mask_field, "")
                    if self.size_unit_combo:
                        _set_combo_items(
                            self.size_unit_combo,
                            [label for _code, label in _SIZE_UNIT_OPTIONS],
                            _SIZE_UNIT_OPTIONS[0][1],
                        )

                elif button_name == "RunButton":
                    try:
                        path = _control_get_text(self.path_field)
                        sheet_name = _control_get_text(self.sheet_field)
                        depth_text = _control_get_text(self.depth_field)
                        mask_text = _control_get_text(self.mask_field)
                        size_unit_text = (
                            _control_get_text(self.size_unit_combo)
                            if self.size_unit_combo
                            else size_unit_label(_DEFAULT_SIZE_UNIT)
                        )

                        path = path_from_user_input(path)
                        sheet_name = sheet_name.strip() if sheet_name else ""
                        depth_text = depth_text.strip() if depth_text else str(_DEFAULT_MAX_DEPTH)
                        mask_text = mask_text.strip() if mask_text else ""
                        size_unit = size_unit_from_label(size_unit_text)

                        if not path:
                            show_message("Укажите путь к папке.")
                            return

                        try:
                            max_depth = int(depth_text)
                            if max_depth < 1:
                                show_message("Глубина должна быть не меньше 1.")
                                return
                        except ValueError:
                            show_message("Глубина должна быть целым числом.")
                            return

                        save_dialog_settings(
                            {
                                "path": path,
                                "sheet_name": sheet_name,
                                "max_depth": max_depth,
                                "masks": mask_text,
                                "size_unit": size_unit,
                            }
                        )
                        self.result = (
                            "run",
                            path,
                            sheet_name,
                            max_depth,
                            mask_text,
                            size_unit,
                        )
                        self.dialog.endExecute()
                    except Exception as e:
                        show_message("Ошибка чтения полей диалога: %s" % e)

                elif button_name == "CancelButton":
                    self.result = ("cancel", None, None, None, None, None)
                    self.dialog.endExecute()

        handler = DialogHandler(dialog)
        handler.path_field = dialog.getControl("PathField")
        handler.sheet_field = dialog.getControl("SheetField")
        handler.depth_field = dialog.getControl("DepthField")
        handler.mask_field = dialog.getControl("MaskField")
        handler.size_unit_combo = dialog.getControl("SizeUnitCombo")
        _set_combo_items(
            handler.size_unit_combo,
            [label for _code, label in _SIZE_UNIT_OPTIONS],
            size_unit_label(saved.get("size_unit", _DEFAULT_SIZE_UNIT)),
        )

        for btn in ("BrowseButton", "ClearButton", "RunButton", "CancelButton"):
            dialog.getControl(btn).addActionListener(handler)

        dialog.createPeer(toolkit, _parent_window())
        dialog.execute()
        return handler.result

    except Exception as e:
        print("Ошибка при создании диалогового окна: %s" % e)
        show_message("Ошибка при создании диалогового окна: %s" % e)
        return None


def analyze_files_dialog(*args):
    try:
        result = create_file_analyzer_dialog()

        if result and result[0] == "run":
            _, path, sheet_name, max_depth, mask_text, size_unit = result
            sheet_arg = sheet_name if sheet_name else None
            print(
                "Запуск: путь=%s, лист=%s, глубина=%s, маски=%s, единицы=%s"
                % (path, sheet_arg, max_depth, mask_text, size_unit)
            )
            success = analyze_files(
                path,
                sheet_name=sheet_arg,
                max_depth=max_depth,
                masks=mask_text,
                size_unit=size_unit,
            )
            if success:
                mask_note = ""
                if mask_text:
                    mask_note = "\nМаски: %s" % mask_text
                show_message(
                    "Сбор завершён.\nПуть: %s\nГлубина: %s\nРазмер: %s%s"
                    % (
                        normalize_path(path),
                        max_depth,
                        size_unit_label(size_unit),
                        mask_note,
                    )
                )
            else:
                show_message(
                    "Сбор не выполнен.\nПроверьте путь, маски и имя листа: %s"
                    % (sheet_arg or "(активный)")
                )
        elif result and result[0] == "cancel":
            print("Операция отменена пользователем")

    except Exception as e:
        error_msg = "Ошибка при выполнении макроса: %s" % e
        show_message(error_msg)
        print(error_msg)
        import traceback

        print(traceback.format_exc())


def get_file_list(*args):
    result = create_custom_dialog()

    if result == "YesButton":
        print("Нажата кнопка 'Да'")
    elif result == "NoButton":
        print("Нажата кнопка 'Нет'")
    elif result == "CancelButton":
        print("Нажата кнопка 'Отмена'")
    else:
        print("Диалог закрыт без выбора")

    doc = XSCRIPTCONTEXT.getDocument()
    show_message_2("Python macros starting")

    if "Параметры" not in [sheet.Name for sheet in doc.Sheets]:
        print("Лист 'Параметры' не найден!")
        return

    params_sheet = doc.Sheets["Параметры"]
    path_cell = params_sheet.getCellByPosition(0, 1)
    path = path_cell.String.strip()

    if not os.path.isdir(path):
        print("Указанный путь '%s' не существует или не является директорией." % path)
        return

    file_data = []
    for file_name in os.listdir(path):
        full_path = os.path.join(path, file_name)
        if os.path.isfile(full_path):
            file_size = os.path.getsize(full_path)
            file_data.append((file_name, file_size))

    if "Результаты" in [sheet.Name for sheet in doc.Sheets]:
        results_sheet = doc.Sheets["Результаты"]
        results_sheet.Rows.removeByIndex(0, results_sheet.Rows.Count)
    else:
        results_sheet = doc.Sheets.insertNewByName("Результаты", len(doc.Sheets))

    headers = ["Имя файла", "Размер (байт)"]
    for col, header in enumerate(headers):
        results_sheet.getCellByPosition(col, 0).String = header

    for row, (file_name, file_size) in enumerate(file_data, start=1):
        results_sheet.getCellByPosition(0, row).String = file_name
        results_sheet.getCellByPosition(1, row).Value = file_size

    print("Обработано %s файлов." % len(file_data))


g_exportedScripts = (get_file_list, analyze_files_dialog)
