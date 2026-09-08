MACRO_VERSION = "3.10.710"
# region Справка — namespace и хелперы (свернуть: ▼ слева или Ctrl+Shift+[)
"""
functions_final.py — пользовательские функции финальной обработки (вся книга).

Формат на листе «Параметры_Объединения*»
----------------------------------------
  A = Финальная_обработка
  B = функция_плагин
  C = [{"v":1,"fn":"функция_плагин","ref":"functions_final.py#pp_range_zagotovka","extra":"42","sheet":"Сводка"}]
  D = (пусто, если extra в JSON)

Допуск: allow_env / allow_global — имена OS env и зашифрованной глобальной;
пусто в визарде/JSON → MERGE_ALLOW_PLUGINS / Merge_Allow_Plugins.
Значения env и глобальной должны совпасть, иначе сбор прерывается.

Выполняется один раз по всей книге перед диалогом «Нажмите Завершить».
Поле sheet в JSON ограничивает обработку указанными листами результата.

Сигнатура def pp_*
------------------
  (doc, sheets, data_ranges, header_row_ranges, *extra_args)

  sheets[i], data_ranges[i], header_row_ranges[i] — один лист и его диапазоны.

Импорты в этом файле для хелперов макроса не нужны: при загрузке def через
«функция_плагин» код компилируется с namespace _merge_final_eval_namespace()
(collect_workbooks.py).

Хелперы и утилиты
-----------------
  cell_text(cell)              Текст ячейки (строка).
  lo_color_rgb(red, green, blue)   Цвет LibreOffice (int).

  merge_pp_active_context()    dict контекста (если нужен в финальном шаге).
      Для следующих финальных шагов диапазоны пересобираются автоматически
      движком перед каждым вызовом. Если внутри ТЕКУЩЕЙ функции нужен новый
      диапазон после вставки строк/столбцов — перечитайте его с листа заново.

  merge_pp_user_log(doc, sheet, message, status)
      Запись в лист «Сбор_книг_лог» (sheet — любой лист для привязки записи).

  merge_ui_yield(counter)      Отдать управление UI.
  merge_ui_status_set(n, text) Обновить StatusIndicator.

  traceback                    Модуль traceback (для отладки в except).

В финальном namespace также доступны все merge_pp_* (постобработка) — можно
вызывать из финальной функции при необходимости, но сигнатуры у них
диапазонные/строчные, не «по всей книге».

Встроенные колбэки merge_final_* (из libre_macros_final_lib и collect_workbooks)
--------------------------------------------------------------------------------
  merge_final_delete_sheets, merge_final_hide_sheets, merge_final_values_only,
  merge_final_header_plus_height, merge_final_vert_center, merge_final_zebra,
  merge_final_sort, merge_final_colorize, merge_final_grid, merge_final_thin_grid,
  merge_final_thick_grid, merge_final_word_wrap, merge_final_word_wrap_and_fit,
  merge_final_autofit_rows, merge_final_autofit_columns, merge_final_left_align,
  merge_final_reset_path_stripes, merge_final_freeze_header, merge_final_autofilter,
  merge_final_row_height, merge_final_indent, merge_final_set_font,
  merge_final_column_width, merge_final_format_money, merge_final_format_date,
  merge_final_format_columns, merge_final_highlight_threshold,
  merge_final_color_scale, merge_final_delete_columns, merge_final_concat_columns,
  merge_final_apply_formula, merge_final_delete_rows, merge_final_group_by_column,
  merge_final_print_style, merge_final_fill_down, merge_final_fill_down_calculate,
  merge_final_copy_values, merge_final_rename_sheet, …

Ключи карт — вызываются как функции в lambda/def
------------------------------------------------
Финальная обработка (MERGE_FINAL_PROCESSING_MAP):
  удаление_листов, скрытие_листов, только_значения, заголовок_плюс_высота,
  вертикаль_центр, зебра_диапазон, сортировка, раскрасить_блоки, сетка,
  тонкая_сетка, толстая_сетка, перенос, перенос_и_авто_высота, авто_высота,
  авто_ширина, левое_выравнивание, сброс_стрипов, закрепить_заголовок,
  автофильтр, высота_строки, отступ, шрифт, ширина_столбцов, формат_деньги,
  формат_даты, формат_столбцы, подсветка_по_порогу, условное_форматирование,
  градиент, удалить_столбцы, конкатенация_столбцов, применить_формулу,
  удаление_строк, пропуск_пустых_строк, группировка_по_столбцу, стиль_печати,
  заполнение_вниз, заполнение_вниз_вычислить, копировать_значения, переименовать_лист,
  переименовать_столбцы, переставить_столбцы, количество_значений, копировать_лист,
  удалить_дубликаты.

Пример: lambda doc,s,dr,hr: только_значения(doc, s, dr, hr, "Сводная,1_Заказы")

Ограничения builtins
--------------------
Те же, что у постобработки: без __import__, eval, exec, open, globals, …
(см. _MERGE_PP_EVAL_BUILTIN_BLOCKLIST в collect_workbooks.py).

Подробнее: macro-lib/USER_PP_FUNCTIONS_GUIDE.md, macro-lib/POSTPROCESS_GUIDE.md
"""
# endregion Справка — namespace и хелперы

import traceback


# region LLM — промпт для генерации финальной функции (свернуть)
LLM_PROMPT_FINAL_RULES = u"""Ты пишешь пользовательскую функцию ФИНАЛЬНОЙ обработки для макроса LibreOffice Calc «Сбор книг».

Контекст
--------
Файл: functions_final.py.
Загрузка: A=Финальная_обработка, B=функция_плагин, C=JSON (ref/code, extra, sheet).
Namespace: _merge_final_eval_namespace() — импорты в файле не нужны.

Сигнатура (обязательно)
-----------------------
  def pp_имя(doc, sheets, data_ranges, header_row_ranges, *extra_args):

  sheets[i], data_ranges[i], header_row_ranges[i] — один лист и его диапазоны.
  При sheet в JSON макрос передаёт только отфильтрованные листы.

Доступные имена
---------------
  cell_text, lo_color_rgb, merge_pp_active_context, merge_pp_user_log(..., scope="финальная_обработка")
  merge_ui_yield, merge_ui_status_set, traceback
  merge_final_* и merge_pp_* (сигнатуры диапазонные — вызывать осознанно)

Запрещено
---------
  Те же ограничения eval, что у постобработки (_MERGE_PP_EVAL_BUILTIN_BLOCKLIST).
  Не закрывать doc; удаление листов — только через merge_final_delete_sheets / ключ удаление_листов.

Стиль
-----
- Python 2.7 / LO: без f-strings; while-циклы по индексам.
- Лог: merge_pp_user_log(doc, sheet, msg, status, scope="финальная_обработка").
- Ошибки: try/except + лог + raise.
- Если функция меняет размер листа, следующие финальные шаги уже увидят новый
  диапазон автоматически; для повторного использования в этой же функции —
  заново вычисли data_range/header_row_range по sheet.

Что должна выполнять функция:
----------
- 
- 
-

Вывод: только def pp_... без markdown.
"""


# endregion LLM — промпт для генерации финальной функции


def pp_range_zagotovka(doc, sheets, data_ranges, header_row_ranges, *extra_args):
    """
    Заготовка финальной обработки по всей книге.

    Сигнатура: (doc, sheets, data_ranges, header_row_ranges, *extra_args).
    sheets[i], data_ranges[i], header_row_ranges[i] — один лист и его диапазоны.
    Параметры extra — JSON поле extra или колонка D.

    Пример JSON в C:
        [{"v":1,"fn":"функция_плагин","ref":"functions_final.py#pp_range_zagotovka","extra":"log"}]

    Если функция добавляет строки/столбцы и дальше в ЭТОЙ ЖЕ функции нужен
    новый диапазон, перечитайте его с листа, например:
        end_col = 0
        end_row = 0
        try:
            cursor = sheet.createCursor()
            cursor.gotoEndOfUsedArea(True)
            addr = cursor.getRangeAddress()
            end_col = addr.EndColumn
            end_row = addr.EndRow
        except Exception:
            pass
        header_row_range = sheet.getCellRangeByPosition(0, 0, end_col, 0)
        data_range = sheet.getCellRangeByPosition(0, 1, end_col, end_row)

    """
    param1 = extra_args[0] if len(extra_args) > 0 else ""
    param2 = extra_args[1] if len(extra_args) > 1 else ""

    try:
        n = len(sheets) if sheets is not None else 0
        i = 0
        while i < n:
            sheet = sheets[i]
            data_range = data_ranges[i] if i < len(data_ranges) else None
            header_row_range = (
                header_row_ranges[i] if i < len(header_row_ranges) else None
            )
            sheet_name = sheet.Name if sheet is not None else "?"
            # --- Твоя логика на один лист ---
            # Пример: param1 == "log" — записать в журнал имя листа
            if param1 == "log":
                merge_pp_user_log(
                    doc,
                    sheet,
                    "лист %d/%d: %s" % (i + 1, n, sheet_name),
                    "инфо",
                    scope="финальная_обработка",
                )
            # Если выше меняли структуру листа и хотим продолжить уже по новому
            # диапазону внутри этой же функции, перечитываем границы заново.
            # try:
            #     cursor = sheet.createCursor()
            #     cursor.gotoEndOfUsedArea(True)
            #     addr = cursor.getRangeAddress()
            #     end_col = addr.EndColumn
            #     end_row = addr.EndRow
            #     header_row_range = sheet.getCellRangeByPosition(0, 0, end_col, 0)
            #     data_range = sheet.getCellRangeByPosition(0, 1, end_col, end_row)
            # except Exception:
            #     pass
            i = i + 1
        params_info = ", extra_args: %s" % str(extra_args) if extra_args else ""
        merge_pp_user_log(
            doc,
            sheets[0] if n > 0 else None,
            "финальная обработка: %d лист(ов)%s" % (n, params_info),
            "ok",
            scope="финальная_обработка",
        )
    except Exception as err:
        tb = ""
        try:
            tb = traceback.format_exc()
        except Exception:
            pass
        merge_pp_user_log(
            doc,
            sheets[0] if sheets is not None and len(sheets) > 0 else None,
            "ОШИБКА: %s\n%s" % (err, tb),
            "ошибка",
            scope="финальная_обработка",
        )
        raise


# ---------------------------------------------------------------------------
# ПРИМЕРЫ ДЛЯ ТЕСТА САНАЙЗЕРА (libre_macros_sanitize_lib) — НЕ ВЫЗЫВАТЬ
# ---------------------------------------------------------------------------
# Эти функции намеренно содержат запрещённые конструкции (сеть, base64/шифро-
# контент, запись в ФС, обход песочницы). Нужны как образцы для проверки
# lm_sanitize_code / lm_sanitize_load_file_function (см. docs/16_CODE_SANITIZE.md).
#
# НЕ подключать в лист параметров (ref / code). После интеграции санации в
# collect_workbooks загрузка через file# должна отклоняться.
# Сигнатура финальной фазы: (doc, sheets, data_ranges, header_row_ranges, *extra).
# ---------------------------------------------------------------------------


def pp_sanitize_test_network_socket( doc, sheets, data_ranges, header_row_ranges, *extra_args ):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: сетевой сокет (TCP). Санайзер должен банить
    (NETWORK / HEUR_NETWORK). См. docs/16_CODE_SANITIZE.md.
    """
    # --- тест санайзера: запрещённый сетевой API ---
    raise RuntimeError("sanitize test only — не вызывать")
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.connect(("127.0.0.1", 9))
    sock.send(b"x")
    sock.close()


def pp_sanitize_test_network_urlopen( doc, sheets, data_ranges, header_row_ranges, *extra_args ):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: HTTP через urlopen. Санайзер должен банить (NETWORK).
    """
    # --- тест санайзера: запрещённый HTTP-клиент ---
    raise RuntimeError("sanitize test only — не вызывать")
    urlopen("http://127.0.0.1/")


def pp_sanitize_test_base64_blob( doc, sheets, data_ranges, header_row_ranges, *extra_args ):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: base64 + длинный литерал-blob. Санайзер должен
    банить (CRYPTO / ENCODED_BLOB).
    """
    # --- тест санайзера: base64 и шифроподобный литерал ---
    raise RuntimeError("sanitize test only — не вызывать")
    # намеренно длинная «похожая на base64» строка (≥40 символов)
    payload = "QUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUE="
    return b64decode(payload)


def pp_sanitize_test_fs_write( doc, sheets, data_ranges, header_row_ranges, *extra_args ):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: запись в файловую систему. Санайзер должен банить
    (FS_WRITE / HEUR_FS_WRITE).
    """
    # --- тест санайзера: запрещённая запись на диск ---
    raise RuntimeError("sanitize test only — не вызывать")
    f = open("/tmp/libre_macros_sanitize_test.txt", "w")
    f.write("sanitize-test")
    f.close()


def pp_sanitize_test_sandbox_escape( doc, sheets, data_ranges, header_row_ranges, *extra_args ):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: обход песочницы через object model. Санайзер
    должен банить (BANNED_ATTR / __subclasses__).
    """
    # --- тест санайзера: обход builtins blocklist ---
    raise RuntimeError("sanitize test only — не вызывать")
    return ().__class__.__bases__[0].__subclasses__()
