MACRO_VERSION = "3.10.727"
# region Справка — namespace и хелперы (свернуть: ▼ слева или Ctrl+Shift+[)
"""
functions_pp.py — пользовательские функции постобработки (диапазон / строка).

Формат на листе «Параметры_Объединения*» (только JSON в C для встроенных функций)
------------------------------------------------------------------------
Диапазон:
  A = Постобработка_Диапазон
  B = функция_плагин
  C = [{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_range_zagotovka","extra":"42,тест","sheet":"Сводка"}]
  D = (пусто, если extra в JSON)

Строка:
  A = Постобработка_Строка
  B = функция_плагин
  C = [{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_row_zagotovka","sheet":"Итог"}]

Допуск: allow_env / allow_global — имена OS env и зашифрованной глобальной;
пусто в визарде/JSON → MERGE_ALLOW_PLUGINS / Merge_Allow_Plugins.
Значения env и глобальной должны совпасть, иначе сбор прерывается.

Legacy: B = functions_pp.py#имя без JSON — по-прежнему работает.

Сигнатуры def pp_*
-------------------
  Диапазон: (doc, sheet, data_range, header_row_range, *extra_args)
  Строка:   (doc, sheet, data_row_range, header_row_range, row_index, *extra_args)

Импорты в этом файле для хелперов макроса не нужны: при загрузке def через
«функция_плагин» код компилируется с namespace _merge_pp_postprocess_eval_namespace()
(collect_workbooks.py). Те же имена доступны в inline lambda/def в колонке B.

Хелперы и утилиты
-----------------
  _merge_pp_range_address(range)
      → (start_col, start_row, end_col, end_row), индексы 0-based.

  _merge_pp_foreach_row(sheet, start_row, end_row, callback, status_prefix="")
      Цикл по строкам данных с merge_ui_yield / merge_ui_status_set.

  _merge_pp_number_format_key(doc, fmt_string, lang="ru", country="RU")
      Ключ числового формата Calc.

  cell_text(cell)              Текст ячейки (строка).
  lo_color_rgb(red, green, blue)   Цвет LibreOffice (int).

  merge_pp_active_context()    dict контекста постобработки (границы, total_row, …).
      Если функция добавила/удалила строки или столбцы и следующие шаги должны
      увидеть новый диапазон, обновите ctx: end_col / end_row / data_start.
      Движок перечитает data_range/header_row_range перед следующим range-шагом.

  merge_pp_user_log(doc, sheet, message, status)
      Запись в лист «Сбор_книг_лог» (status: ok | пропуск | ошибка | …).

  merge_ui_yield(counter)      Отдать управление UI (каждые MERGE_UI_YIELD_EVERY_ROWS).
  merge_ui_status_set(n, text) Обновить StatusIndicator (каждые MERGE_UI_STATUS_EVERY_ROWS).

  traceback                    Модуль traceback (для отладки в except).

  MERGE_UI_STATUS_EVERY_ROWS   Интервал статуса (по умолчанию 300).
  MERGE_UI_YIELD_EVERY_ROWS    Интервал yield UI (по умолчанию 10).

Встроенные колбэки merge_pp_* (из libre_macros_lib)
-----------------------------------------------------
После install.sh доступны все merge_pp_range_* и merge_pp_row_* из библиотеки, например:
  merge_pp_range_thin_grid_borders, merge_pp_range_rename_sheet,
  merge_pp_range_pivot_table, merge_pp_range_sort_data, merge_pp_range_colorize_data,
  merge_pp_row_zebra_stripes, merge_pp_row_height_pad, merge_pp_row_convert_numeric_strings,
  merge_pp_vlookup_join_sheets, …
Полный перечень — LM_PP_PUBLIC_NAMES в pythonpath/libre_macros_lib.py
(имя merge_* = lm_* без префикса lm_).

Ключи карт — вызываются как функции в lambda/def
------------------------------------------------
Диапазон (MERGE_RESULT_POSTPROCESS_RANGE_MAP):
  сетка, тонкая_сетка, толстая_сетка, заголовок_плюс_высота, перенос, перенос_и_авто_высота,
  авто_высота, авто_ширина, высота_строки, левое_выравнивание, отступ, шрифт, зебра_диапазон,
  формат_деньги, формат_даты, формат_столбцы, подсветка_по_порогу, условное_форматирование,
  градиент, удалить_столбцы, конкатенация_столбцов, применить_формулу, удаление_строк,
  пропуск_пустых_строк, сброс_стрипов, группировка_по_столбцу, закрепить_заголовок,
  ширина_столбцов, автофильтр, стиль_печати, заполнение_вниз, заполнение_вниз_вычислить,
  копировать_значения, переименовать_лист, переименовать_столбцы, переставить_столбцы,
  количество_значений, копировать_лист,
  удалить_дубликаты, группировать_строки,
  сводная_таблица, сортировка, раскрасить_блоки.

Строка (MERGE_RESULT_POSTPROCESS_ROW_MAP):
  высота_строки, серые_служебные, полосы_по_пути, чередующиеся_границы,
  копировать_формат_заголовка, подсветка_по_заголовку, цвет_текста_по_значению,
  жирный_по_пути, преобразовать_числа.

Пример: lambda doc,sheet,dr,hr: переименовать_лист(doc, sheet, dr, hr, "Итог")

Ограничения builtins
--------------------
В eval/exec запрещены: __import__, eval, exec, compile, open, input, breakpoint,
globals, locals, vars, dir и др. (_MERGE_PP_EVAL_BUILTIN_BLOCKLIST).

Подробнее: macro-lib/USER_PP_FUNCTIONS_GUIDE.md, macro-lib/POSTPROCESS_GUIDE.md
"""
# endregion Справка — namespace и хелперы

import traceback


# region LLM — промпт для генерации пользовательской функции (свернуть)
LLM_PROMPT_PP_RULES = u"""Ты пишешь пользовательскую функцию постобработки для макроса LibreOffice Calc «Сбор книг».

Контекст
--------
Файл: functions_pp.py (или functions_final.py для финальной фазы).
Загрузка: B=функция_плагин, C=JSON с полем ref (functions_pp.py#имя) или code (inline).
Namespace при компиляции: _merge_pp_postprocess_eval_namespace() — импорты в файле не нужны.

Сигнатуры (обязательно)
-----------------------
Диапазон (Постобработка_Диапазон):
  def pp_имя(doc, sheet, data_range, header_row_range, *extra_args):

Строка (Постобработка_Строка):
  def pp_имя(doc, sheet, data_row_range, header_row_range, row_index, *extra_args):

Параметры extra_args — строки из JSON extra / колонки D (через запятую).

Доступные имена (без import)
----------------------------
  _merge_pp_range_address(range) → (sc, sr, ec, er) 0-based
  _merge_pp_foreach_row(sheet, start_row, end_row, callback, status_prefix="")
  cell_text(cell), lo_color_rgb(r,g,b)
  merge_pp_active_context(), merge_pp_user_log(doc, sheet, msg, status)
  merge_ui_yield(counter), merge_ui_status_set(n, text)
  MERGE_UI_STATUS_EVERY_ROWS, MERGE_UI_YIELD_EVERY_ROWS
  traceback
  Все merge_pp_range_* и merge_pp_row_* из библиотеки (тонкая_сетка, сортировка, …)
  Ключи карт как функции: lambda doc,s,dr,hr: переименовать_лист(doc,s,dr,hr,"Итог")

Запрещено в inline-коде (eval/exec)
-----------------------------------
  __import__, eval, exec, compile, open, input, breakpoint,
  globals, locals, vars, dir, getattr на опасных объектах — см. _MERGE_PP_EVAL_BUILTIN_BLOCKLIST.

Стиль кода
----------
- Python 2.7 / LibreOffice Scripting: без f-strings; while вместо for по индексам где в образце.
- Индексы ячеек 0-based; Calc-формулы — 1-based.
- Ошибки: try/except, merge_pp_user_log(..., "ошибка"), raise после лога.
- Долгие циклы: merge_ui_yield каждые ~10 строк, merge_ui_status_set каждые ~300.
- Не менять глобальное состояние макроса кроме ячеек целевого листа.
- Если функция меняет размер диапазона (вставка/удаление строк, новые столбцы),
  обнови merge_pp_active_context(): ctx["end_col"], ctx["end_row"], при нужде
  ctx["data_start"]. Следующий шаг цепочки получит уже новый data_range.
- Не вызывать doc.close(), не удалять листы (для постобработки диапазона/строки).


Что должна выполнять функция:
----------
- 
- 
-


Вывод
-----
Верни только тело функции def pp_... (или полный def), без markdown-ограждений.
Имя функции должно совпадать с ref в JSON (functions_pp.py#имя).
"""


# endregion LLM — промпт для генерации пользовательской функции


def pp_range_zagotovka(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Заготовка диапазонной постобработки: один проход по строкам данных,
    статус в строке Calc — раз в MERGE_UI_STATUS_EVERY_ROWS, отдача UI — чаще.

    Сигнатура как у merge_pp_range_*: (doc, sheet, data_range, header_row_range, *extra_args).
    Параметры extra — JSON поле extra или колонка D (через запятую).

    Пример JSON в C:
        [{"v":1,"fn":"функция_плагин","ref":"functions_pp.py#pp_range_zagotovka","extra":"42","sheet":"Сводка"}]

    Замените тело цикла while r <= er на свою логику.

    Если функция меняет размер диапазона, обновите контекст, например:
        ctx = merge_pp_active_context()
        if ctx is not None:
            ctx["end_col"] = max(int(ctx.get("end_col", end_col)), new_end_col)
            ctx["end_row"] = new_end_row
            ctx["data_start"] = header_row + 1
    Тогда следующий range-шаг получит уже новый data_range/header_row_range.
    """
    # Пример работы с extra_args (параметры из колонки C)
    # extra_args — кортеж строк, например ("42",) или ("100", "200", "test")
    param1 = extra_args[0] if len(extra_args) > 0 else ""
    param2 = extra_args[1] if len(extra_args) > 1 else ""

    try:
        start_col, start_row, end_col, end_row = _merge_pp_range_address(data_range)
        _hdr_col, header_row, _hdr_end_col, _hdr_end_row = _merge_pp_range_address(
            header_row_range
        )
        if start_row <= header_row:
            start_row = header_row + 1
        if end_row < start_row:
            merge_pp_user_log(doc, sheet, "нет строк данных", "пропуск")
            return

        row_index = start_row
        rows_done = 0
        while row_index <= end_row:
            # --- Твоя логика на одну строку данных (0-based row_index) ---
            # Пример использования extra_args: проверка условия из параметра
            col_index = start_col
            while col_index <= end_col:
                cell = sheet.getCellByPosition(col_index, row_index)
                # Пример: закрасить ячейку если param1 == "highlight"
                # if param1 == "highlight":
                #     cell.CellBackColor = 16776960  # yellow
                col_index = col_index + 1
            rows_done = rows_done + 1
            # Статус: 1-я строка, каждые MERGE_UI_STATUS_EVERY_ROWS (500), последняя
            is_last_row = row_index == end_row
            if rows_done <= 1 or is_last_row or (
                rows_done % MERGE_UI_STATUS_EVERY_ROWS == 0
            ):
                merge_ui_status_set(
                    rows_done,
                    "диапазон — строка %d/%d"
                    % (row_index + 1, end_row + 1),
                )
            # Отдача UI (внутри — каждые MERGE_UI_YIELD_EVERY_ROWS, по умолчанию 10)
            merge_ui_yield(counter=rows_done)
            row_index = row_index + 1
        # Если выше были добавлены/удалены строки или столбцы, обновите контекст
        # для следующих range-шагов pipeline.
        # ctx = merge_pp_active_context()
        # if ctx is not None:
        #     ctx["end_col"] = max(int(ctx.get("end_col", end_col)), new_end_col)
        #     ctx["end_row"] = new_end_row
        #     ctx["data_start"] = header_row + 1
        total_rows = end_row - start_row + 1
        # В лог можно вывести использованные параметры из extra_args
        params_info = ", extra_args: %s" % str(extra_args) if extra_args else ""
        merge_pp_user_log(
            doc,
            sheet,
            "готово, строк %d (статус каждые %d)%s"
            % (total_rows, MERGE_UI_STATUS_EVERY_ROWS, params_info),
            "ok",
        )
    except Exception as err:
        tb = ""
        try:
            tb = traceback.format_exc()
        except Exception:
            pass
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s\n%s" % (err, tb), "ошибка")
        raise


def pp_stroka_itogo(doc, sheet, data_range, header_row_range):
    try:
        addr = data_range.getRangeAddress()
        hdr = header_row_range.getRangeAddress()
        sc = addr.StartColumn
        ec = addr.EndColumn
        sr = addr.StartRow
        er = addr.EndRow
        if sr <= hdr.EndRow:
            sr = hdr.EndRow + 1
        if er < sr:
            return
        total_r = er + 1
        first_1 = sr + 1
        last_1 = er + 1
        sample_r = sr
        hdr_r = hdr.StartRow
        header_bg = (217 << 16) | (217 << 8) | 217
        char_weight_bold = 150
        vert_center = 2
        hori_center = 2
        hori_right = 3
        sum_fmt_key = _merge_pp_number_format_key(doc, "# ##0,00")
        c = sc
        while c <= ec:
            col_numeric = False
            data_cell = sheet.getCellByPosition(c, sample_r)
            kind = None
            try:
                ctype = int(data_cell.getType())
            except Exception:
                ctype = -1
            if ctype == 1:
                kind = True
            elif ctype == 3:
                try:
                    v = data_cell.Value
                    if v is not None and isinstance(v, (int, float)):
                        kind = True
                except Exception:
                    pass
            if kind is None and ctype in (2, 3, -1):
                txt = ""
                try:
                    s = data_cell.String
                    if s is not None and str(s).strip() != "":
                        txt = str(s).strip()
                except Exception:
                    pass
                if txt == "":
                    try:
                        v = data_cell.Value
                        if v is not None and isinstance(v, (int, float)) and v != 0:
                            kind = True
                    except Exception:
                        pass
                if kind is None and txt != "":
                    has_letter = False
                    ti = 0
                    while ti < len(txt):
                        if txt[ti].isalpha():
                            has_letter = True
                        ti = ti + 1
                    if not has_letter:
                        norm = txt
                        for sep in (" ", "\t", "\xa0", "\u202f", "\u2009"):
                            norm = norm.replace(sep, "")
                        norm = norm.replace(",", ".")
                        try:
                            float(norm)
                            kind = True
                        except ValueError:
                            kind = False
                    else:
                        kind = False
            if kind is True:
                col_numeric = True
            n = c + 1
            letters = ""
            while n > 0:
                n, rem = divmod(n - 1, 26)
                letters = chr(65 + rem) + letters
            cell = sheet.getCellByPosition(c, total_r)
            ref_cell = sheet.getCellByPosition(c, er)
            if c == sc:
                cell.String = "Итого"
            elif col_numeric:
                cell.Formula = "=SUM(%s%d:%s%d)" % (letters, first_1, letters, last_1)
                cell.NumberFormat = sum_fmt_key
                try:
                    cell.HoriJustify = hori_right
                except Exception:
                    pass
            else:
                cell.String = ""
            for side in ("TopBorder", "BottomBorder", "LeftBorder", "RightBorder"):
                try:
                    setattr(cell, side, getattr(ref_cell, side))
                except Exception:
                    pass
            if c == sc or not col_numeric:
                try:
                    cell.HoriJustify = sheet.getCellByPosition(c, hdr_r).HoriJustify
                except Exception:
                    try:
                        cell.HoriJustify = hori_center
                    except Exception:
                        pass
            c = c + 1
        total_range = sheet.getCellRangeByPosition(sc, total_r, ec, total_r)
        total_range.IsCellBackgroundTransparent = False
        total_range.CellBackColor = header_bg
        total_range.CharWeight = char_weight_bold
        try:
            total_range.VertJustify = vert_center
        except Exception:
            pass
        ctx = merge_pp_active_context()
        if ctx is not None:
            ctx["total_row"] = total_r
            # Добавили строку «Итого»: следующие range-шаги должны видеть
            # обновлённую высоту листа и новый конец диапазона.
            if total_r > int(ctx.get("end_row", total_r)):
                ctx["end_row"] = total_r
            if ec > int(ctx.get("end_col", ec)):
                ctx["end_col"] = ec
            ctx["data_start"] = hdr_r + 1
    except Exception as err:
        tb = ""
        try:
            tb = traceback.format_exc()
        except Exception:
            pass
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s\n%s" % (err, tb), "ошибка")
        raise


def pp_range_fill_down(doc, sheet, data_range, header_row_range, *extra_args):
    """
    Пользовательский вариант заполнения вниз: пустые ячейки → значение из строки выше.

    Сигнатура как у merge_pp_range_*.
    Параметры из JSON extra / колонки C: имена столбцов — в JSON блоке «заполнение_вниз»
    или в extra плагина. Для этой функции: extra_args — список столбцов.

    Использование: ref functions_pp.py#pp_range_fill_down в JSON плагина.

    Если вы расширите пример и будете добавлять столбцы/строки, не забудьте
    обновить merge_pp_active_context() в конце функции.
    """
    try:
        sc, sr, ec, er = _merge_pp_range_address(data_range)
        _hc, hdr_r, _he, _hr = _merge_pp_range_address(header_row_range)
        if sr <= hdr_r:
            sr = hdr_r + 1
        if er < sr:
            merge_pp_user_log(doc, sheet, "нет строк данных", "пропуск")
            return

        # Разобрать параметры C
        spec = ", ".join(str(x) for x in extra_args) if extra_args else ""
        cols_to_fill = []
        if spec:
            for part in spec.replace(";", ",").split(","):
                p = part.strip()
                if not p:
                    continue
                # Поиск столбца по номеру, букве или имени заголовка
                col_idx = None
                try:
                    col_idx = int(p) - 1  # 1-based → 0-based
                except ValueError:
                    pass
                if col_idx is None and len(p) <= 3 and p.isalpha():
                    # Буква столбца (A, AA...)
                    col_idx = 0
                    for ch in p.upper():
                        col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
                    col_idx -= 1
                if col_idx is None:
                    # Поиск по заголовкам
                    c = sc
                    while c <= ec:
                        hdr_text = cell_text(sheet.getCellByPosition(c, hdr_r))
                        if p.lower() in hdr_text.lower():
                            col_idx = c
                            break
                        c = c + 1
                if col_idx is not None and col_idx not in cols_to_fill:
                    cols_to_fill.append(col_idx)

        filled = 0
        for col in cols_to_fill:
            row = sr + 1
            while row <= er:
                cell = sheet.getCellByPosition(col, row)
                if cell_text(cell) == "":
                    src = sheet.getCellByPosition(col, row - 1)
                    try:
                        cell.String = src.String
                        filled = filled + 1
                    except Exception:
                        try:
                            cell.Value = src.Value
                            filled = filled + 1
                        except Exception:
                            pass
                row = row + 1

        merge_pp_user_log(
            doc, sheet,
            "заполнение_вниз (functions_pp.py): столбцов %d, заполнено %d" % (len(cols_to_fill), filled),
            "ok"
        )
    except Exception as err:
        tb = ""
        try:
            tb = traceback.format_exc()
        except Exception:
            pass
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s\n%s" % (err, tb), "ошибка")
        raise


def pp_range_color_nth_col(doc, sheet, data_range, header_row_range, *extra_args):
    try:
        start_col, start_row, end_col, end_row = _merge_pp_range_address(data_range)
        _hdr_col, header_row, _hdr_end_col, _hdr_end_row = _merge_pp_range_address(
            header_row_range
        )
        if start_row <= header_row:
            start_row = header_row + 1
        if end_row < start_row:
            merge_pp_user_log(doc, sheet, "нет строк данных", "пропуск")
            return

        # --- Парсинг extra_args: кратность и HEX-цвет ---
        # Разделители: запятая и точка с запятой (могут быть в одном аргументе или между)
        raw = ""
        i = 0
        while i < len(extra_args):
            if i > 0:
                raw = raw + ","
            raw = raw + str(extra_args[i])
            i = i + 1

        # Нормализуем разделители
        normalized = raw.replace(";", ",")
        tokens = []
        for part in normalized.split(","):
            p = part.strip()
            if p != "":
                tokens.append(p)

        # Значения по умолчанию
        step = 3
        color_light_blue = lo_color_rgb(173, 216, 230)  # #ADD8E6
        color_hex_str = "#ADD8E6"

        if len(tokens) >= 1:
            try:
                step = int(tokens[0])
                if step <= 0:
                    step = 3
            except Exception:
                merge_pp_user_log(
                    doc, sheet,
                    "некорректная кратность '%s', используется 3" % tokens[0],
                    "предупреждение"
                )
                step = 3

        if len(tokens) >= 2:
            hex_val = tokens[1].strip().lstrip("#").upper()
            valid_hex = True
            if len(hex_val) == 6:
                try:
                    r = int(hex_val[0:2], 16)
                    g = int(hex_val[2:4], 16)
                    b = int(hex_val[4:6], 16)
                    color_light_blue = lo_color_rgb(r, g, b)
                    color_hex_str = "#" + hex_val
                except Exception:
                    valid_hex = False
            else:
                valid_hex = False
            if not valid_hex:
                merge_pp_user_log(
                    doc, sheet,
                    "некорректный HEX '%s', используется #ADD8E6" % tokens[1],
                    "предупреждение"
                )
                color_light_blue = lo_color_rgb(173, 216, 230)
                color_hex_str = "#ADD8E6"

        row_index = start_row
        rows_done = 0
        while row_index <= end_row:
            col_index = start_col
            while col_index <= end_col:
                # Каждый N-й столбец диапазона (индексы step-1, 2*step-1, ...)
                if (col_index - start_col) % step == (step - 1):
                    cell = sheet.getCellByPosition(col_index, row_index)
                    cell.CellBackColor = color_light_blue
                col_index = col_index + 1

            rows_done = rows_done + 1
            is_last_row = row_index == end_row
            if rows_done <= 1 or is_last_row or (
                rows_done % MERGE_UI_STATUS_EVERY_ROWS == 0
            ):
                merge_ui_status_set(
                    rows_done,
                    "заливка каждого %d-го столбца — строка %d/%d"
                    % (step, row_index + 1, end_row + 1),
                )
            merge_ui_yield(counter=rows_done)
            row_index = row_index + 1

        total_rows = end_row - start_row + 1
        merge_pp_user_log(
            doc,
            sheet,
            "готово, строк %d, шаг %d, цвет %s" % (total_rows, step, color_hex_str),
            "ok",
        )
        # Если на базе этого примера вы добавляете столбцы справа, обновите:
        # ctx = merge_pp_active_context()
        # if ctx is not None:
        #     ctx["end_col"] = new_end_col
        #     ctx["end_row"] = end_row
    except Exception as err:
        tb = ""
        try:
            tb = traceback.format_exc()
        except Exception:
            pass
        merge_pp_user_log(doc, sheet, "ОШИБКА: %s\n%s" % (err, tb), "ошибка")
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
# ---------------------------------------------------------------------------


def pp_sanitize_test_network_socket(doc, sheet, data_range, header_row_range, *extra_args):
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


def pp_sanitize_test_network_urlopen(doc, sheet, data_range, header_row_range, *extra_args):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: HTTP через urlopen. Санайзер должен банить (NETWORK).
    """
    # --- тест санайзера: запрещённый HTTP-клиент ---
    raise RuntimeError("sanitize test only — не вызывать")
    urlopen("http://127.0.0.1/")


def pp_sanitize_test_base64_blob(doc, sheet, data_range, header_row_range, *extra_args):
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


def pp_sanitize_test_fs_write(doc, sheet, data_range, header_row_range, *extra_args):
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


def pp_sanitize_test_sandbox_escape(doc, sheet, data_range, header_row_range, *extra_args):
    """
    ТЕСТ САНАЙЗЕРА / ПРИМЕР — НЕ ИСПОЛЬЗОВАТЬ В РАБОЧИХ ШАБЛОНАХ.

    Заведомо небезопасный код: обход песочницы через object model. Санайзер
    должен банить (BANNED_ATTR / __subclasses__).
    """
    # --- тест санайзера: обход builtins blocklist ---
    raise RuntimeError("sanitize test only — не вызывать")
    return ().__class__.__bases__[0].__subclasses__()
