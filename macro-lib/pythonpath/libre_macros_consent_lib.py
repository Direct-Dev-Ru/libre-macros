# -*- coding: utf-8 -*-
"""
Согласие с условиями использования и предупреждение о функциях плагинов.

Файлы в каталоге пресетов (~/.config/libre-macros/collect_workbooks/):
  terms_of_use.txt      — текст условий
  user_agreement.json   — журнал принятий
"""

from __future__ import print_function
MACRO_VERSION = "3.10.706"
import datetime
import json
import os
import platform
import re
import socket
import sys
import traceback

import uno
import unohelper
from com.sun.star.awt import XActionListener

try:
    unicode
except NameError:
    unicode = str

CONSENT_TERMS_VERSION = 4
CONSENT_JSON_VERSION = 1
CONSENT_TERMS_FILE = u"terms_of_use.txt"
CONSENT_ACCEPTANCE_FILE = u"user_agreement.json"
CONSENT_DATA_SUBDIR = u"collect_workbooks"
CONSENT_TERMS_MARKER = u"# TERMS_VERSION=%d\n" % CONSENT_TERMS_VERSION
CONSENT_OFFICE_NAME = u"LibreOffice/AlterOffice"


def lm_consent_terms_marker_for_version(version):
    return u"# TERMS_VERSION=%d\n" % int(version)


def lm_consent_parse_terms_version_header(text):
    """Разбор первой строки # TERMS_VERSION=N → (версия|None, тело без заголовка)."""
    text = unicode(text or u"")
    if text.startswith(u"\ufeff"):
        text = text.lstrip(u"\ufeff")
    first_line, _sep, rest = text.partition(u"\n")
    m = re.match(r"^#\s*TERMS_VERSION=(\d+)\s*$", first_line.strip())
    if m:
        return int(m.group(1)), rest
    return None, text


def lm_consent_required_terms_version():
    """Актуальная версия условий из terms_of_use.txt (или CONSENT_TERMS_VERSION)."""
    lm_consent_ensure_terms_file()
    path = lm_consent_terms_path()
    try:
        with open(path, "rb") as f:
            head = f.read(256).decode("utf-8", "replace")
        ver, _body = lm_consent_parse_terms_version_header(head)
        if ver is not None:
            return int(ver)
    except Exception:
        pass
    return int(CONSENT_TERMS_VERSION)


def lm_consent_data_dir():
    """Каталог данных (тот же, что и merge_param_presets.json)."""
    if sys.platform == "win32":
        base = os.environ.get("APPDATA", os.path.expanduser("~"))
    else:
        base = os.environ.get(
            "XDG_CONFIG_HOME", os.path.join(os.path.expanduser("~"), ".config")
        )
    path = os.path.join(base, "libre-macros", CONSENT_DATA_SUBDIR)
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


def lm_consent_terms_path():
    return os.path.join(lm_consent_data_dir(), CONSENT_TERMS_FILE)


def lm_consent_acceptance_path():
    return os.path.join(lm_consent_data_dir(), CONSENT_ACCEPTANCE_FILE)


def lm_consent_default_terms_text():
    office = CONSENT_OFFICE_NAME
    return (
        (u"УСЛОВИЯ ИСПОЛЬЗОВАНИЯ МАКРОСА «ОБЪЕДИНЕНИЕ КНИГ» (%s Calc)\n" % office)
        + u"\n"
        + u"1. НАЗНАЧЕНИЕ\n"
        + u"Настоящий программный макрос предназначен для автоматизации типовых операций "
        + u"сбора, объединения, разделения и постобработки данных в электронных таблицах. Он помогает "
        + u"сократить ручной труд, но не заменяет контроль пользователя над результатом.\n"
        + u"\n"
        + u"2. ОГРАНИЧЕНИЕ ОТВЕТСТВЕННОСТИ РАЗРАБОТЧИКА\n"
        + u"Макрос проходил тестирование в ходе разработки, однако автор не гарантирует:\n"
        + u"  • отсутствие ошибок при сборе, преобразовании или записи данных;\n"
        + u"  • корректность результата при определенных сочетаниях параметров и порядке их следования;\n"
        + u"  • соответствие результата вашим бизнес-правилам без дополнительной проверки с вашей стороны.\n"
        + u"Некоторые комбинации настройки листа «Параметры_Объединения» могут привести к потере фрагментов "
        + u"данных, дублированию строк, неверным формулам, некорректному форматированию "
        + u"или иному непредсказуемому поведению.\n"
        + u"\n"
        + u"3. ИСХОДНЫЕ ДАННЫЕ И ФАЙЛ РЕЗУЛЬТАТА\n"
        + u"Макрос спроектирован так, чтобы по возможности не изменять файлы-источники. "
        + u"Тем не менее вы обязаны самостоятельно убедиться перед запуском, что:\n"
        + u"  • пути к источникам указаны верно;\n"
        + u"  • резервные копии важных файлов созданы до запуска;\n"
        + u"  • сохранены и закрыты все открытые книги кроме книги результата с парметрами объединения;\n"
        + u"после запуска, что:\n"        
        + u"  • содержимое книги результата проверено визуально и по контрольным суммам/выборкам.\n"
        + u"Ответственность за корректность итоговых данных в файле результата несёт "
        + u"пользователь, запускающий макрос.\n"
        + u"\n"
        + u"4. ПОЛЬЗОВАТЕЛЬСКИЕ ФУНКЦИИ И ПЛАГИНЫ\n"
        + u"На листе параметров допускается подключение пользовательского кода: функции "
        + u"плагинов (functions_pp.py, functions_final.py, functions_*.py), произвольный "
        + u"Python-код и расширения, заданные вами при настройке.\n"
        + (u"Такой код выполняется от вашего имени в среде %s. Разработчик макроса " % office)
        + u"не проверяет сторонний код и не несёт ответственности за его действия, побочные "
        + u"эффекты, утечку данных, повреждение книг или некорректные вычисления.\n"
        + u"Ответственность за подключённые функции плагинов и произвольный код полностью "
        + u"лежит на пользователе, указавшем их в параметрах работы макроса.\n"
        + u"\n"
        + u"5. РЕКОМЕНДАЦИИ ПЕРЕД ЗАПУСКОМ\n"
        + u"  • Сохраните книгу результата и источники.\n"
        + u"  • Закройте все открытые книги кроме книги результата с параметрами объединения.\n"
        + u"  • Проверьте параметры на тестовой копии данных.\n"
        + u"  • Просмотрите журнал «Сбор_книг_лог» после выполнения.\n"
        + u"  • Не используйте макрос для критичных отчётов без независимой сверки с исходными данными.\n"
        + u"\n"
        + u"6. ПРИНЯТИЕ УСЛОВИЙ\n"
        + u"Нажимая кнопку «Принимаю», вы подтверждаете, что:\n"
        + u"  • ознакомились с настоящими условиями;\n"
        + u"  • понимаете риски автоматизированной обработки данных;\n"
        + u"  • принимаете на себя ответственность за параметры запуска, подключённый код "
        + u"и полученный результат;\n"
        + u"  • согласны продолжить использование макроса на изложенных условиях.\n"
        + u"\n"
        + u"Без принятия условий запуск макроса «Объединение книг» невозможен.\n"
    )


def lm_consent_terms_file_is_current():
    """Файл terms_of_use.txt соответствует встроенной CONSENT_TERMS_VERSION."""
    path = lm_consent_terms_path()
    if not os.path.isfile(path):
        return False
    try:
        with open(path, "rb") as f:
            head = f.read(96).decode("utf-8", "replace")
        ver, _body = lm_consent_parse_terms_version_header(head)
        return ver == int(CONSENT_TERMS_VERSION)
    except Exception:
        return False


def lm_consent_ensure_terms_file():
    path = lm_consent_terms_path()
    if lm_consent_terms_file_is_current():
        return path
    text = CONSENT_TERMS_MARKER + lm_consent_default_terms_text()
    bundled = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), u"terms_of_use_ru.txt"
    )
    if os.path.isfile(bundled):
        try:
            with open(bundled, "rb") as f:
                raw = f.read().decode("utf-8", "replace")
            if raw.strip() != u"":
                text = CONSENT_TERMS_MARKER + raw.lstrip(u"\ufeff")
        except Exception:
            pass
    try:
        lm_consent_data_dir()
        with open(path, "wb") as f:
            f.write(text.encode("utf-8"))
    except Exception:
        pass
    return path


def lm_consent_load_terms_text():
    lm_consent_ensure_terms_file()
    path = lm_consent_terms_path()
    try:
        with open(path, "rb") as f:
            raw = f.read()
        for enc in ("utf-8", "utf-8-sig", "cp1251"):
            try:
                text = raw.decode(enc)
                _ver, body = lm_consent_parse_terms_version_header(text)
                if _ver is not None:
                    return body
                return text
            except Exception:
                pass
        text = raw.decode("utf-8", "replace")
        _ver, body = lm_consent_parse_terms_version_header(text)
        if _ver is not None:
            return body
        return text
    except Exception:
        return lm_consent_default_terms_text()


def lm_consent_collect_system_info(macro_version=u"", doc=None):
    info = {
        u"platform": unicode(sys.platform or u""),
        u"os_name": unicode(platform.system() or u""),
        u"os_release": unicode(platform.release() or u""),
        u"os_version": unicode(platform.version() or u""),
        u"machine": unicode(platform.machine() or u""),
        u"python_version": unicode(sys.version or u"").split()[0],
        u"os_user": unicode(os.environ.get("USER") or os.environ.get("USERNAME") or u""),
        u"hostname": u"",
        u"macro_version": unicode(macro_version or u""),
        u"terms_version": lm_consent_required_terms_version(),
    }
    try:
        info[u"hostname"] = unicode(socket.gethostname() or u"")
    except Exception:
        pass
    if doc is not None and hasattr(doc, "getPropertyValue"):
        for prop in (u"Title", u"URL"):
            try:
                info[u"doc_" + prop.lower()] = unicode(doc.getPropertyValue(prop) or u"")
            except Exception:
                pass
    return info


def lm_consent_load_acceptance_data():
    path = lm_consent_acceptance_path()
    if not os.path.isfile(path):
        return {u"version": CONSENT_JSON_VERSION, u"acceptances": []}
    try:
        with open(path, "rb") as f:
            data = json.loads(f.read().decode("utf-8"))
        if not isinstance(data, dict):
            return {u"version": CONSENT_JSON_VERSION, u"acceptances": []}
        if not isinstance(data.get(u"acceptances"), list):
            data[u"acceptances"] = []
        return data
    except Exception:
        return {u"version": CONSENT_JSON_VERSION, u"acceptances": []}


def lm_consent_save_acceptance(macro_version=u"", doc=None):
    data = lm_consent_load_acceptance_data()
    entry = lm_consent_collect_system_info(macro_version=macro_version, doc=doc)
    entry[u"accepted_at"] = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    entry[u"accepted"] = True
    data.setdefault(u"acceptances", []).append(entry)
    data[u"version"] = CONSENT_JSON_VERSION
    path = lm_consent_acceptance_path()
    try:
        lm_consent_data_dir()
        with open(path, "wb") as f:
            f.write(
                json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True).encode(
                    "utf-8"
                )
            )
        try:
            from libre_macros_global_settings_lib import notify_config_changed

            notify_config_changed()
        except Exception:
            pass
        return True
    except Exception as err:
        print(u"consent: не удалось сохранить %s: %s" % (path, err))
        return False


def lm_consent_is_accepted():
    """
    True — есть принятие с terms_version, совпадающей с # TERMS_VERSION в terms_of_use.txt.
    """
    lm_consent_ensure_terms_file()
    required = lm_consent_required_terms_version()
    data = lm_consent_load_acceptance_data()
    for item in reversed(data.get(u"acceptances") or []):
        if not isinstance(item, dict):
            continue
        if not item.get(u"accepted"):
            continue
        try:
            tv = int(item.get(u"terms_version") or 0)
        except Exception:
            tv = 0
        if tv == required:
            return True
    return False


def _consent_component_context():
    try:
        return uno.getComponentContext()
    except Exception:
        pass
    return None


def _consent_dialog_create_peer(dlg, toolkit, doc=None, parent_window=None):
    parents = []
    if parent_window is not None:
        parents.append(parent_window)
        try:
            peer = parent_window.getPeer()
            if peer is not None:
                parents.append(peer)
        except Exception:
            pass
    if doc is not None:
        try:
            w = doc.getCurrentController().getFrame().getContainerWindow()
            if w is not None and w not in parents:
                parents.append(w)
        except Exception:
            pass
    parents.append(None)
    i = 0
    while i < len(parents):
        try:
            dlg.createPeer(toolkit, parents[i])
            return True
        except Exception:
            pass
        i += 1
    return False


def _consent_dlg_model(ctx):
    sm = ctx.getServiceManager()
    return sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)


def _consent_dlg_add_fixed(dm, name, label, x, y, w, h, multiline=False):
    m = dm.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
    m.Name = str(name)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    m.Label = unicode(label)
    if multiline:
        m.MultiLine = True
    dm.insertByName(str(name), m)
    return m


def _consent_dlg_add_edit(dm, name, x, y, w, h, multiline=False, readonly=False):
    m = dm.createInstance("com.sun.star.awt.UnoControlEditModel")
    m.Name = str(name)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    if multiline:
        m.MultiLine = True
        try:
            m.VScroll = True
        except Exception:
            pass
    if readonly:
        m.ReadOnly = True
    dm.insertByName(str(name), m)
    return m


def _consent_dlg_add_button(dm, name, label, x, y, w, h):
    m = dm.createInstance("com.sun.star.awt.UnoControlButtonModel")
    m.Name = str(name)
    m.PositionX = int(x)
    m.PositionY = int(y)
    m.Width = int(w)
    m.Height = int(h)
    m.Label = unicode(label)
    dm.insertByName(str(name), m)
    return m


def _consent_report_error(on_error, text):
    print(u"consent: %s" % text)
    if on_error is not None:
        try:
            on_error(unicode(text))
        except Exception:
            pass


def lm_consent_dialog_show_terms( doc, macro_version=u"", toolkit=None, parent_window=None, on_error=None ):
    """
    Диалог условий использования.
    Возвращает True, если пользователь нажал «Принимаю».
    """
    try:
        if toolkit is None:
            _consent_report_error(
                on_error,
                u"Не удалось открыть диалог условий (Toolkit недоступен). Запуск отменён.",
            )
            return False
        ctx = _consent_component_context()
        if ctx is None:
            _consent_report_error(
                on_error,
                u"Не удалось открыть диалог условий (контекст UNO недоступен). Запуск отменён.",
            )
            return False

        terms = lm_consent_load_terms_text()
        dm = _consent_dlg_model(ctx)
        m = 12
        dw = 620
        btn_h = 24
        edit_h = 360
        dh = m + 16 + edit_h + 12 + btn_h + m
        dm.PositionX = 80
        dm.PositionY = 60
        dm.Width = dw
        dm.Height = dh
        dm.Title = u"Условия использования макроса «Объединение книг»"

        _consent_dlg_add_fixed(
            dm,
            "TermsHint",
            u"Внимательно прочитайте условия. Без принятия запуск макроса невозможен.",
            m,
            m,
            dw - 2 * m,
            16,
        )
        _consent_dlg_add_edit(
            dm,
            "TermsText",
            m,
            m + 18,
            dw - 2 * m,
            edit_h,
            multiline=True,
            readonly=True,
        )
        by = m + 18 + edit_h + 10
        _consent_dlg_add_button(dm, "CancelBtn", u"Отмена", dw - m - 210, by, 96, btn_h)
        _consent_dlg_add_button(dm, "AcceptBtn", u"Принимаю", dw - m - 104, by, 96, btn_h)

        sm = ctx.getServiceManager()
        dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dlg.setModel(dm)
        if not _consent_dialog_create_peer(dlg, toolkit, doc=doc, parent_window=parent_window):
            _consent_report_error(
                on_error,
                u"Не удалось создать окно диалога условий. Запуск отменён.",
            )
            return False

        terms_ctl = dlg.getControl("TermsText")
        if terms_ctl is not None:
            try:
                terms_ctl.setText(terms)
            except Exception:
                try:
                    terms_ctl.Model.Text = terms
                except Exception:
                    pass

        state = {"accepted": False}

        class _Handler(unohelper.Base, XActionListener):
            def disposing(self, event):
                pass

            def actionPerformed(self, event):
                try:
                    name = event.Source.getModel().Name
                except Exception:
                    return
                if name == "AcceptBtn":
                    state["accepted"] = True
                    dlg.endExecute()
                elif name == "CancelBtn":
                    state["accepted"] = False
                    dlg.endExecute()

        handler = _Handler()
        for btn in ("AcceptBtn", "CancelBtn"):
            try:
                dlg.getControl(btn).addActionListener(handler)
            except Exception:
                pass

        dlg.execute()
        if state["accepted"]:
            return lm_consent_save_acceptance(macro_version=macro_version, doc=doc)
        return False
    except Exception as err:
        print(u"consent: ошибка диалога условий:\n%s" % traceback.format_exc())
        _consent_report_error(
            on_error,
            u"Ошибка диалога условий использования:\n%s\nЗапуск отменён." % err,
        )
        return False


def lm_consent_dialog_show_plugin_warning( doc, plugin_refs, toolkit=None, parent_window=None, on_error=None ):
    """
    Предупреждение о пользовательских функциях плагинов.
    Возвращает True, если пользователь выбрал продолжить.
    """
    refs = [unicode(x).strip() for x in (plugin_refs or []) if unicode(x).strip()]
    if len(refs) == 0:
        return True
    try:
        if toolkit is None:
            _consent_report_error(
                on_error,
                u"Не удалось показать предупреждение о плагинах (Toolkit недоступен).",
            )
            return False
        ctx = _consent_component_context()
        if ctx is None:
            _consent_report_error(
                on_error,
                u"Не удалось показать предупреждение о плагинах (контекст UNO недоступен).",
            )
            return False

        lines = [
            u"ВНИМАНИЕ: в параметрах обнаружены пользовательские функции плагинов.",
            u"",
            (
                u"Такой код выполняется в среде %s от вашего имени. Разработчик "
                % CONSENT_OFFICE_NAME
                + u"основного макроса не проверяет сторонний код и не гарантирует корректность "
                + u"его работы. Любые ошибки, потеря данных или искажение результата — зона "
                + u"ответственности пользователя, указавшего эти функции на листе параметров."
            ),
            u"",
            u"Обнаружено:",
        ]
        i = 0
        while i < len(refs):
            lines.append(u"  • " + refs[i])
            i += 1
        lines.append(u"")
        lines.append(
            u"Продолжайте только если полностью доверяете указанному коду и понимаете риски."
        )
        body = u"\n".join(lines)

        dm = _consent_dlg_model(ctx)
        m = 12
        dw = 560
        btn_h = 24
        edit_h = 250
        # Soft-gray тема сама добавит высоту полосы-титла.
        dh = m + edit_h + 12 + btn_h + m
        dm.PositionX = 90
        dm.PositionY = 70
        dm.Width = dw
        dm.Height = dh
        dm.Title = u"Пользовательские функции плагинов"

        edit_m = _consent_dlg_add_edit(
            dm,
            "PluginWarnText",
            m,
            m,
            dw - 2 * m,
            edit_h,
            multiline=True,
            readonly=True,
        )
        try:
            from libre_macros_ui_theme import PLUGIN_WARN_TEXT_FONT_HEIGHT

            fh = int(PLUGIN_WARN_TEXT_FONT_HEIGHT)
            fd = edit_m.FontDescriptor
            fd.Height = fh
            edit_m.FontDescriptor = fd
        except Exception:
            try:
                edit_m.FontHeight = 12
            except Exception:
                pass
        by = m + edit_h + 10
        _consent_dlg_add_button(dm, "CancelBtn", u"Отмена", dw - m - 220, by, 104, btn_h)
        _consent_dlg_add_button(
            dm, "ContinueBtn", u"Продолжить", dw - m - 104, by, 96, btn_h
        )

        try:
            from libre_macros_ui_theme import apply_soft_gray_orange_theme

            apply_soft_gray_orange_theme(dm, title_text=dm.Title)
        except Exception:
            pass

        sm = ctx.getServiceManager()
        dlg = sm.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
        dlg.setModel(dm)
        if not _consent_dialog_create_peer(dlg, toolkit, doc=doc, parent_window=parent_window):
            _consent_report_error(
                on_error,
                u"Не удалось создать окно предупреждения о плагинах. Запуск отменён.",
            )
            return False
        try:
            from libre_macros_ui_theme import ORANGE_TITLE_BG, ORANGE_TITLE_FG, try_paint_titlebar

            try_paint_titlebar(dlg, title_bg=ORANGE_TITLE_BG, title_fg=ORANGE_TITLE_FG)
        except Exception:
            pass

        body_ctl = dlg.getControl("PluginWarnText")
        if body_ctl is not None:
            try:
                body_ctl.setText(body)
            except Exception:
                pass
            # После темы снова крупный шрифт (тема могла сбросить descriptor).
            try:
                from libre_macros_ui_theme import PLUGIN_WARN_TEXT_FONT_HEIGHT

                fh = int(PLUGIN_WARN_TEXT_FONT_HEIGHT)
                bm = body_ctl.Model
                fd = bm.FontDescriptor
                fd.Height = fh
                bm.FontDescriptor = fd
            except Exception:
                try:
                    body_ctl.Model.FontHeight = 12
                except Exception:
                    pass

        state = {"continue": False}

        class _Handler(unohelper.Base, XActionListener):
            def disposing(self, event):
                pass

            def actionPerformed(self, event):
                try:
                    name = event.Source.getModel().Name
                except Exception:
                    return
                if name == "ContinueBtn":
                    state["continue"] = True
                    dlg.endExecute()
                elif name == "CancelBtn":
                    state["continue"] = False
                    dlg.endExecute()

        handler = _Handler()
        for btn in ("ContinueBtn", "CancelBtn"):
            try:
                dlg.getControl(btn).addActionListener(handler)
            except Exception:
                pass

        dlg.execute()
        return bool(state["continue"])
    except Exception as err:
        print(u"consent: ошибка диалога плагинов:\n%s" % traceback.format_exc())
        _consent_report_error(
            on_error,
            u"Ошибка предупреждения о плагинах:\n%s\nЗапуск отменён." % err,
        )
        return False


def lm_consent_ensure_before_run( doc, macro_version=u"", plugin_refs=None, toolkit=None, parent_window=None, on_error=None):
    """
    Проверка условий и предупреждение о плагинах перед запуском сбора.
    Возвращает True, если можно продолжать.
    """
    lm_consent_ensure_terms_file()
    if not lm_consent_is_accepted():
        if not lm_consent_dialog_show_terms(
            doc,
            macro_version=macro_version,
            toolkit=toolkit,
            parent_window=parent_window,
            on_error=on_error,
        ):
            return False
    refs = plugin_refs or []
    if len(refs) > 0:
        if not lm_consent_dialog_show_plugin_warning(
            doc,
            refs,
            toolkit=toolkit,
            parent_window=parent_window,
            on_error=on_error,
        ):
            return False
    return True
