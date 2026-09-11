# -*- coding: utf-8 -*-
"""
Генерация карточек с QR-кодами для печати.

Лист «Генератор_QR» — данные; шаблон «QR_Print» — макет карточки.
"""
MACRO_VERSION = "3.10.725"
from libre_macros_qr_codes_lib import generate_qr_codes as _generate_qr_codes


def generate_qr_codes(*args):
    return _generate_qr_codes(*args)


g_exportedScripts = (generate_qr_codes,)
