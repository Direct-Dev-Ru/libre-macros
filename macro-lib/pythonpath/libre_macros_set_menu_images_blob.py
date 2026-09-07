# -*- coding: utf-8 -*-
"""Вшитые lc_imagelist.xml + lc_userimages.png для set_menu.
Пересборка: python3 macro-lib/bundle_set_menu_images.py
"""
from __future__ import print_function, unicode_literals
MACRO_VERSION = "3.10.690"
import base64
import os

LC_IMAGELIST_FILENAME = "lc_imagelist.xml"
LC_USERIMAGES_PNG_FILENAME = "lc_userimages.png"
LC_USERIMAGES_PNG_NAME = LC_USERIMAGES_PNG_FILENAME

LC_IMAGELIST_XML = '''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE image:imagecontainer PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "image.dtd">
<image:imagescontainer xmlns:image="http://openoffice.org/2001/image" xmlns:xlink="http://www.w3.org/1999/xlink">
 <image:images xlink:type="simple">
  <image:entry image:command="vnd.sun.star.script:collect_workbooks.py$collect_pack?language=Python&amp;location=user"/>
  <image:entry image:command="vnd.sun.star.script:param_wizard.py$set_merge_param?language=Python&amp;location=user"/>
  <image:entry image:command="vnd.sun.star.script:param_wizard.py$merge_param_go_up?language=Python&amp;location=user"/>
  <image:entry image:command="vnd.sun.star.script:param_wizard.py$merge_param_go_down?language=Python&amp;location=user"/>
  <image:entry image:command="vnd.sun.star.script:param_wizard.py$toggle_merge_param_disabled?language=Python&amp;location=user"/>
  <image:entry image:command="vnd.sun.star.script:pick_source_file.py$select_source_file?language=Python&amp;location=user"/>
  <image:entry image:command="vnd.sun.star.script:copy_param_rows.py$copy_param_rows?language=Python&amp;location=user"/>
 </image:images>
</image:imagescontainer>
'''

# binary length 2919
LC_USERIMAGES_PNG_B64 = (
"""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
AAAAAElFTkSuQmCC
"""
)


def lc_imagelist_xml_text():
    return LC_IMAGELIST_XML


def lc_userimages_png_bytes():
    return base64.b64decode(LC_USERIMAGES_PNG_B64.encode("ascii"))


def write_lc_imagelist_xml(path):
    text = lc_imagelist_xml_text()
    path = os.path.abspath(str(path))
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(text)
        if not text.endswith("\n"):
            f.write("\n")
    return path


def write_lc_userimages_png(path):
    data = lc_userimages_png_bytes()
    path = os.path.abspath(str(path))
    parent = os.path.dirname(path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    with open(path, "wb") as f:
        f.write(data)
    return path
