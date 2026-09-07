# -*- coding: utf-8 -*-
"""
Защита ODS от изменений (LibreOffice / AlterOffice).

В settings.xml:
  - LoadReadonly = true  (открытие только для чтения)
  - ModifyPasswordInfo   (пароль на снятие защиты / запись)

Патч текстовый: ElementTree ломает namespace/структуру, и LO игнорирует настройки.
"""
import base64
import hashlib
import os
import re
import zipfile

DEFAULT_INSTALLER_PASSWORD = "macros_libre_123456"
PASSWORD_ENV_VAR = "MACRO_INSTALLER_PASSWORD"
PASSWORD_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "installer_password.txt")

PBKDF2_ITERATIONS = 100000
PBKDF2_HASH_LEN = 16

_MODIFY_PASSWORD_INFO_RE = re.compile(
    r"<config:config-item-set config:name=\"ModifyPasswordInfo\">.*?</config:config-item-set>\s*",
    re.DOTALL,
)
_LOAD_READONLY_RE = re.compile(
    r'(<config:config-item config:name="LoadReadonly" config:type="boolean">)'
    r"(?:false|true)"
    r"(</config:config-item>)",
)


def lo_modify_password_digest(password, iteration_count=PBKDF2_ITERATIONS):
    """Соль и hash (base64) для ModifyPasswordInfo."""
    salt = os.urandom(16)
    hash_bytes = hashlib.pbkdf2_hmac(
        "sha1",
        password.encode("utf-8"),
        salt,
        iteration_count,
        dklen=PBKDF2_HASH_LEN,
    )
    return (
        base64.b64encode(salt).decode("ascii"),
        iteration_count,
        base64.b64encode(hash_bytes).decode("ascii"),
    )


def _build_modify_password_info_xml(password):
    salt_b64, iters, hash_b64 = lo_modify_password_digest(password)
    return (
        '<config:config-item-set config:name="ModifyPasswordInfo">'
        '<config:config-item config:name="algorithm-name" config:type="string">PBKDF2</config:config-item>'
        f'<config:config-item config:name="salt" config:type="base64Binary">{salt_b64}</config:config-item>'
        f'<config:config-item config:name="iteration-count" config:type="int">{iters}</config:config-item>'
        f'<config:config-item config:name="hash" config:type="base64Binary">{hash_b64}</config:config-item>'
        "</config:config-item-set>"
    )


def patch_settings_xml_for_modify_password(settings_data, password):
    """Патч settings.xml без пересборки дерева XML."""
    if isinstance(settings_data, bytes):
        text = settings_data.decode("utf-8")
        encoded = True
    else:
        text = settings_data
        encoded = False

    if 'config:name="ooo:configuration-settings"' not in text:
        raise RuntimeError("В settings.xml не найден ooo:configuration-settings")

    text = _MODIFY_PASSWORD_INFO_RE.sub("", text)
    text, n = _LOAD_READONLY_RE.subn(r"\1true\2", text, count=1)
    if n == 0:
        insert_at = text.rfind("</config:config-item-set></office:settings>")
        if insert_at == -1:
            raise RuntimeError("В settings.xml не найден LoadReadonly и точка вставки")
        load_readonly_item = (
            '<config:config-item config:name="LoadReadonly" config:type="boolean">true</config:config-item>'
        )
        text = text[:insert_at] + load_readonly_item + text[insert_at:]

    modify_xml = _build_modify_password_info_xml(password)
    anchor = "</config:config-item-set></office:settings>"
    insert_at = text.rfind(anchor)
    if insert_at == -1:
        raise RuntimeError("В settings.xml не найден конец office:settings")
    text = text[:insert_at] + modify_xml + text[insert_at:]

    return text.encode("utf-8") if encoded else text


def protect_ods_modify_password(ods_path, password):
    """Записать в ODS защиту от изменений с паролем password."""
    ods_path = os.path.abspath(ods_path)
    with zipfile.ZipFile(ods_path, "r") as zin:
        infos = zin.infolist()
        payload = {info.filename: zin.read(info.filename) for info in infos}

    if "settings.xml" not in payload:
        raise RuntimeError(f"В {ods_path} нет settings.xml")

    payload["settings.xml"] = patch_settings_xml_for_modify_password(
        payload["settings.xml"], password
    )

    tmp_path = ods_path + ".tmp"
    with zipfile.ZipFile(tmp_path, "w") as zout:
        for info in infos:
            data = payload[info.filename]
            if info.filename == "mimetype":
                zout.writestr(info, data, compress_type=zipfile.ZIP_STORED)
            else:
                zout.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)
    os.replace(tmp_path, ods_path)


def read_password_file(path=PASSWORD_FILE):
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as f:
        for line in f:
            text = line.strip()
            if text and not text.startswith("#"):
                return text
    return None


def write_password_file(password, path=PASSWORD_FILE):
    with open(path, "w", encoding="utf-8") as f:
        f.write(
            "# Пароль защиты MacroInstaller_bundled.ods от изменений (LibreOffice).\n"
            "# Переопределение: переменная среды "
            + PASSWORD_ENV_VAR
            + "\n"
        )
        f.write(password + "\n")


def resolve_installer_password(password_arg=None, password_file=PASSWORD_FILE):
    """
    Пароль: аргумент CLI → env → файл → значение по умолчанию.
    Итоговый пароль записывается в installer_password.txt.
    """
    if password_arg is not None and str(password_arg).strip():
        password = str(password_arg).strip()
    else:
        env_val = os.environ.get(PASSWORD_ENV_VAR, "").strip()
        if env_val:
            password = env_val
        else:
            file_val = read_password_file(password_file)
            password = file_val if file_val else DEFAULT_INSTALLER_PASSWORD
    write_password_file(password, password_file)
    return password
