# -*- coding: utf-8 -*-
from __future__ import print_function, unicode_literals
"""
Санация пользовательского Python перед eval/exec и загрузкой file#func.

Пока не подключена к collect_workbooks.py — только API для будущей интеграции.

Покрывает:
  • inline-код из колонки B / JSON field ``code`` (lambda / def);
  • исходник функции, вырезанный из файла по ссылке ``path.py#func_name``;
  • разрешение пути file# (каталог макроса, запрет ``..`` и произвольных abs);
  • запрет сети/сокетов, base64/шифроконтента, записи в ФС.

См. docs/16_CODE_SANITIZE.md.
"""
MACRO_VERSION = "3.10.688"
import ast
import os
import re

try:
    unicode
except NameError:
    unicode = str

# ---------------------------------------------------------------------------
# Политики и константы
# ---------------------------------------------------------------------------

# Уровни строгости (строка или константа).
LM_SANITIZE_POLICY_STRICT = "strict"
LM_SANITIZE_POLICY_STANDARD = "standard"
LM_SANITIZE_POLICY_PERMISSIVE = "permissive"

# --- песочница / introspection ---
_LM_SANITIZE_BANNED_CALLS_SANDBOX = frozenset(
    {
        "__import__",
        "eval",
        "exec",
        "execfile",
        "compile",
        "input",
        "raw_input",
        "breakpoint",
        "help",
        "exit",
        "quit",
        "globals",
        "locals",
        "vars",
        "dir",
        "getattr",
        "setattr",
        "delattr",
        "hasattr",
        "memoryview",
        "bytearray",
        "buffer",
        "file",
        "reload",
        "__builtins__",
    }
)

# --- сеть / сокеты любых протоколов ---
# Узнаваемые сетевые API + типичные методы сокета (в контексте pp ложные срабатывания редки).
_LM_SANITIZE_BANNED_CALLS_NETWORK = frozenset(
    {
        "socket",
        "socketpair",
        "create_connection",
        "create_server",
        "fromfd",
        "connect",
        "connect_ex",
        "listen",
        "bind",
        "accept",
        "send",
        "sendall",
        "sendto",
        "sendmsg",
        "recv",
        "recvfrom",
        "recv_into",
        "recvfrom_into",
        "recvmsg",
        "setsockopt",
        "getsockopt",
        "getsockname",
        "getpeername",
        "getaddrinfo",
        "gethostbyname",
        "gethostbyname_ex",
        "gethostbyaddr",
        "getnameinfo",
        "getprotobyname",
        "getservbyname",
        "getservbyport",
        "urlopen",
        "urlretrieve",
        "build_opener",
        "install_opener",
        "HTTPConnection",
        "HTTPSConnection",
        "HTTPHandler",
        "HTTPSHandler",
        "FTP",
        "FTP_TLS",
        "SMTP",
        "SMTP_SSL",
        "Telnet",
        "wrap_socket",
        "SSLContext",
        "create_default_context",
        "open_connection",
        "start_server",
        "create_datagram_endpoint",
        "open_unix_connection",
        "start_unix_server",
        "ServerProxy",
        "WebSocket",
        "WebSocketApp",
    }
)

# --- base64 / кодирование-шифрование полезной нагрузки ---
_LM_SANITIZE_BANNED_CALLS_CRYPTO = frozenset(
    {
        "b64encode",
        "b64decode",
        "standard_b64encode",
        "standard_b64decode",
        "urlsafe_b64encode",
        "urlsafe_b64decode",
        "b32encode",
        "b32decode",
        "b16encode",
        "b16decode",
        "b85encode",
        "b85decode",
        "a85encode",
        "a85decode",
        "encodebytes",
        "decodebytes",
        "encodestring",
        "decodestring",
        "b2a_base64",
        "a2b_base64",
        "b2a_hex",
        "a2b_hex",
        "hexlify",
        "unhexlify",
        "encrypt",
        "decrypt",
        "encryptor",
        "decryptor",
        "Fernet",
        "AES",
        "DES",
        "DES3",
        "Blowfish",
        "ChaCha20",
        "PKCS1_OAEP",
        "PKCS1_v1_5",
        "Cipher",
        "CipherContext",
    }
)

# --- запись в файловую систему (и UNO store*) ---
# Без общих имён copy/remove/replace (ложные срабатывания на list/str).
_LM_SANITIZE_BANNED_CALLS_FS_WRITE = frozenset(
    {
        "open",
        "writelines",
        "truncate",
        "unlink",
        "renames",
        "rmdir",
        "removedirs",
        "mkdir",
        "makedirs",
        "chmod",
        "chown",
        "symlink",
        "write_text",
        "write_bytes",
        "copy2",
        "copyfile",
        "copyfileobj",
        "copytree",
        "rmtree",
        "make_archive",
        "unpack_archive",
        "fdopen",
        "NamedTemporaryFile",
        "TemporaryFile",
        "SpooledTemporaryFile",
        "mkstemp",
        "mkdtemp",
        "storeAsURL",
        "storeToURL",
        "store",  # XStorable.store() — запись текущего файла
        "write",
    }
)

# Имена, запрещённые как вызовы (Call с Name.id / Attribute.attr).
_LM_SANITIZE_BANNED_CALL_NAMES = (
    _LM_SANITIZE_BANNED_CALLS_SANDBOX
    | _LM_SANITIZE_BANNED_CALLS_NETWORK
    | _LM_SANITIZE_BANNED_CALLS_CRYPTO
    | _LM_SANITIZE_BANNED_CALLS_FS_WRITE
)

# Идентификаторы модулей/API, запрещённые как Name (даже без вызова).
_LM_SANITIZE_BANNED_NAMES = frozenset(
    {
        "__builtins__",
        "__import__",
        "__loader__",
        "__spec__",
        # сеть
        "socket",
        "ssl",
        "select",
        "selectors",
        "asyncio",
        "urllib",
        "urllib2",
        "urllib3",
        "http",
        "httplib",
        "httpx",
        "requests",
        "aiohttp",
        "ftplib",
        "smtplib",
        "telnetlib",
        "xmlrpc",
        "xmlrpclib",
        "websocket",
        "websockets",
        "paramiko",
        "twisted",
        # base64 / crypto
        "base64",
        "binascii",
        "Crypto",
        "Cryptodome",
        "cryptography",
        "nacl",
        "PyNaCl",
        # ФС-запись
        "shutil",
        "tempfile",
        "pickle",
        "shelve",
        "marshal",
    }
)

# Мягкие вызовы: в permissive → warning (introspection).
_LM_SANITIZE_SOFT_CALLS = frozenset(
    {
        "getattr",
        "setattr",
        "delattr",
        "hasattr",
        "vars",
        "dir",
    }
)

# Атрибуты, доступ к которым запрещён (обход песочницы через object model).
_LM_SANITIZE_BANNED_ATTRS = frozenset(
    {
        "__class__",
        "__bases__",
        "__mro__",
        "__subclasses__",
        "__mro_entries__",
        "__globals__",
        "__code__",
        "__closure__",
        "__defaults__",
        "__kwdefaults__",
        "__dict__",
        "__getattribute__",
        "__setattr__",
        "__delattr__",
        "__reduce__",
        "__reduce_ex__",
        "__builtins__",
        "__import__",
        "__loader__",
        "__spec__",
        "__module__",
        "func_globals",
        "func_code",
        "func_closure",
        "gi_frame",
        "gi_code",
        "f_builtins",
        "f_globals",
        "f_locals",
        "f_code",
        "f_back",
        "tb_frame",
        "tb_next",
        "co_code",
        "co_consts",
    }
)

# Разрешённые «безопасные» dunder-атрибуты (имена/документация).
_LM_SANITIZE_ALLOWED_DUNDER_ATTRS = frozenset(
    {
        "__name__",
        "__doc__",
        "__qualname__",
    }
)

# Узлы AST, запрещённые целиком (импорты, exec-statement Py2 и т.п.).
_LM_SANITIZE_BANNED_NODE_TYPES = (
    ast.Import,
    ast.ImportFrom,
)

# Доп. типы, если есть в версии Python.
for _name in ("Exec", "Print"):  # Py2
    _cls = getattr(ast, _name, None)
    if _cls is not None:
        _LM_SANITIZE_BANNED_NODE_TYPES = _LM_SANITIZE_BANNED_NODE_TYPES + (_cls,)

# Builtins blocklist — совместим с _MERGE_PP_EVAL_BUILTIN_BLOCKLIST в макросе.
LM_SANITIZE_BUILTIN_BLOCKLIST = frozenset(
    {
        "__import__",
        "eval",
        "exec",
        "compile",
        "open",
        "input",
        "breakpoint",
        "help",
        "exit",
        "quit",
        "copyright",
        "credits",
        "license",
        "globals",
        "locals",
        "vars",
        "dir",
    }
)

# Идентификатор функции в file#func.
_LM_SANITIZE_FUNC_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
# Кириллические имена функций (как в картах постобработки) — допускаем в ref.
_LM_SANITIZE_FUNC_NAME_UNICODE_RE = re.compile(
    r"^[A-Za-zА-Яа-яЁё_][A-Za-zА-Яа-яЁё0-9_]*$", re.UNICODE
)

# Текстовые маркеры сети / crypto / ФС (casefold-поиск по исходнику).
_LM_SANITIZE_TEXT_MARKERS_NETWORK = (
    (u"socket.", "HEUR_NETWORK"),
    (u"ssl.", "HEUR_NETWORK"),
    (u"asyncio.", "HEUR_NETWORK"),
    (u"urllib", "HEUR_NETWORK"),
    (u"http.client", "HEUR_NETWORK"),
    (u"httplib", "HEUR_NETWORK"),
    (u"requests.", "HEUR_NETWORK"),
    (u"aiohttp", "HEUR_NETWORK"),
    (u"ftplib", "HEUR_NETWORK"),
    (u"smtplib", "HEUR_NETWORK"),
    (u"telnetlib", "HEUR_NETWORK"),
    (u"xmlrpc", "HEUR_NETWORK"),
    (u"websocket", "HEUR_NETWORK"),
    (u"create_connection", "HEUR_NETWORK"),
    (u"urlopen", "HEUR_NETWORK"),
    (u"af_inet", "HEUR_NETWORK"),
    (u"af_inet6", "HEUR_NETWORK"),
    (u"af_unix", "HEUR_NETWORK"),
    (u"sock_stream", "HEUR_NETWORK"),
    (u"sock_dgram", "HEUR_NETWORK"),
    (u"sock_raw", "HEUR_NETWORK"),
)

_LM_SANITIZE_TEXT_MARKERS_CRYPTO = (
    (u"base64", "HEUR_CRYPTO"),
    (u"b64encode", "HEUR_CRYPTO"),
    (u"b64decode", "HEUR_CRYPTO"),
    (u"binascii", "HEUR_CRYPTO"),
    (u"b2a_base64", "HEUR_CRYPTO"),
    (u"a2b_base64", "HEUR_CRYPTO"),
    (u"cryptography", "HEUR_CRYPTO"),
    (u"from crypto", "HEUR_CRYPTO"),
    (u"import crypto", "HEUR_CRYPTO"),
    (u"cryptodome", "HEUR_CRYPTO"),
    (u"fernet", "HEUR_CRYPTO"),
    (u"aes.new", "HEUR_CRYPTO"),
    (u"pkcs1", "HEUR_CRYPTO"),
)

_LM_SANITIZE_TEXT_MARKERS_FS_WRITE = (
    (u"os.remove", "HEUR_FS_WRITE"),
    (u"os.unlink", "HEUR_FS_WRITE"),
    (u"os.rename", "HEUR_FS_WRITE"),
    (u"os.mkdir", "HEUR_FS_WRITE"),
    (u"os.makedirs", "HEUR_FS_WRITE"),
    (u"os.rmdir", "HEUR_FS_WRITE"),
    (u"shutil.", "HEUR_FS_WRITE"),
    (u"tempfile", "HEUR_FS_WRITE"),
    (u"storeasurl", "HEUR_FS_WRITE"),
    (u"storetour", "HEUR_FS_WRITE"),
    (u"write_text", "HEUR_FS_WRITE"),
    (u"write_bytes", "HEUR_FS_WRITE"),
)

# Отдельные regex-маркеры ФС (чтобы не ловить urlopen / setString.write и т.п.).
_LM_SANITIZE_FS_OPEN_RE = re.compile(r"(?<![a-z0-9_])open\s*\(")
_LM_SANITIZE_FS_WRITE_RE = re.compile(r"\.write\s*\(")
_LM_SANITIZE_FS_WRITELINES_RE = re.compile(r"\.writelines\s*\(")

# Подозрительные строковые литералы: длинный base64 / hex blob.
_LM_SANITIZE_B64_LITERAL_RE = re.compile(r"^[A-Za-z0-9+/]{40,}={0,2}$")
_LM_SANITIZE_B64URL_LITERAL_RE = re.compile(r"^[A-Za-z0-9_-]{40,}={0,2}$")
_LM_SANITIZE_HEX_LITERAL_RE = re.compile(r"^[0-9a-fA-F]{48,}$")
_LM_SANITIZE_BLOB_MIN_LEN = 40


class LmSanitizeError(Exception):
    """Нарушение политики санации (код или путь)."""

    def __init__(self, message, code="SANITIZE", details=None):
        Exception.__init__(self, message)
        self.code = str(code or "SANITIZE")
        self.details = details if details is not None else []
        self.message = unicode(message)


class LmSanitizeResult(object):
    """
    Результат проверки исходника.

    ok          — True, если нарушений нет (или только warnings при allow_warnings).
    source      — исходный текст (как передан).
    issues      — список dict: {severity, code, message, lineno, col}.
    tree        — ast.AST или None при SyntaxError.
    policy      — имя политики.
    """

    def __init__(self, ok, source, issues=None, tree=None, policy=None):
        self.ok = bool(ok)
        self.source = source
        self.issues = list(issues or [])
        self.tree = tree
        self.policy = policy or LM_SANITIZE_POLICY_STANDARD

    @property
    def errors(self):
        return [i for i in self.issues if i.get("severity") == "error"]

    @property
    def warnings(self):
        return [i for i in self.issues if i.get("severity") == "warning"]

    def raise_if_failed(self):
        if self.ok:
            return self
        msgs = []
        for i in self.errors:
            loc = ""
            if i.get("lineno"):
                loc = " (стр. %s)" % i["lineno"]
            msgs.append("%s%s: %s" % (i.get("code", "?"), loc, i.get("message", "")))
        raise LmSanitizeError(
            "; ".join(msgs) if msgs else "sanitize failed",
            code="SANITIZE_CODE",
            details=self.issues,
        )


class LmSanitizeFileRefResult(object):
    """
    Результат разбора и проверки ссылки file#func.

    ok, ref, path_part, func_name, abs_path, issues.
    """

    def __init__( self, ok, ref, path_part=None, func_name=None, abs_path=None, issues=None):
        self.ok = bool(ok)
        self.ref = ref
        self.path_part = path_part
        self.func_name = func_name
        self.abs_path = abs_path
        self.issues = list(issues or [])

    def raise_if_failed(self):
        if self.ok:
            return self
        msgs = [i.get("message", "") for i in self.issues if i.get("severity") == "error"]
        raise LmSanitizeError(
            "; ".join(msgs) if msgs else "sanitize file_ref failed",
            code="SANITIZE_FILE_REF",
            details=self.issues,
        )


# ---------------------------------------------------------------------------
# Вспомогательные
# ---------------------------------------------------------------------------


def _lm_sanitize_u(s):
    if s is None:
        return u""
    if isinstance(s, unicode):
        return s
    try:
        return unicode(s)
    except Exception:
        return unicode(str(s))


def _lm_sanitize_issue(severity, code, message, lineno=None, col=None):
    return {
        "severity": severity,
        "code": code,
        "message": _lm_sanitize_u(message),
        "lineno": lineno,
        "col": col,
    }


def _lm_sanitize_normalize_policy(policy):
    p = _lm_sanitize_u(policy or LM_SANITIZE_POLICY_STANDARD).strip().casefold()
    if p in ("strict", "строгий"):
        return LM_SANITIZE_POLICY_STRICT
    if p in ("permissive", "мягкий", "слабый"):
        return LM_SANITIZE_POLICY_PERMISSIVE
    return LM_SANITIZE_POLICY_STANDARD


def _lm_sanitize_attr_name(node):
    """Имя атрибута из ast.Attribute (str / ast.Name в старых деревьях)."""
    attr = getattr(node, "attr", None)
    if isinstance(attr, (str, unicode)):
        return attr
    return None


def _lm_sanitize_call_name(node):
    """Простое имя вызываемой функции: foo(...) или obj.bar(...)-> bar."""
    func = getattr(node, "func", None)
    if func is None:
        return None
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute):
        return _lm_sanitize_attr_name(func)
    return None


def _lm_sanitize_is_dunder(name):
    n = _lm_sanitize_u(name)
    return len(n) >= 4 and n.startswith("__") and n.endswith("__")


def _lm_sanitize_call_issue_code(name):
    """Код issue по категории запрещённого вызова/атрибута."""
    if name in _LM_SANITIZE_BANNED_CALLS_NETWORK:
        return "NETWORK"
    if name in _LM_SANITIZE_BANNED_CALLS_CRYPTO:
        return "CRYPTO"
    if name in _LM_SANITIZE_BANNED_CALLS_FS_WRITE:
        return "FS_WRITE"
    return "BANNED_CALL"


def _lm_sanitize_name_issue_code(name):
    """Код issue для запрещённого идентификатора Name."""
    if name in (
        "socket",
        "ssl",
        "select",
        "selectors",
        "asyncio",
        "urllib",
        "urllib2",
        "urllib3",
        "http",
        "httplib",
        "httpx",
        "requests",
        "aiohttp",
        "ftplib",
        "smtplib",
        "telnetlib",
        "xmlrpc",
        "xmlrpclib",
        "websocket",
        "websockets",
        "paramiko",
        "twisted",
    ):
        return "NETWORK"
    if name in (
        "base64",
        "binascii",
        "Crypto",
        "Cryptodome",
        "cryptography",
        "nacl",
        "PyNaCl",
    ):
        return "CRYPTO"
    if name in ("shutil", "tempfile", "pickle", "shelve", "marshal"):
        return "FS_WRITE"
    return "BANNED_NAME"


def _lm_sanitize_text_fold(s):
    t = _lm_sanitize_u(s)
    try:
        return t.casefold()
    except Exception:
        return t.lower()


def _lm_sanitize_scan_text_markers(text, policy):
    """Эвристики по тексту исходника (сеть / crypto / ФС)."""
    issues = []
    folded = _lm_sanitize_text_fold(text)
    seen_codes = set()
    groups = (
        _LM_SANITIZE_TEXT_MARKERS_NETWORK,
        _LM_SANITIZE_TEXT_MARKERS_CRYPTO,
        _LM_SANITIZE_TEXT_MARKERS_FS_WRITE,
    )
    for group in groups:
        for needle, code in group:
            if needle in folded:
                if code in seen_codes:
                    continue
                seen_codes.add(code)
                sev = "error"
                if policy == LM_SANITIZE_POLICY_PERMISSIVE:
                    sev = "warning"
                issues.append(
                    _lm_sanitize_issue(
                        sev,
                        code,
                        "В исходнике встречается маркер «%s»" % needle,
                    )
                )
    # Regex-маркеры записи в ФС (не substring open( внутри urlopen).
    if "HEUR_FS_WRITE" not in seen_codes:
        fs_hit = None
        if _LM_SANITIZE_FS_OPEN_RE.search(folded):
            fs_hit = u"open("
        elif _LM_SANITIZE_FS_WRITE_RE.search(folded):
            fs_hit = u".write("
        elif _LM_SANITIZE_FS_WRITELINES_RE.search(folded):
            fs_hit = u".writelines("
        if fs_hit is not None:
            sev = "warning" if policy == LM_SANITIZE_POLICY_PERMISSIVE else "error"
            issues.append(
                _lm_sanitize_issue(
                    sev,
                    "HEUR_FS_WRITE",
                    "В исходнике встречается маркер «%s»" % fs_hit,
                )
            )
    return issues


# ---------------------------------------------------------------------------
# AST visitor
# ---------------------------------------------------------------------------


class _LmSanitizeVisitor(ast.NodeVisitor):
    """Обход AST: собирает issues по политике."""

    def __init__(self, policy, extra_banned_calls=None, extra_allowed_attrs=None):
        self.policy = policy
        self.issues = []
        self.extra_banned_calls = frozenset(extra_banned_calls or ())
        self.extra_allowed_attrs = frozenset(extra_allowed_attrs or ())

    def _err(self, code, message, node=None):
        lineno = getattr(node, "lineno", None) if node is not None else None
        col = getattr(node, "col_offset", None) if node is not None else None
        self.issues.append(
            _lm_sanitize_issue("error", code, message, lineno=lineno, col=col)
        )

    def _warn(self, code, message, node=None):
        lineno = getattr(node, "lineno", None) if node is not None else None
        col = getattr(node, "col_offset", None) if node is not None else None
        self.issues.append(
            _lm_sanitize_issue("warning", code, message, lineno=lineno, col=col)
        )

    def generic_visit(self, node):
        if isinstance(node, _LM_SANITIZE_BANNED_NODE_TYPES):
            self._err(
                "BANNED_NODE",
                "Запрещённый узел AST: %s" % type(node).__name__,
                node,
            )
            return
        ast.NodeVisitor.generic_visit(self, node)

    def visit_Import(self, node):
        self._err("IMPORT", "import запрещён в пользовательском коде", node)

    def visit_ImportFrom(self, node):
        self._err("IMPORT", "from … import запрещён в пользовательском коде", node)

    def visit_Call(self, node):
        name = _lm_sanitize_call_name(node)
        if name and (
            name in _LM_SANITIZE_BANNED_CALL_NAMES or name in self.extra_banned_calls
        ):
            soft = name in _LM_SANITIZE_SOFT_CALLS
            code = _lm_sanitize_call_issue_code(name)
            if self.policy == LM_SANITIZE_POLICY_PERMISSIVE and soft:
                self._warn(
                    code,
                    "Вызов %s() нежелателен политикой санации" % name,
                    node,
                )
            else:
                self._err(
                    code,
                    "Вызов %s() запрещён политикой санации" % name,
                    node,
                )
        self.generic_visit(node)

    def visit_Attribute(self, node):
        attr = _lm_sanitize_attr_name(node)
        if attr:
            allowed = _LM_SANITIZE_ALLOWED_DUNDER_ATTRS | self.extra_allowed_attrs
            if attr in allowed:
                pass
            elif attr in _LM_SANITIZE_BANNED_ATTRS:
                self._err(
                    "BANNED_ATTR",
                    "Доступ к атрибуту .%s запрещён" % attr,
                    node,
                )
            elif attr in _LM_SANITIZE_BANNED_CALL_NAMES:
                code = _lm_sanitize_call_issue_code(attr)
                self._err(
                    code,
                    "Атрибут/API .%s запрещён политикой санации" % attr,
                    node,
                )
            elif _lm_sanitize_is_dunder(attr):
                if self.policy == LM_SANITIZE_POLICY_PERMISSIVE:
                    self._warn(
                        "DUNDER_ATTR",
                        "Dunder-атрибут .%s подозрителен" % attr,
                        node,
                    )
                else:
                    self._err(
                        "DUNDER_ATTR",
                        "Dunder-атрибут .%s запрещён" % attr,
                        node,
                    )
        self.generic_visit(node)

    def visit_Name(self, node):
        nid = getattr(node, "id", None)
        if nid and nid in _LM_SANITIZE_BANNED_NAMES:
            code = _lm_sanitize_name_issue_code(nid)
            self._err(
                code,
                "Имя %s запрещено политикой санации" % nid,
                node,
            )
        self.generic_visit(node)

    def visit_Constant(self, node):
        self._check_string_literal(getattr(node, "value", None), node)
        self.generic_visit(node)

    def visit_Str(self, node):
        self._check_string_literal(getattr(node, "s", None), node)
        self.generic_visit(node)

    def visit_Bytes(self, node):
        raw = getattr(node, "s", None)
        if raw is not None:
            try:
                text = raw.decode("ascii", "ignore")
            except Exception:
                text = ""
            self._check_string_literal(text, node)
        self.generic_visit(node)

    def _check_string_literal(self, value, node):
        if not isinstance(value, (str, unicode, bytes)):
            return
        if isinstance(value, bytes):
            try:
                s = value.decode("ascii", "ignore")
            except Exception:
                return
        else:
            s = _lm_sanitize_u(value).strip()
        if len(s) < _LM_SANITIZE_BLOB_MIN_LEN:
            return
        compact = u"".join(s.split())
        if len(compact) < _LM_SANITIZE_BLOB_MIN_LEN:
            return
        if _LM_SANITIZE_B64_LITERAL_RE.match(compact) or _LM_SANITIZE_B64URL_LITERAL_RE.match(
            compact
        ):
            self._err(
                "ENCODED_BLOB",
                "Строковый литерал похож на base64/шифроконтент (длина %s)"
                % len(compact),
                node,
            )
            return
        if _LM_SANITIZE_HEX_LITERAL_RE.match(compact) and (len(compact) % 2 == 0):
            self._err(
                "ENCODED_BLOB",
                "Строковый литерал похож на hex-шифроконтент (длина %s)"
                % len(compact),
                node,
            )

    def visit_FunctionDef(self, node):
        self.generic_visit(node)

    def visit_Lambda(self, node):
        self.generic_visit(node)


# ---------------------------------------------------------------------------
# Публичный API: код
# ---------------------------------------------------------------------------


def lm_sanitize_code( source, policy=LM_SANITIZE_POLICY_STANDARD, extra_banned_calls=None, extra_allowed_attrs=None, allow_warnings=True):
    """
    Статическая проверка исходника пользовательского Python (AST).

    Не выполняет код. Возвращает LmSanitizeResult.

    policy:
      strict     — максимум запретов (в т.ч. почти все dunder-атрибуты);
      standard   — рекомендуемый режим для B/C и file# (по умолчанию);
      permissive — опасные вызовы getattr и т.п. → warning, не error.

    allow_warnings=True: ok=True при отсутствии error (warnings допустимы).
    """
    text = _lm_sanitize_u(source)
    pol = _lm_sanitize_normalize_policy(policy)
    issues = []
    tree = None
    if text.strip() == u"":
        issues.append(
            _lm_sanitize_issue("error", "EMPTY", "Пустой исходный код")
        )
        return LmSanitizeResult(False, text, issues=issues, tree=None, policy=pol)
    try:
        tree = ast.parse(text)
    except SyntaxError as err:
        issues.append(
            _lm_sanitize_issue(
                "error",
                "SYNTAX",
                "Синтаксическая ошибка: %s" % err,
                lineno=getattr(err, "lineno", None),
                col=getattr(err, "offset", None),
            )
        )
        return LmSanitizeResult(False, text, issues=issues, tree=None, policy=pol)

    visitor = _LmSanitizeVisitor(
        pol,
        extra_banned_calls=extra_banned_calls,
        extra_allowed_attrs=extra_allowed_attrs,
    )
    visitor.visit(tree)
    issues.extend(visitor.issues)

    # Эвристика по тексту: динамическая сборка имён может обойти AST Attribute.
    if pol != LM_SANITIZE_POLICY_PERMISSIVE:
        for needle, code in (
            (u"__subclasses__", "HEUR_SUBCLASSES"),
            (u"__globals__", "HEUR_GLOBALS"),
            (u"__builtins__", "HEUR_BUILTINS"),
        ):
            if needle not in text:
                continue
            already = any(i.get("code") == code for i in issues)
            if already:
                continue
            attr_hit = any(
                i.get("code") in ("BANNED_ATTR", "DUNDER_ATTR")
                and needle in i.get("message", u"")
                for i in issues
            )
            if attr_hit:
                continue
            issues.append(
                _lm_sanitize_issue(
                    "warning" if pol == LM_SANITIZE_POLICY_STANDARD else "error",
                    code,
                    "В исходнике встречается %s" % needle,
                )
            )

    # Сеть / crypto / запись ФС — текстовые маркеры (дополняют AST).
    # Не дублируем, если AST уже выставил error той же категории.
    marker_issues = _lm_sanitize_scan_text_markers(text, pol)
    existing_cats = set()
    for i in issues:
        c = i.get("code") or ""
        if c in (
            "NETWORK",
            "CRYPTO",
            "FS_WRITE",
            "ENCODED_BLOB",
            "HEUR_NETWORK",
            "HEUR_CRYPTO",
            "HEUR_FS_WRITE",
        ):
            existing_cats.add(c.replace("HEUR_", ""))
    for mi in marker_issues:
        cat = (mi.get("code") or "").replace("HEUR_", "")
        # Если уже есть AST-error той же категории — текстовый маркер не добавляем.
        if cat in existing_cats and any(
            (x.get("code") or "") in (cat, "ENCODED_BLOB")
            and x.get("severity") == "error"
            for x in issues
        ):
            continue
        issues.append(mi)

    errors = [i for i in issues if i.get("severity") == "error"]
    ok = len(errors) == 0 if allow_warnings else len(issues) == 0
    return LmSanitizeResult(ok, text, issues=issues, tree=tree, policy=pol)


def lm_sanitize_code_or_raise(source, policy=LM_SANITIZE_POLICY_STANDARD, **kwargs):
    """Как lm_sanitize_code, но при ошибках бросает LmSanitizeError."""
    return lm_sanitize_code(source, policy=policy, **kwargs).raise_if_failed()


# ---------------------------------------------------------------------------
# Публичный API: file#func
# ---------------------------------------------------------------------------


def lm_sanitize_parse_file_ref(ref):
    """
    Разобрать строку «path.py#func_name».

    Возвращает (path_part, func_name) или (None, None).
    """
    t = _lm_sanitize_u(ref).strip()
    if t == u"" or u"\n" in t:
        return None, None
    if t.count(u"#") != 1:
        return None, None
    if t.startswith(u"def ") or t.startswith(u"lambda"):
        return None, None
    path_part, fn_name = t.split(u"#", 1)
    path_part = path_part.strip()
    fn_name = fn_name.strip()
    if path_part == u"" or fn_name == u"":
        return None, None
    return path_part, fn_name


def lm_sanitize_is_safe_func_name(name):
    """True, если имя функции допустимо в file#ref (ASCII или кириллица)."""
    n = _lm_sanitize_u(name).strip()
    if n == u"":
        return False
    if _LM_SANITIZE_FUNC_NAME_RE.match(n):
        return True
    if _LM_SANITIZE_FUNC_NAME_UNICODE_RE.match(n):
        return True
    # Python 3 isidentifier
    try:
        if n.isidentifier():
            return True
    except Exception:
        pass
    return False


def lm_sanitize_resolve_script_path( path_part, base_dir, allow_absolute=False, allowed_roots=None):
    """
    Разрешить путь к .py относительно base_dir.

    По умолчанию (allow_absolute=False):
      • абсолютные пути отклоняются;
      • ``..`` и выход за base_dir отклоняются;
      • результат — нормализованный abs путь внутри base_dir.

    allow_absolute=True: abs допускается, но если задан allowed_roots —
    путь должен лежать под одним из корней.

    Возвращает abs_path (str) или None.
    """
    p = _lm_sanitize_u(path_part).strip()
    if p == u"":
        return None
    base = os.path.abspath(os.path.normpath(_lm_sanitize_u(base_dir)))
    if not base or not os.path.isdir(base):
        return None

    if os.path.isabs(p):
        if not allow_absolute:
            return None
        cand = os.path.abspath(os.path.normpath(p))
        if allowed_roots:
            if not lm_sanitize_path_under_roots(cand, allowed_roots):
                return None
        return cand if os.path.isfile(cand) else None

    # Относительный: запрет «сырого» .. в сегментах до join — normpath всё равно
    # проверим, что результат под base.
    if u"\0" in p:
        return None
    cand = os.path.abspath(os.path.normpath(os.path.join(base, p)))
    if not lm_sanitize_path_under_roots(cand, [base]):
        return None
    if not os.path.isfile(cand):
        return None
    return cand


def lm_sanitize_path_under_roots(abs_path, roots):
    """True, если abs_path лежит внутри хотя бы одного root (после normpath)."""
    try:
        real = os.path.realpath(abs_path)
    except OSError:
        real = os.path.abspath(os.path.normpath(abs_path))
    for root in roots or ():
        try:
            root_real = os.path.realpath(root)
        except OSError:
            root_real = os.path.abspath(os.path.normpath(root))
        if real == root_real:
            return True
        prefix = root_real.rstrip(os.sep) + os.sep
        if real.startswith(prefix):
            return True
    return False


def lm_sanitize_file_ref( ref, base_dir, allow_absolute=False, allowed_roots=None, require_py_suffix=True):
    """
    Полная проверка ссылки ``file.py#func``: разбор, имя, путь.

    Не читает файл и не компилирует код.
    """
    raw = _lm_sanitize_u(ref).strip()
    issues = []
    path_part, func_name = lm_sanitize_parse_file_ref(raw)
    if path_part is None:
        issues.append(
            _lm_sanitize_issue(
                "error",
                "BAD_REF",
                "Ожидается ссылка вида path.py#func_name, получено: %r"
                % (raw[:80],),
            )
        )
        return LmSanitizeFileRefResult(False, raw, issues=issues)

    if not lm_sanitize_is_safe_func_name(func_name):
        issues.append(
            _lm_sanitize_issue(
                "error",
                "BAD_FUNC_NAME",
                "Недопустимое имя функции: %r" % (func_name,),
            )
        )

    if require_py_suffix:
        low = path_part.casefold() if hasattr(path_part, "casefold") else path_part.lower()
        if not low.endswith(u".py"):
            issues.append(
                _lm_sanitize_issue(
                    "error",
                    "NOT_PY",
                    "Ссылка file# должна указывать на .py, получено: %r" % (path_part,),
                )
            )

    if os.path.isabs(path_part) and not allow_absolute:
        issues.append(
            _lm_sanitize_issue(
                "error",
                "ABS_PATH",
                "Абсолютный путь к скрипту запрещён: %r" % (path_part,),
            )
        )

    abs_path = None
    if not any(i.get("severity") == "error" for i in issues):
        abs_path = lm_sanitize_resolve_script_path(
            path_part,
            base_dir,
            allow_absolute=allow_absolute,
            allowed_roots=allowed_roots,
        )
        if abs_path is None:
            issues.append(
                _lm_sanitize_issue(
                    "error",
                    "PATH_DENIED",
                    "Путь вне каталога макроса или файл не найден: %r (base=%r)"
                    % (path_part, base_dir),
                )
            )

    ok = not any(i.get("severity") == "error" for i in issues)
    return LmSanitizeFileRefResult(
        ok,
        raw,
        path_part=path_part,
        func_name=func_name,
        abs_path=abs_path,
        issues=issues,
    )


def lm_sanitize_file_ref_or_raise(ref, base_dir, **kwargs):
    return lm_sanitize_file_ref(ref, base_dir, **kwargs).raise_if_failed()


# ---------------------------------------------------------------------------
# Извлечение def + санация содержимого файла
# ---------------------------------------------------------------------------


def lm_sanitize_extract_function_source(file_text, func_name):
    """
    Вырезать исходник ``def func_name`` верхнего уровня (как в макросе).

    Сначала AST, иначе построчный fallback. Возвращает str или None.
    """
    text = _lm_sanitize_u(file_text)
    name = _lm_sanitize_u(func_name).strip()
    if text == u"" or name == u"":
        return None
    try:
        tree = ast.parse(text)
    except SyntaxError:
        return _lm_sanitize_extract_function_source_lines(text, name)
    lines = text.splitlines()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name == name:
            if hasattr(ast, "get_source_segment"):
                seg = ast.get_source_segment(text, node)
                if seg is not None and _lm_sanitize_u(seg).strip() != u"":
                    return seg
            end = getattr(node, "end_lineno", None)
            if end is not None:
                return u"\n".join(lines[node.lineno - 1 : end])
            return _lm_sanitize_extract_function_source_lines(text, name)
    return _lm_sanitize_extract_function_source_lines(text, name)


def _lm_sanitize_extract_function_source_lines(file_text, func_name):
    lines = file_text.splitlines()
    def_prefix = u"def " + func_name + u"("
    start = None
    i = 0
    while i < len(lines):
        line = lines[i]
        if start is None:
            if line.lstrip().startswith(def_prefix):
                start = i
        else:
            stripped = line.lstrip()
            if stripped.startswith(u"def ") or stripped.startswith(u"class "):
                if line and line[0] not in (u" ", u"\t"):
                    break
        i = i + 1
    if start is None:
        return None
    return u"\n".join(lines[start:i])


def lm_sanitize_load_file_function( ref, base_dir, policy=LM_SANITIZE_POLICY_STANDARD, allow_absolute=False, allowed_roots=None, file_reader=None):
    """
    Разрешить file#ref → прочитать файл → вырезать def → прогнать lm_sanitize_code.

    Не компилирует и не вызывает eval/exec.

    Возвращает dict:
      {
        "ok": bool,
        "ref_result": LmSanitizeFileRefResult,
        "source": str|None,
        "code_result": LmSanitizeResult|None,
        "issues": list,
      }

    file_reader(abs_path) -> str  — опционально (по умолчанию utf-8 open).
    """
    ref_res = lm_sanitize_file_ref(
        ref,
        base_dir,
        allow_absolute=allow_absolute,
        allowed_roots=allowed_roots,
    )
    out_issues = list(ref_res.issues)
    if not ref_res.ok:
        return {
            "ok": False,
            "ref_result": ref_res,
            "source": None,
            "code_result": None,
            "issues": out_issues,
        }

    reader = file_reader
    if reader is None:

        def _default_reader(path):
            import io

            f = io.open(path, "r", encoding="utf-8")
            try:
                return f.read()
            finally:
                f.close()

        reader = _default_reader

    try:
        file_text = reader(ref_res.abs_path)
    except Exception as err:
        out_issues.append(
            _lm_sanitize_issue(
                "error",
                "READ",
                "Не удалось прочитать %s: %s" % (ref_res.abs_path, err),
            )
        )
        return {
            "ok": False,
            "ref_result": ref_res,
            "source": None,
            "code_result": None,
            "issues": out_issues,
        }

    source = lm_sanitize_extract_function_source(file_text, ref_res.func_name)
    if source is None or _lm_sanitize_u(source).strip() == u"":
        out_issues.append(
            _lm_sanitize_issue(
                "error",
                "NO_FUNC",
                "В файле %s нет функции def %s"
                % (ref_res.abs_path, ref_res.func_name),
            )
        )
        return {
            "ok": False,
            "ref_result": ref_res,
            "source": None,
            "code_result": None,
            "issues": out_issues,
        }

    code_res = lm_sanitize_code(source, policy=policy)
    out_issues.extend(code_res.issues)
    ok = ref_res.ok and code_res.ok
    return {
        "ok": ok,
        "ref_result": ref_res,
        "source": source,
        "code_result": code_res,
        "issues": out_issues,
    }


# ---------------------------------------------------------------------------
# Builtins helper (для будущей интеграции с eval namespace)
# ---------------------------------------------------------------------------


def lm_sanitize_filtered_builtins(blocklist=None, source_module=None):
    """
    Словарь builtins без имён из blocklist (и без имён на «_»).

    Аналог _merge_pp_postprocess_eval_builtins(), но без кэша макроса.
    """
    bl = frozenset(blocklist) if blocklist is not None else LM_SANITIZE_BUILTIN_BLOCKLIST
    if source_module is None:
        try:
            import builtins as source_module
        except ImportError:
            import __builtin__ as source_module  # Py2
    out = {}
    for name, obj in vars(source_module).items():
        if name.startswith("_"):
            continue
        if name in bl:
            continue
        out[name] = obj
    return out


def lm_sanitize_format_issues(issues, max_items=20):
    """Краткая строка для лога макроса."""
    parts = []
    n = 0
    for i in issues or []:
        if n >= max_items:
            parts.append(u"…")
            break
        loc = u""
        if i.get("lineno"):
            loc = u":%s" % i["lineno"]
        parts.append(
            u"[%s]%s %s — %s"
            % (
                i.get("severity", u"?"),
                loc,
                i.get("code", u"?"),
                i.get("message", u""),
            )
        )
        n += 1
    return u"; ".join(parts)
