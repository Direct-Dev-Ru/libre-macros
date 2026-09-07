# -*- coding: utf-8 -*-
"""
odfpy (import odf) + defusedxml в одном файле для LibreOffice pythonpath.

Версии: defusedxml 0.7.1, odf 1.4.1
Сборка: python macro-lib/bundle_odfpy.py

Перед import целевых пакетов выполните:
  import odf_bundled
"""
from __future__ import annotations
MACRO_VERSION = "3.10.688"
import importlib.abc
import importlib.util
import sys

_FINDER = None
_BUNDLE_TAG = "odf_bundled"

_PREFIXES = ('defusedxml', 'odf')
_SOURCES = {
    'defusedxml': {
        'is_package': True,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defuse XML bomb denial of service vulnerabilities
"""
from __future__ import print_function, absolute_import

import warnings

from .common import (
    DefusedXmlException,
    DTDForbidden,
    EntitiesForbidden,
    ExternalReferenceForbidden,
    NotSupportedError,
    _apply_defusing,
)


def defuse_stdlib():
    """Monkey patch and defuse all stdlib packages

    :warning: The monkey patch is an EXPERIMETNAL feature.
    """
    defused = {}

    with warnings.catch_warnings():
        from . import cElementTree
    from . import ElementTree
    from . import minidom
    from . import pulldom
    from . import sax
    from . import expatbuilder
    from . import expatreader
    from . import xmlrpc

    xmlrpc.monkey_patch()
    defused[xmlrpc] = None

    defused_mods = [
        cElementTree,
        ElementTree,
        minidom,
        pulldom,
        sax,
        expatbuilder,
        expatreader,
    ]

    for defused_mod in defused_mods:
        stdlib_mod = _apply_defusing(defused_mod)
        defused[defused_mod] = stdlib_mod

    return defused


__version__ = "0.7.1"

__all__ = [
    "DefusedXmlException",
    "DTDForbidden",
    "EntitiesForbidden",
    "ExternalReferenceForbidden",
    "NotSupportedError",
]

''',
    },
    'defusedxml.ElementTree': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.etree.ElementTree facade
"""
from __future__ import print_function, absolute_import

import sys
import warnings
from xml.etree.ElementTree import ParseError
from xml.etree.ElementTree import TreeBuilder as _TreeBuilder
from xml.etree.ElementTree import parse as _parse
from xml.etree.ElementTree import tostring

from .common import PY3

if PY3:
    import importlib
else:
    from xml.etree.ElementTree import XMLParser as _XMLParser
    from xml.etree.ElementTree import iterparse as _iterparse


from .common import (
    DTDForbidden,
    EntitiesForbidden,
    ExternalReferenceForbidden,
    _generate_etree_functions,
)

__origin__ = "xml.etree.ElementTree"


def _get_py3_cls():
    """Python 3.3 hides the pure Python code but defusedxml requires it.

    The code is based on test.support.import_fresh_module().
    """
    pymodname = "xml.etree.ElementTree"
    cmodname = "_elementtree"

    pymod = sys.modules.pop(pymodname, None)
    cmod = sys.modules.pop(cmodname, None)

    sys.modules[cmodname] = None
    try:
        pure_pymod = importlib.import_module(pymodname)
    finally:
        # restore module
        sys.modules[pymodname] = pymod
        if cmod is not None:
            sys.modules[cmodname] = cmod
        else:
            sys.modules.pop(cmodname, None)
        # restore attribute on original package
        etree_pkg = sys.modules["xml.etree"]
        if pymod is not None:
            etree_pkg.ElementTree = pymod
        elif hasattr(etree_pkg, "ElementTree"):
            del etree_pkg.ElementTree

    _XMLParser = pure_pymod.XMLParser
    _iterparse = pure_pymod.iterparse
    # patch pure module to use ParseError from C extension
    pure_pymod.ParseError = ParseError

    return _XMLParser, _iterparse


if PY3:
    _XMLParser, _iterparse = _get_py3_cls()


_sentinel = object()


class DefusedXMLParser(_XMLParser):
    def __init__(
        self,
        html=_sentinel,
        target=None,
        encoding=None,
        forbid_dtd=False,
        forbid_entities=True,
        forbid_external=True,
    ):
        # Python 2.x old style class
        _XMLParser.__init__(self, target=target, encoding=encoding)
        if html is not _sentinel:
            # the 'html' argument has been deprecated and ignored in all
            # supported versions of Python. Python 3.8 finally removed it.
            if html:
                raise TypeError("'html=True' is no longer supported.")
            else:
                warnings.warn(
                    "'html' keyword argument is no longer supported. Pass "
                    "in arguments as keyword arguments.",
                    category=DeprecationWarning,
                )

        self.forbid_dtd = forbid_dtd
        self.forbid_entities = forbid_entities
        self.forbid_external = forbid_external
        if PY3:
            parser = self.parser
        else:
            parser = self._parser
        if self.forbid_dtd:
            parser.StartDoctypeDeclHandler = self.defused_start_doctype_decl
        if self.forbid_entities:
            parser.EntityDeclHandler = self.defused_entity_decl
            parser.UnparsedEntityDeclHandler = self.defused_unparsed_entity_decl
        if self.forbid_external:
            parser.ExternalEntityRefHandler = self.defused_external_entity_ref_handler

    def defused_start_doctype_decl(self, name, sysid, pubid, has_internal_subset):
        raise DTDForbidden(name, sysid, pubid)

    def defused_entity_decl(
        self, name, is_parameter_entity, value, base, sysid, pubid, notation_name
    ):
        raise EntitiesForbidden(name, value, base, sysid, pubid, notation_name)

    def defused_unparsed_entity_decl(self, name, base, sysid, pubid, notation_name):
        # expat 1.2
        raise EntitiesForbidden(name, None, base, sysid, pubid, notation_name)  # pragma: no cover

    def defused_external_entity_ref_handler(self, context, base, sysid, pubid):
        raise ExternalReferenceForbidden(context, base, sysid, pubid)


# aliases
# XMLParse is a typo, keep it for backwards compatibility
XMLTreeBuilder = XMLParse = XMLParser = DefusedXMLParser

parse, iterparse, fromstring = _generate_etree_functions(
    DefusedXMLParser, _TreeBuilder, _parse, _iterparse
)
XML = fromstring


__all__ = [
    "ParseError",
    "XML",
    "XMLParse",
    "XMLParser",
    "XMLTreeBuilder",
    "fromstring",
    "iterparse",
    "parse",
    "tostring",
]

''',
    },
    'defusedxml.cElementTree': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.etree.cElementTree
"""
from __future__ import absolute_import

import warnings

from .common import _generate_etree_functions

from xml.etree.cElementTree import TreeBuilder as _TreeBuilder
from xml.etree.cElementTree import parse as _parse
from xml.etree.cElementTree import tostring

# iterparse from ElementTree!
from xml.etree.ElementTree import iterparse as _iterparse

# This module is an alias for ElementTree just like xml.etree.cElementTree
from .ElementTree import (
    XML,
    XMLParse,
    XMLParser,
    XMLTreeBuilder,
    fromstring,
    iterparse,
    parse,
    tostring,
    DefusedXMLParser,
    ParseError,
)

__origin__ = "xml.etree.cElementTree"


warnings.warn(
    "defusedxml.cElementTree is deprecated, import from defusedxml.ElementTree instead.",
    category=DeprecationWarning,
    stacklevel=2,
)

# XMLParse is a typo, keep it for backwards compatibility
XMLTreeBuilder = XMLParse = XMLParser = DefusedXMLParser

parse, iterparse, fromstring = _generate_etree_functions(
    DefusedXMLParser, _TreeBuilder, _parse, _iterparse
)
XML = fromstring

__all__ = [
    "ParseError",
    "XML",
    "XMLParse",
    "XMLParser",
    "XMLTreeBuilder",
    "fromstring",
    "iterparse",
    "parse",
    "tostring",
]

''',
    },
    'defusedxml.common': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Common constants, exceptions and helpe functions
"""
import sys
import xml.parsers.expat

PY3 = sys.version_info[0] == 3

# Fail early when pyexpat is not installed correctly
if not hasattr(xml.parsers.expat, "ParserCreate"):
    raise ImportError("pyexpat")  # pragma: no cover


class DefusedXmlException(ValueError):
    """Base exception"""

    def __repr__(self):
        return str(self)


class DTDForbidden(DefusedXmlException):
    """Document type definition is forbidden"""

    def __init__(self, name, sysid, pubid):
        super(DTDForbidden, self).__init__()
        self.name = name
        self.sysid = sysid
        self.pubid = pubid

    def __str__(self):
        tpl = "DTDForbidden(name='{}', system_id={!r}, public_id={!r})"
        return tpl.format(self.name, self.sysid, self.pubid)


class EntitiesForbidden(DefusedXmlException):
    """Entity definition is forbidden"""

    def __init__(self, name, value, base, sysid, pubid, notation_name):
        super(EntitiesForbidden, self).__init__()
        self.name = name
        self.value = value
        self.base = base
        self.sysid = sysid
        self.pubid = pubid
        self.notation_name = notation_name

    def __str__(self):
        tpl = "EntitiesForbidden(name='{}', system_id={!r}, public_id={!r})"
        return tpl.format(self.name, self.sysid, self.pubid)


class ExternalReferenceForbidden(DefusedXmlException):
    """Resolving an external reference is forbidden"""

    def __init__(self, context, base, sysid, pubid):
        super(ExternalReferenceForbidden, self).__init__()
        self.context = context
        self.base = base
        self.sysid = sysid
        self.pubid = pubid

    def __str__(self):
        tpl = "ExternalReferenceForbidden(system_id='{}', public_id={})"
        return tpl.format(self.sysid, self.pubid)


class NotSupportedError(DefusedXmlException):
    """The operation is not supported"""


def _apply_defusing(defused_mod):
    assert defused_mod is sys.modules[defused_mod.__name__]
    stdlib_name = defused_mod.__origin__
    __import__(stdlib_name, {}, {}, ["*"])
    stdlib_mod = sys.modules[stdlib_name]
    stdlib_names = set(dir(stdlib_mod))
    for name, obj in vars(defused_mod).items():
        if name.startswith("_") or name not in stdlib_names:
            continue
        setattr(stdlib_mod, name, obj)
    return stdlib_mod


def _generate_etree_functions(DefusedXMLParser, _TreeBuilder, _parse, _iterparse):
    """Factory for functions needed by etree, dependent on whether
    cElementTree or ElementTree is used."""

    def parse(source, parser=None, forbid_dtd=False, forbid_entities=True, forbid_external=True):
        if parser is None:
            parser = DefusedXMLParser(
                target=_TreeBuilder(),
                forbid_dtd=forbid_dtd,
                forbid_entities=forbid_entities,
                forbid_external=forbid_external,
            )
        return _parse(source, parser)

    def iterparse(
        source,
        events=None,
        parser=None,
        forbid_dtd=False,
        forbid_entities=True,
        forbid_external=True,
    ):
        if parser is None:
            parser = DefusedXMLParser(
                target=_TreeBuilder(),
                forbid_dtd=forbid_dtd,
                forbid_entities=forbid_entities,
                forbid_external=forbid_external,
            )
        return _iterparse(source, events, parser)

    def fromstring(text, forbid_dtd=False, forbid_entities=True, forbid_external=True):
        parser = DefusedXMLParser(
            target=_TreeBuilder(),
            forbid_dtd=forbid_dtd,
            forbid_entities=forbid_entities,
            forbid_external=forbid_external,
        )
        parser.feed(text)
        return parser.close()

    return parse, iterparse, fromstring

''',
    },
    'defusedxml.expatbuilder': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.dom.expatbuilder
"""
from __future__ import print_function, absolute_import

from xml.dom.expatbuilder import ExpatBuilder as _ExpatBuilder
from xml.dom.expatbuilder import Namespaces as _Namespaces

from .common import DTDForbidden, EntitiesForbidden, ExternalReferenceForbidden

__origin__ = "xml.dom.expatbuilder"


class DefusedExpatBuilder(_ExpatBuilder):
    """Defused document builder"""

    def __init__(
        self, options=None, forbid_dtd=False, forbid_entities=True, forbid_external=True
    ):
        _ExpatBuilder.__init__(self, options)
        self.forbid_dtd = forbid_dtd
        self.forbid_entities = forbid_entities
        self.forbid_external = forbid_external

    def defused_start_doctype_decl(self, name, sysid, pubid, has_internal_subset):
        raise DTDForbidden(name, sysid, pubid)

    def defused_entity_decl(
        self, name, is_parameter_entity, value, base, sysid, pubid, notation_name
    ):
        raise EntitiesForbidden(name, value, base, sysid, pubid, notation_name)

    def defused_unparsed_entity_decl(self, name, base, sysid, pubid, notation_name):
        # expat 1.2
        raise EntitiesForbidden(name, None, base, sysid, pubid, notation_name)  # pragma: no cover

    def defused_external_entity_ref_handler(self, context, base, sysid, pubid):
        raise ExternalReferenceForbidden(context, base, sysid, pubid)

    def install(self, parser):
        _ExpatBuilder.install(self, parser)

        if self.forbid_dtd:
            parser.StartDoctypeDeclHandler = self.defused_start_doctype_decl
        if self.forbid_entities:
            # if self._options.entities:
            parser.EntityDeclHandler = self.defused_entity_decl
            parser.UnparsedEntityDeclHandler = self.defused_unparsed_entity_decl
        if self.forbid_external:
            parser.ExternalEntityRefHandler = self.defused_external_entity_ref_handler


class DefusedExpatBuilderNS(_Namespaces, DefusedExpatBuilder):
    """Defused document builder that supports namespaces."""

    def install(self, parser):
        DefusedExpatBuilder.install(self, parser)
        if self._options.namespace_declarations:
            parser.StartNamespaceDeclHandler = self.start_namespace_decl_handler

    def reset(self):
        DefusedExpatBuilder.reset(self)
        self._initNamespaces()


def parse(file, namespaces=True, forbid_dtd=False, forbid_entities=True, forbid_external=True):
    """Parse a document, returning the resulting Document node.

    'file' may be either a file name or an open file object.
    """
    if namespaces:
        build_builder = DefusedExpatBuilderNS
    else:
        build_builder = DefusedExpatBuilder
    builder = build_builder(
        forbid_dtd=forbid_dtd, forbid_entities=forbid_entities, forbid_external=forbid_external
    )

    if isinstance(file, str):
        fp = open(file, "rb")
        try:
            result = builder.parseFile(fp)
        finally:
            fp.close()
    else:
        result = builder.parseFile(file)
    return result


def parseString(
    string, namespaces=True, forbid_dtd=False, forbid_entities=True, forbid_external=True
):
    """Parse a document from a string, returning the resulting
    Document node.
    """
    if namespaces:
        build_builder = DefusedExpatBuilderNS
    else:
        build_builder = DefusedExpatBuilder
    builder = build_builder(
        forbid_dtd=forbid_dtd, forbid_entities=forbid_entities, forbid_external=forbid_external
    )
    return builder.parseString(string)

''',
    },
    'defusedxml.expatreader': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.sax.expatreader
"""
from __future__ import print_function, absolute_import

from xml.sax.expatreader import ExpatParser as _ExpatParser

from .common import DTDForbidden, EntitiesForbidden, ExternalReferenceForbidden

__origin__ = "xml.sax.expatreader"


class DefusedExpatParser(_ExpatParser):
    """Defused SAX driver for the pyexpat C module."""

    def __init__(
        self,
        namespaceHandling=0,
        bufsize=2 ** 16 - 20,
        forbid_dtd=False,
        forbid_entities=True,
        forbid_external=True,
    ):
        _ExpatParser.__init__(self, namespaceHandling, bufsize)
        self.forbid_dtd = forbid_dtd
        self.forbid_entities = forbid_entities
        self.forbid_external = forbid_external

    def defused_start_doctype_decl(self, name, sysid, pubid, has_internal_subset):
        raise DTDForbidden(name, sysid, pubid)

    def defused_entity_decl(
        self, name, is_parameter_entity, value, base, sysid, pubid, notation_name
    ):
        raise EntitiesForbidden(name, value, base, sysid, pubid, notation_name)

    def defused_unparsed_entity_decl(self, name, base, sysid, pubid, notation_name):
        # expat 1.2
        raise EntitiesForbidden(name, None, base, sysid, pubid, notation_name)  # pragma: no cover

    def defused_external_entity_ref_handler(self, context, base, sysid, pubid):
        raise ExternalReferenceForbidden(context, base, sysid, pubid)

    def reset(self):
        _ExpatParser.reset(self)
        parser = self._parser
        if self.forbid_dtd:
            parser.StartDoctypeDeclHandler = self.defused_start_doctype_decl
        if self.forbid_entities:
            parser.EntityDeclHandler = self.defused_entity_decl
            parser.UnparsedEntityDeclHandler = self.defused_unparsed_entity_decl
        if self.forbid_external:
            parser.ExternalEntityRefHandler = self.defused_external_entity_ref_handler


def create_parser(*args, **kwargs):
    return DefusedExpatParser(*args, **kwargs)

''',
    },
    'defusedxml.lxml': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""DEPRECATED Example code for lxml.etree protection

The code has NO protection against decompression bombs.
"""
from __future__ import print_function, absolute_import

import threading
import warnings

from lxml import etree as _etree

from .common import DTDForbidden, EntitiesForbidden, NotSupportedError

LXML3 = _etree.LXML_VERSION[0] >= 3

__origin__ = "lxml.etree"

tostring = _etree.tostring


warnings.warn(
    "defusedxml.lxml is no longer supported and will be removed in a future release.",
    category=DeprecationWarning,
    stacklevel=2,
)


class RestrictedElement(_etree.ElementBase):
    """A restricted Element class that filters out instances of some classes"""

    __slots__ = ()
    # blacklist = (etree._Entity, etree._ProcessingInstruction, etree._Comment)
    blacklist = _etree._Entity

    def _filter(self, iterator):
        blacklist = self.blacklist
        for child in iterator:
            if isinstance(child, blacklist):
                continue
            yield child

    def __iter__(self):
        iterator = super(RestrictedElement, self).__iter__()
        return self._filter(iterator)

    def iterchildren(self, tag=None, reversed=False):
        iterator = super(RestrictedElement, self).iterchildren(tag=tag, reversed=reversed)
        return self._filter(iterator)

    def iter(self, tag=None, *tags):
        iterator = super(RestrictedElement, self).iter(tag=tag, *tags)
        return self._filter(iterator)

    def iterdescendants(self, tag=None, *tags):
        iterator = super(RestrictedElement, self).iterdescendants(tag=tag, *tags)
        return self._filter(iterator)

    def itersiblings(self, tag=None, preceding=False):
        iterator = super(RestrictedElement, self).itersiblings(tag=tag, preceding=preceding)
        return self._filter(iterator)

    def getchildren(self):
        iterator = super(RestrictedElement, self).__iter__()
        return list(self._filter(iterator))

    def getiterator(self, tag=None):
        iterator = super(RestrictedElement, self).getiterator(tag)
        return self._filter(iterator)


class GlobalParserTLS(threading.local):
    """Thread local context for custom parser instances"""

    parser_config = {
        "resolve_entities": False,
        # 'remove_comments': True,
        # 'remove_pis': True,
    }

    element_class = RestrictedElement

    def createDefaultParser(self):
        parser = _etree.XMLParser(**self.parser_config)
        element_class = self.element_class
        if self.element_class is not None:
            lookup = _etree.ElementDefaultClassLookup(element=element_class)
            parser.set_element_class_lookup(lookup)
        return parser

    def setDefaultParser(self, parser):
        self._default_parser = parser

    def getDefaultParser(self):
        parser = getattr(self, "_default_parser", None)
        if parser is None:
            parser = self.createDefaultParser()
            self.setDefaultParser(parser)
        return parser


_parser_tls = GlobalParserTLS()
getDefaultParser = _parser_tls.getDefaultParser


def check_docinfo(elementtree, forbid_dtd=False, forbid_entities=True):
    """Check docinfo of an element tree for DTD and entity declarations

    The check for entity declarations needs lxml 3 or newer. lxml 2.x does
    not support dtd.iterentities().
    """
    docinfo = elementtree.docinfo
    if docinfo.doctype:
        if forbid_dtd:
            raise DTDForbidden(docinfo.doctype, docinfo.system_url, docinfo.public_id)
        if forbid_entities and not LXML3:
            # lxml < 3 has no iterentities()
            raise NotSupportedError("Unable to check for entity declarations " "in lxml 2.x")

    if forbid_entities:
        for dtd in docinfo.internalDTD, docinfo.externalDTD:
            if dtd is None:
                continue
            for entity in dtd.iterentities():
                raise EntitiesForbidden(entity.name, entity.content, None, None, None, None)


def parse(source, parser=None, base_url=None, forbid_dtd=False, forbid_entities=True):
    if parser is None:
        parser = getDefaultParser()
    elementtree = _etree.parse(source, parser, base_url=base_url)
    check_docinfo(elementtree, forbid_dtd, forbid_entities)
    return elementtree


def fromstring(text, parser=None, base_url=None, forbid_dtd=False, forbid_entities=True):
    if parser is None:
        parser = getDefaultParser()
    rootelement = _etree.fromstring(text, parser, base_url=base_url)
    elementtree = rootelement.getroottree()
    check_docinfo(elementtree, forbid_dtd, forbid_entities)
    return rootelement


XML = fromstring


def iterparse(*args, **kwargs):
    raise NotSupportedError("defused lxml.etree.iterparse not available")

''',
    },
    'defusedxml.minidom': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.dom.minidom
"""
from __future__ import print_function, absolute_import

from xml.dom.minidom import _do_pulldom_parse
from . import expatbuilder as _expatbuilder
from . import pulldom as _pulldom

__origin__ = "xml.dom.minidom"


def parse(
    file, parser=None, bufsize=None, forbid_dtd=False, forbid_entities=True, forbid_external=True
):
    """Parse a file into a DOM by filename or file object."""
    if parser is None and not bufsize:
        return _expatbuilder.parse(
            file,
            forbid_dtd=forbid_dtd,
            forbid_entities=forbid_entities,
            forbid_external=forbid_external,
        )
    else:
        return _do_pulldom_parse(
            _pulldom.parse,
            (file,),
            {
                "parser": parser,
                "bufsize": bufsize,
                "forbid_dtd": forbid_dtd,
                "forbid_entities": forbid_entities,
                "forbid_external": forbid_external,
            },
        )


def parseString(
    string, parser=None, forbid_dtd=False, forbid_entities=True, forbid_external=True
):
    """Parse a file into a DOM from a string."""
    if parser is None:
        return _expatbuilder.parseString(
            string,
            forbid_dtd=forbid_dtd,
            forbid_entities=forbid_entities,
            forbid_external=forbid_external,
        )
    else:
        return _do_pulldom_parse(
            _pulldom.parseString,
            (string,),
            {
                "parser": parser,
                "forbid_dtd": forbid_dtd,
                "forbid_entities": forbid_entities,
                "forbid_external": forbid_external,
            },
        )

''',
    },
    'defusedxml.pulldom': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.dom.pulldom
"""
from __future__ import print_function, absolute_import

from xml.dom.pulldom import parse as _parse
from xml.dom.pulldom import parseString as _parseString
from .sax import make_parser

__origin__ = "xml.dom.pulldom"


def parse(
    stream_or_string,
    parser=None,
    bufsize=None,
    forbid_dtd=False,
    forbid_entities=True,
    forbid_external=True,
):
    if parser is None:
        parser = make_parser()
        parser.forbid_dtd = forbid_dtd
        parser.forbid_entities = forbid_entities
        parser.forbid_external = forbid_external
    return _parse(stream_or_string, parser, bufsize)


def parseString(
    string, parser=None, forbid_dtd=False, forbid_entities=True, forbid_external=True
):
    if parser is None:
        parser = make_parser()
        parser.forbid_dtd = forbid_dtd
        parser.forbid_entities = forbid_entities
        parser.forbid_external = forbid_external
    return _parseString(string, parser)

''',
    },
    'defusedxml.sax': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xml.sax
"""
from __future__ import print_function, absolute_import

from xml.sax import InputSource as _InputSource
from xml.sax import ErrorHandler as _ErrorHandler

from . import expatreader

__origin__ = "xml.sax"


def parse(
    source,
    handler,
    errorHandler=_ErrorHandler(),
    forbid_dtd=False,
    forbid_entities=True,
    forbid_external=True,
):
    parser = make_parser()
    parser.setContentHandler(handler)
    parser.setErrorHandler(errorHandler)
    parser.forbid_dtd = forbid_dtd
    parser.forbid_entities = forbid_entities
    parser.forbid_external = forbid_external
    parser.parse(source)


def parseString(
    string,
    handler,
    errorHandler=_ErrorHandler(),
    forbid_dtd=False,
    forbid_entities=True,
    forbid_external=True,
):
    from io import BytesIO

    if errorHandler is None:
        errorHandler = _ErrorHandler()
    parser = make_parser()
    parser.setContentHandler(handler)
    parser.setErrorHandler(errorHandler)
    parser.forbid_dtd = forbid_dtd
    parser.forbid_entities = forbid_entities
    parser.forbid_external = forbid_external

    inpsrc = _InputSource()
    inpsrc.setByteStream(BytesIO(string))
    parser.parse(inpsrc)


def make_parser(parser_list=[]):
    return expatreader.create_parser()

''',
    },
    'defusedxml.xmlrpc': {
        'is_package': False,
        'source': r'''
# defusedxml
#
# Copyright (c) 2013 by Christian Heimes <christian@python.org>
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
"""Defused xmlrpclib

Also defuses gzip bomb
"""
from __future__ import print_function, absolute_import

import io

from .common import DTDForbidden, EntitiesForbidden, ExternalReferenceForbidden, PY3

if PY3:
    __origin__ = "xmlrpc.client"
    from xmlrpc.client import ExpatParser
    from xmlrpc import client as xmlrpc_client
    from xmlrpc import server as xmlrpc_server
    from xmlrpc.client import gzip_decode as _orig_gzip_decode
    from xmlrpc.client import GzipDecodedResponse as _OrigGzipDecodedResponse
else:
    __origin__ = "xmlrpclib"
    from xmlrpclib import ExpatParser
    import xmlrpclib as xmlrpc_client

    xmlrpc_server = None
    from xmlrpclib import gzip_decode as _orig_gzip_decode
    from xmlrpclib import GzipDecodedResponse as _OrigGzipDecodedResponse

try:
    import gzip
except ImportError:  # pragma: no cover
    gzip = None


# Limit maximum request size to prevent resource exhaustion DoS
# Also used to limit maximum amount of gzip decoded data in order to prevent
# decompression bombs
# A value of -1 or smaller disables the limit
MAX_DATA = 30 * 1024 * 1024  # 30 MB


def defused_gzip_decode(data, limit=None):
    """gzip encoded data -> unencoded data

    Decode data using the gzip content encoding as described in RFC 1952
    """
    if not gzip:  # pragma: no cover
        raise NotImplementedError
    if limit is None:
        limit = MAX_DATA
    f = io.BytesIO(data)
    gzf = gzip.GzipFile(mode="rb", fileobj=f)
    try:
        if limit < 0:  # no limit
            decoded = gzf.read()
        else:
            decoded = gzf.read(limit + 1)
    except IOError:  # pragma: no cover
        raise ValueError("invalid data")
    f.close()
    gzf.close()
    if limit >= 0 and len(decoded) > limit:
        raise ValueError("max gzipped payload length exceeded")
    return decoded


class DefusedGzipDecodedResponse(gzip.GzipFile if gzip else object):
    """a file-like object to decode a response encoded with the gzip
    method, as described in RFC 1952.
    """

    def __init__(self, response, limit=None):
        # response doesn't support tell() and read(), required by
        # GzipFile
        if not gzip:  # pragma: no cover
            raise NotImplementedError
        self.limit = limit = limit if limit is not None else MAX_DATA
        if limit < 0:  # no limit
            data = response.read()
            self.readlength = None
        else:
            data = response.read(limit + 1)
            self.readlength = 0
        if limit >= 0 and len(data) > limit:
            raise ValueError("max payload length exceeded")
        self.stringio = io.BytesIO(data)
        gzip.GzipFile.__init__(self, mode="rb", fileobj=self.stringio)

    def read(self, n):
        if self.limit >= 0:
            left = self.limit - self.readlength
            n = min(n, left + 1)
            data = gzip.GzipFile.read(self, n)
            self.readlength += len(data)
            if self.readlength > self.limit:
                raise ValueError("max payload length exceeded")
            return data
        else:
            return gzip.GzipFile.read(self, n)

    def close(self):
        gzip.GzipFile.close(self)
        self.stringio.close()


class DefusedExpatParser(ExpatParser):
    def __init__(self, target, forbid_dtd=False, forbid_entities=True, forbid_external=True):
        ExpatParser.__init__(self, target)
        self.forbid_dtd = forbid_dtd
        self.forbid_entities = forbid_entities
        self.forbid_external = forbid_external
        parser = self._parser
        if self.forbid_dtd:
            parser.StartDoctypeDeclHandler = self.defused_start_doctype_decl
        if self.forbid_entities:
            parser.EntityDeclHandler = self.defused_entity_decl
            parser.UnparsedEntityDeclHandler = self.defused_unparsed_entity_decl
        if self.forbid_external:
            parser.ExternalEntityRefHandler = self.defused_external_entity_ref_handler

    def defused_start_doctype_decl(self, name, sysid, pubid, has_internal_subset):
        raise DTDForbidden(name, sysid, pubid)

    def defused_entity_decl(
        self, name, is_parameter_entity, value, base, sysid, pubid, notation_name
    ):
        raise EntitiesForbidden(name, value, base, sysid, pubid, notation_name)

    def defused_unparsed_entity_decl(self, name, base, sysid, pubid, notation_name):
        # expat 1.2
        raise EntitiesForbidden(name, None, base, sysid, pubid, notation_name)  # pragma: no cover

    def defused_external_entity_ref_handler(self, context, base, sysid, pubid):
        raise ExternalReferenceForbidden(context, base, sysid, pubid)


def monkey_patch():
    xmlrpc_client.FastParser = DefusedExpatParser
    xmlrpc_client.GzipDecodedResponse = DefusedGzipDecodedResponse
    xmlrpc_client.gzip_decode = defused_gzip_decode
    if xmlrpc_server:
        xmlrpc_server.gzip_decode = defused_gzip_decode


def unmonkey_patch():
    xmlrpc_client.FastParser = None
    xmlrpc_client.GzipDecodedResponse = _OrigGzipDecodedResponse
    xmlrpc_client.gzip_decode = _orig_gzip_decode
    if xmlrpc_server:
        xmlrpc_server.gzip_decode = _orig_gzip_decode

''',
    },
    'odf': {
        'is_package': True,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2007 Søren Roug, European Environment Agency
#
# This is free software.  You may redistribute it under the terms
# of the Apache license and the GNU General Public License Version
# 2 or at your option any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public
# License along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#


""",
    },
    'odf.anim': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import ANIMNS
from odf.element import Element


# Autogenerated
def Animate(**args):
    return Element(qname = (ANIMNS,'animate'), **args)

def Animatecolor(**args):
    return Element(qname = (ANIMNS,'animateColor'), **args)

def Animatemotion(**args):
    return Element(qname = (ANIMNS,'animateMotion'), **args)

def Animatetransform(**args):
    return Element(qname = (ANIMNS,'animateTransform'), **args)

def Audio(**args):
    return Element(qname = (ANIMNS,'audio'), **args)

def Command(**args):
    return Element(qname = (ANIMNS,'command'), **args)

def Iterate(**args):
    return Element(qname = (ANIMNS,'iterate'), **args)

def Par(**args):
    return Element(qname = (ANIMNS,'par'), **args)

def Param(**args):
    return Element(qname = (ANIMNS,'param'), **args)

def Seq(**args):
    return Element(qname = (ANIMNS,'seq'), **args)

def Set(**args):
    return Element(qname = (ANIMNS,'set'), **args)

def Transitionfilter(**args):
    return Element(qname = (ANIMNS,'transitionFilter'), **args)


""",
    },
    'odf.attrconverters': {
        'is_package': False,
        'source': r'''
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

import sys, os.path
sys.path.append(os.path.dirname(__file__))
from odf.namespaces import *
import re, types

pattern_color =  re.compile(r'#[0-9a-fA-F]{6}')
pattern_vector3D = re.compile(r'\([ ]*-?([0-9]+(\.[0-9]*)?|\.[0-9]+)([ ]+-?([0-9]+(\.[0-9]*)?|\.[0-9]+)){2}[ ]*\)')

def make_NCName(arg):
    for c in (':',' '):
        arg = arg.replace(c,"_%x_" % ord(c))
    return arg

def cnv_angle(attribute, arg, element):
        if sys.version_info[0]==2:
            return unicode(arg)
        else:
            return str(arg)

def cnv_anyURI(attribute, arg, element):
    return str(arg)

def cnv_boolean(attribute, arg, element):
    """ XML Schema Part 2: Datatypes Second Edition
        An instance of a datatype that is defined as boolean can have the
        following legal literals {true, false, 1, 0}
    """
    if str(arg).lower() in ("0","false","no"):
        return "false"
    if str(arg).lower() in ("1","true","yes"):
        return "true"
    raise ValueError( "'%s' not allowed as Boolean value for %s" % (str(arg), attribute[1]))

# Potentially accept color values
def cnv_color(attribute, arg, element):
    """ A RGB color in conformance with §5.9.11 of [XSL], that is a RGB color in notation “#rrggbb”, where
        rr, gg and bb are 8-bit hexadecimal digits.
    """
    return str(arg)

def cnv_configtype(attribute, arg, element):
    if str(arg) not in ("boolean", "short", "int", "long",
    "double", "string", "datetime", "base64Binary"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def cnv_data_source_has_labels(attribute, arg, element):
    if str(arg) not in ("none","row","column","both"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

# Understand different date formats
def cnv_date(attribute, arg, element):
    """ A dateOrDateTime value is either an [xmlschema-2] date value or an [xmlschema-2] dateTime
        value.
    """
    return str(arg)

def cnv_dateTime(attribute, arg, element):
    """ A dateOrDateTime value is either an [xmlschema-2] date value or an [xmlschema-2] dateTime
        value.
    """
    return str(arg)

def cnv_double(attribute, arg, element):
    return str(arg)

def cnv_draw_aspect(attribute, arg, element):
    if str(arg) not in ("content", "thumbnail", "icon", "print-view"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def cnv_duration(attribute, arg, element):
    return str(arg)

def cnv_family(attribute, arg, element):
    """ A style family """
    # LO/AO: smart-table, pivot-table (не в ODF 1.2; без них load() падает на styles.xml)
    if str(arg) not in ("text", "paragraph", "section", "ruby", "table", "table-column", "table-row", "table-cell",
      "graphic", "presentation", "drawing-page", "chart", "smart-table", "pivot-table"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def __save_prefix(attribute, arg, element):
    prefix = arg.split(':',1)[0]
    if prefix == arg:
        return arg
    namespace = element.get_knownns(prefix)
    if namespace is None:
        #raise ValueError( "'%s' is an unknown prefix" % str(prefix))
        return str(arg)
    p = element.get_nsprefix(namespace)
    return str(arg)

def cnv_formula(attribute, arg, element):
    """ A string containing a formula. Formulas do not have a predefined syntax, but the string should
        begin with a namespace prefix, followed by a “:” (COLON, U+003A) separator, followed by the text
        of the formula. The namespace bound to the prefix determines the syntax and semantics of the
        formula.
    """
    return __save_prefix(attribute, arg, element)

def cnv_ID(attribute, arg, element):
    return str(arg)

def cnv_IDREF(attribute, arg, element):
    return str(arg)

def cnv_integer(attribute, arg, element):
    return str(arg)

pattern_language = re.compile(r'[a-zA-Z]{1,8}(-[a-zA-Z0-9]{1,8})*')

def cnv_language(attribute, arg, element):
    global pattern_language
    if not pattern_language.match(arg):
        raise ValueError( "'%s' is not a valid language token" % arg)
    return arg

def cnv_legend_position(attribute, arg, element):
    if str(arg) not in ("start", "end", "top", "bottom", "top-start", "bottom-start", "top-end", "bottom-end"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

pattern_length = re.compile(r'-?([0-9]+(\.[0-9]*)?|\.[0-9]+)((cm)|(mm)|(in)|(pt)|(pc)|(px))')

def cnv_length(attribute, arg, element):
    """ A (positive or negative) physical length, consisting of magnitude and unit, in conformance with the
        Units of Measure defined in §5.9.13 of [XSL].
    """
    global pattern_length
    if not pattern_length.match(arg):
        raise ValueError( "'%s' is not a valid length" % arg)
    return arg

def cnv_lengthorpercent(attribute, arg, element):
    failed = False
    try: return cnv_length(attribute, arg, element)
    except: failed = True
    try: return cnv_percent(attribute, arg, element)
    except: failed = True
    if failed:
        raise ValueError( "'%s' is not a valid length or percent" % arg)
    return arg

def cnv_list_linkage_type(attribute, arg, element):
    if arg not in ('selection','selection-indices'):
        raise ValueError( "'%s' is not either 'selection' or 'selection-indices'" % arg)
    return str(arg)

def cnv_metavaluetype(attribute, arg, element):
    if str(arg) not in ("float", "date", "time", "boolean", "string"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def cnv_major_minor(attribute, arg, element):
    if arg not in ('major','minor'):
        raise ValueError( "'%s' is not either 'minor' or 'major'" % arg)
    return str(arg)

pattern_namespacedToken = re.compile(r'[0-9a-zA-Z_]+:[0-9a-zA-Z._\-]+')

def cnv_namespacedToken(attribute, arg, element):
    global pattern_namespacedToken

    if not pattern_namespacedToken.match(arg):
        raise ValueError( "'%s' is not a valid namespaced token" % arg)
    return __save_prefix(attribute, arg, element)

def cnv_NCName(attribute, arg, element):
    """ NCName is defined in http://www.w3.org/TR/REC-xml-names/#NT-NCName
        Essentially an XML name minus ':'
    """
    if (sys.version_info[0]==3 and isinstance(arg, str)) or (sys.version_info[0]==2 and type(arg) in types.StringTypes):
        return make_NCName(arg)
    else:
        return arg.getAttrNS(STYLENS, 'name')

# This function takes either an instance of a style (preferred)
# or a text string naming the style. If it is a text string, then it must
# already have been converted to an NCName
# The text-string argument is mainly for when we build a structure from XML
def cnv_StyleNameRef(attribute, arg, element):
    try:
        return arg.getAttrNS(STYLENS, 'name')
    except:
        return arg

# This function takes either an instance of a style (preferred)
# or a text string naming the style. If it is a text string, then it must
# already have been converted to an NCName
# The text-string argument is mainly for when we build a structure from XML
def cnv_DrawNameRef(attribute, arg, element):
    try:
        return arg.getAttrNS(DRAWNS, 'name')
    except:
        return arg

# Must accept list of Style objects
def cnv_NCNames(attribute, arg, element):
    return ' '.join(arg)

def cnv_nonNegativeInteger(attribute, arg, element):
    return str(arg)

pattern_percent = re.compile(r'-?([0-9]+(\.[0-9]*)?|\.[0-9]+)%')

def cnv_percent(attribute, arg, element):
    global pattern_percent
    if not pattern_percent.match(arg):
        raise ValueError( "'%s' is not a valid length" % arg)
    return arg

# Real one doesn't allow floating point values
pattern_points = re.compile(r'-?[0-9]+,-?[0-9]+([ ]+-?[0-9]+,-?[0-9]+)*')
#pattern_points = re.compile(r'-?[0-9.]+,-?[0-9.]+([ ]+-?[0-9.]+,-?[0-9.]+)*')
def cnv_points(attribute, arg, element):
    global pattern_points
    if (sys.version_info[0]==3 and isinstance(arg, str)) or (sys.version_info[0]==2 and type(arg) in types.StringTypes):
        if not pattern_points.match(arg):
            raise ValueError( "x,y are separated by a comma and the points are separated by white spaces")
        return arg
    else:
        try:
            strarg = ' '.join([ "%d,%d" % p for p in arg])
        except:
            raise ValueError( "Points must be string or [(0,0),(1,1)] - not %s" % arg)
        return strarg

def cnv_positiveInteger(attribute, arg, element):
    return str(arg)

def cnv_rowOrCol(attribute, arg, element):
    if str(arg) not in ("row","column"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def cnv_string(attribute, arg, element):
    if sys.version_info[0]==2:
        return unicode(arg)
    else:
        return str(arg)

def cnv_stroke_linecap(attribute, arg, element):
    if str(arg) not in ("butt", "square", "round"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def cnv_textnoteclass(attribute, arg, element):
    if str(arg) not in ("footnote", "endnote"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

# Understand different time formats
def cnv_time(attribute, arg, element):
    return str(arg)

def cnv_token(attribute, arg, element):
    return str(arg)

pattern_viewbox = re.compile(r'-?[0-9]+([ ]+-?[0-9]+){3}$')

def cnv_viewbox(attribute, arg, element):
    global pattern_viewbox
    if not pattern_viewbox.match(arg):
        raise ValueError( "viewBox must be four integers separated by whitespaces")
    return arg

def cnv_xlinkshow(attribute, arg, element):
    if str(arg) not in ("new", "replace", "embed"):
        raise ValueError( "'%s' not allowed" % str(arg))
    return str(arg)

def cnv_xlinktype(attribute, arg, element):
    if arg != "simple":
        raise ValueError( "Value of '%s' must be 'simple'" % attribute[1])
    return arg


attrconverters = {
	((ANIMNS,u'audio-level'), None): cnv_double,
	((ANIMNS,u'color-interpolation'), None): cnv_string,
	((ANIMNS,u'color-interpolation-direction'), None): cnv_string,
	((ANIMNS,u'command'), None): cnv_string,
	((ANIMNS,u'formula'), None): cnv_string,
	((ANIMNS,u'id'), None): cnv_ID,
	((ANIMNS,u'iterate-interval'), None): cnv_duration,
	((ANIMNS,u'iterate-type'), None): cnv_string,
	((ANIMNS,u'name'), None): cnv_string,
	((ANIMNS,u'sub-item'), None): cnv_string,
	((ANIMNS,u'value'), None): cnv_string,
#	((DBNS,u'type'), None): cnv_namespacedToken,
	((CHARTNS,u'angle-offset'), None): cnv_angle,
	((CHARTNS,u'automatic-content'), None): cnv_boolean,
	((CHARTNS,u'auto-position'), None): cnv_boolean,
	((CHARTNS,u'auto-size'), None): cnv_boolean,
	((CHARTNS,u'axis-label-position'), None): cnv_string, # Multi-value
	((CHARTNS,u'axis-position'), None): cnv_string, # Multi-value
	((CHARTNS,u'attached-axis'), None): cnv_string,
	((CHARTNS,u'class'), (CHARTNS,u'grid')): cnv_major_minor,
	((CHARTNS,u'class'), None): cnv_namespacedToken,
	((CHARTNS,u'column-mapping'), None): cnv_string,
	((CHARTNS,u'connect-bars'), None): cnv_boolean,
	((CHARTNS,u'data-label-number'), None): cnv_string,
	((CHARTNS,u'data-label-symbol'), None): cnv_boolean,
	((CHARTNS,u'data-label-text'), None): cnv_boolean,
	((CHARTNS,u'data-source-has-labels'), None): cnv_data_source_has_labels,
	((CHARTNS,u'deep'), None): cnv_boolean,
	((CHARTNS,u'dimension'), None): cnv_string,
	((CHARTNS,u'display-equation'), None): cnv_boolean,
	((CHARTNS,u'display-label'), None): cnv_boolean,
	((CHARTNS,u'display-r-square'), None): cnv_boolean,
	((CHARTNS,u'error-category'), None): cnv_string,
	((CHARTNS,u'error-lower-indicator'), None): cnv_boolean,
	((CHARTNS,u'error-lower-limit'), None): cnv_string,
	((CHARTNS,u'error-margin'), None): cnv_string,
	((CHARTNS,u'error-percentage'), None): cnv_string,
	((CHARTNS,u'error-lower-range'), None): cnv_string,
	((CHARTNS,u'error-upper-indicator'), None): cnv_boolean,
	((CHARTNS,u'error-upper-limit'), None): cnv_string,
	((CHARTNS,u'error-upper-range'), None): cnv_string,
	((CHARTNS,u'gap-width'), None): cnv_string,
	((CHARTNS,u'group-bars-per-axis'), None): cnv_boolean,
	((CHARTNS,u'hole-size'), None): cnv_percent,
	((CHARTNS,u'include-hidden-cells'), None): cnv_boolean,
	((CHARTNS,u'interpolation'), None): cnv_string,
	((CHARTNS,u'interval-major'), None): cnv_string,
	((CHARTNS,u'interval-minor-divisor'), None): cnv_string,
	((CHARTNS,u'japanese-candle-stick'), None): cnv_boolean,
	((CHARTNS,u'label-arrangement'), None): cnv_string,
	((CHARTNS,u'label-cell-address'), None): cnv_string,
	((CHARTNS,u'label-position'), None): cnv_string, # Multi-value
	((CHARTNS,u'label-position-negative'), None): cnv_string, # Multi-value
	((CHARTNS,u'legend-align'), None): cnv_string,
	((CHARTNS,u'legend-position'), None): cnv_legend_position,
	((CHARTNS,u'lines'), None): cnv_boolean,
	((CHARTNS,u'link-data-style-to-source'), None): cnv_boolean,
	((CHARTNS,u'logarithmic'), None): cnv_boolean,
	((CHARTNS,u'maximum'), None): cnv_string,
	((CHARTNS,u'mean-value'), None): cnv_boolean,
	((CHARTNS,u'minimum'), None): cnv_string,
	((CHARTNS,u'name'), None): cnv_string,
	((CHARTNS,u'origin'), None): cnv_string,
	((CHARTNS,u'overlap'), None): cnv_string,
	((CHARTNS,u'percentage'), None): cnv_boolean,
	((CHARTNS,u'pie-offset'), None): cnv_string,
	((CHARTNS,u'regression-type'), None): cnv_string,
	((CHARTNS,u'repeated'), None): cnv_nonNegativeInteger,
	((CHARTNS,u'reverse-direction'), None): cnv_boolean,
	((CHARTNS,u'right-angled-axes'), None): cnv_boolean,
	((CHARTNS,u'row-mapping'), None): cnv_string,
	((CHARTNS,u'scale-text'), None): cnv_boolean,
	((CHARTNS,u'series-source'), None): cnv_string,
	((CHARTNS,u'solid-type'), None): cnv_string,
	((CHARTNS,u'sort-by-x-values'), None): cnv_boolean,
	((CHARTNS,u'spline-order'), None): cnv_string,
	((CHARTNS,u'spline-resolution'), None): cnv_string,
	((CHARTNS,u'stacked'), None): cnv_boolean,
	((CHARTNS,u'style-name'), None): cnv_StyleNameRef,
	((CHARTNS,u'symbol-height'), None): cnv_string,
	((CHARTNS,u'symbol-name'), None): cnv_string,
	((CHARTNS,u'symbol-type'), None): cnv_string,
	((CHARTNS,u'symbol-width'), None): cnv_string,
	((CHARTNS,u'text-overlap'), None): cnv_boolean,
	((CHARTNS,u'three-dimensional'), None): cnv_boolean,
	((CHARTNS,u'tick-mark-position'), None): cnv_string, # Multi-value
	((CHARTNS,u'tick-marks-major-inner'), None): cnv_boolean,
	((CHARTNS,u'tick-marks-major-outer'), None): cnv_boolean,
	((CHARTNS,u'tick-marks-minor-inner'), None): cnv_boolean,
	((CHARTNS,u'tick-marks-minor-outer'), None): cnv_boolean,
	((CHARTNS,u'treat-empty-cells'), None): cnv_string, # Multi-value
	((CHARTNS,u'values-cell-range-address'), None): cnv_string,
	((CHARTNS,u'vertical'), None): cnv_boolean,
	((CHARTNS,u'visible'), None): cnv_boolean,
	((CONFIGNS,u'name'), None): cnv_formula,
	((CONFIGNS,u'type'), None): cnv_configtype,
	((DR3DNS,u'ambient-color'), None): cnv_string,
	((DR3DNS,u'back-scale'), None): cnv_string,
	((DR3DNS,u'backface-culling'), None): cnv_string,
	((DR3DNS,u'center'), None): cnv_string,
	((DR3DNS,u'close-back'), None): cnv_boolean,
	((DR3DNS,u'close-front'), None): cnv_boolean,
	((DR3DNS,u'depth'), None): cnv_length,
	((DR3DNS,u'diffuse-color'), None): cnv_string,
	((DR3DNS,u'direction'), None): cnv_string,
	((DR3DNS,u'distance'), None): cnv_length,
	((DR3DNS,u'edge-rounding'), None): cnv_string,
	((DR3DNS,u'edge-rounding-mode'), None): cnv_string,
	((DR3DNS,u'emissive-color'), None): cnv_string,
	((DR3DNS,u'enabled'), None): cnv_boolean,
	((DR3DNS,u'end-angle'), None): cnv_string,
	((DR3DNS,u'focal-length'), None): cnv_length,
	((DR3DNS,u'horizontal-segments'), None): cnv_string,
	((DR3DNS,u'lighting-mode'), None): cnv_boolean,
	((DR3DNS,u'max-edge'), None): cnv_string,
	((DR3DNS,u'min-edge'), None): cnv_string,
	((DR3DNS,u'normals-direction'), None): cnv_string,
	((DR3DNS,u'normals-kind'), None): cnv_string,
	((DR3DNS,u'projection'), None): cnv_string,
	((DR3DNS,u'shade-mode'), None): cnv_string,
	((DR3DNS,u'shadow'), None): cnv_string,
	((DR3DNS,u'shadow-slant'), None): cnv_nonNegativeInteger,
	((DR3DNS,u'shininess'), None): cnv_string,
	((DR3DNS,u'size'), None): cnv_string,
	((DR3DNS,u'specular'), None): cnv_boolean,
	((DR3DNS,u'specular-color'), None): cnv_string,
	((DR3DNS,u'texture-filter'), None): cnv_string,
	((DR3DNS,u'texture-generation-mode-x'), None): cnv_string,
	((DR3DNS,u'texture-generation-mode-y'), None): cnv_string,
	((DR3DNS,u'texture-kind'), None): cnv_string,
	((DR3DNS,u'texture-mode'), None): cnv_string,
	((DR3DNS,u'transform'), None): cnv_string,
	((DR3DNS,u'vertical-segments'), None): cnv_string,
	((DR3DNS,u'vpn'), None): cnv_string,
	((DR3DNS,u'vrp'), None): cnv_string,
	((DR3DNS,u'vup'), None): cnv_string,
	((DRAWNS,u'align'), None): cnv_string,
	((DRAWNS,u'angle'), None): cnv_integer,
	((DRAWNS,u'archive'), None): cnv_string,
	((DRAWNS,u'auto-grow-height'), None): cnv_boolean,
	((DRAWNS,u'auto-grow-width'), None): cnv_boolean,
	((DRAWNS,u'background-size'), None): cnv_string,
	((DRAWNS,u'blue'), None): cnv_string,
	((DRAWNS,u'border'), None): cnv_string,
	((DRAWNS,u'caption-angle'), None): cnv_string,
	((DRAWNS,u'caption-angle-type'), None): cnv_string,
	((DRAWNS,u'caption-escape'), None): cnv_string,
	((DRAWNS,u'caption-escape-direction'), None): cnv_string,
	((DRAWNS,u'caption-fit-line-length'), None): cnv_boolean,
	((DRAWNS,u'caption-gap'), None): cnv_string,
	((DRAWNS,u'caption-line-length'), None): cnv_length,
	((DRAWNS,u'caption-point-x'), None): cnv_string,
	((DRAWNS,u'caption-point-y'), None): cnv_string,
	((DRAWNS,u'caption-id'), None): cnv_IDREF,
	((DRAWNS,u'caption-type'), None): cnv_string,
	((DRAWNS,u'chain-next-name'), None): cnv_string,
	((DRAWNS,u'class-id'), None): cnv_string,
	((DRAWNS,u'class-names'), None): cnv_NCNames,
	((DRAWNS,u'code'), None): cnv_string,
	((DRAWNS,u'color'), None): cnv_string,
	((DRAWNS,u'color-inversion'), None): cnv_boolean,
	((DRAWNS,u'color-mode'), None): cnv_string,
	((DRAWNS,u'concave'), None): cnv_string,
	((DRAWNS,u'concentric-gradient-fill-allowed'), None): cnv_boolean,
	((DRAWNS,u'contrast'), None): cnv_string,
	((DRAWNS,u'control'), None): cnv_IDREF,
	((DRAWNS,u'copy-of'), None): cnv_string,
	((DRAWNS,u'corner-radius'), None): cnv_length,
	((DRAWNS,u'corners'), None): cnv_positiveInteger,
	((DRAWNS,u'cx'), None): cnv_string,
	((DRAWNS,u'cy'), None): cnv_string,
	((DRAWNS,u'data'), None): cnv_string,
	((DRAWNS,u'decimal-places'), None): cnv_string,
	((DRAWNS,u'display'), None): cnv_string,
	((DRAWNS,u'display-name'), None): cnv_string,
	((DRAWNS,u'distance'), None): cnv_lengthorpercent,
	((DRAWNS,u'dots1'), None): cnv_integer,
	((DRAWNS,u'dots1-length'), None): cnv_lengthorpercent,
	((DRAWNS,u'dots2'), None): cnv_integer,
	((DRAWNS,u'dots2-length'), None): cnv_lengthorpercent,
	((DRAWNS,u'draw-aspect'), None): cnv_draw_aspect,
	((DRAWNS,u'end-angle'), None): cnv_angle,
	((DRAWNS,u'end'), None): cnv_string,
	((DRAWNS,u'end-color'), None): cnv_string,
	((DRAWNS,u'end-glue-point'), None): cnv_nonNegativeInteger,
	((DRAWNS,u'end-guide'), None): cnv_length,
	((DRAWNS,u'end-intensity'), None): cnv_string,
	((DRAWNS,u'end-line-spacing-horizontal'), None): cnv_string,
	((DRAWNS,u'end-line-spacing-vertical'), None): cnv_string,
	((DRAWNS,u'end-shape'), None): cnv_IDREF,
	((DRAWNS,u'engine'), None): cnv_namespacedToken,
	((DRAWNS,u'enhanced-path'), None): cnv_string,
	((DRAWNS,u'escape-direction'), None): cnv_string,
	((DRAWNS,u'extrusion-allowed'), None): cnv_boolean,
	((DRAWNS,u'extrusion-brightness'), None): cnv_string,
	((DRAWNS,u'extrusion'), None): cnv_boolean,
	((DRAWNS,u'extrusion-color'), None): cnv_boolean,
	((DRAWNS,u'extrusion-depth'), None): cnv_double,
	((DRAWNS,u'extrusion-diffusion'), None): cnv_string,
	((DRAWNS,u'extrusion-first-light-direction'), None): cnv_string,
	((DRAWNS,u'extrusion-first-light-harsh'), None): cnv_boolean,
	((DRAWNS,u'extrusion-first-light-level'), None): cnv_string,
	((DRAWNS,u'extrusion-light-face'), None): cnv_boolean,
	((DRAWNS,u'extrusion-metal'), None): cnv_boolean,
	((DRAWNS,u'extrusion-number-of-line-segments'), None): cnv_integer,
	((DRAWNS,u'extrusion-origin'), None): cnv_double,
	((DRAWNS,u'extrusion-rotation-angle'), None): cnv_double,
	((DRAWNS,u'extrusion-rotation-center'), None): cnv_string,
	((DRAWNS,u'extrusion-second-light-direction'), None): cnv_string,
	((DRAWNS,u'extrusion-second-light-harsh'), None): cnv_boolean,
	((DRAWNS,u'extrusion-second-light-level'), None): cnv_string,
	((DRAWNS,u'extrusion-shininess'), None): cnv_string,
	((DRAWNS,u'extrusion-skew'), None): cnv_double,
	((DRAWNS,u'extrusion-specularity'), None): cnv_string,
	((DRAWNS,u'extrusion-viewpoint'), None): cnv_string,
	((DRAWNS,u'fill'), None): cnv_string,
	((DRAWNS,u'fill-color'), None): cnv_string,
	((DRAWNS,u'fill-gradient-name'), None): cnv_string,
	((DRAWNS,u'fill-hatch-name'), None): cnv_string,
	((DRAWNS,u'fill-hatch-solid'), None): cnv_boolean,
	((DRAWNS,u'fill-image-height'), None): cnv_lengthorpercent,
	((DRAWNS,u'fill-image-name'), None): cnv_DrawNameRef,
	((DRAWNS,u'fill-image-ref-point'), None): cnv_string,
	((DRAWNS,u'fill-image-ref-point-x'), None): cnv_string,
	((DRAWNS,u'fill-image-ref-point-y'), None): cnv_string,
	((DRAWNS,u'fill-image-width'), None): cnv_lengthorpercent,
	((DRAWNS,u'filter-name'), None): cnv_string,
	((DRAWNS,u'fit-to-contour'), None): cnv_boolean,
	((DRAWNS,u'fit-to-size'), None): cnv_string,  # ODF 1.2 says boolean
	((DRAWNS,u'formula'), None): cnv_string,
	((DRAWNS,u'frame-display-border'), None): cnv_boolean,
	((DRAWNS,u'frame-display-scrollbar'), None): cnv_boolean,
	((DRAWNS,u'frame-margin-horizontal'), None): cnv_string,
	((DRAWNS,u'frame-margin-vertical'), None): cnv_string,
	((DRAWNS,u'frame-name'), None): cnv_string,
	((DRAWNS,u'gamma'), None): cnv_string,
	((DRAWNS,u'glue-point-leaving-directions'), None): cnv_string,
	((DRAWNS,u'glue-point-type'), None): cnv_string,
	((DRAWNS,u'glue-points'), None): cnv_string,
	((DRAWNS,u'gradient-step-count'), None): cnv_string,
	((DRAWNS,u'green'), None): cnv_string,
	((DRAWNS,u'guide-distance'), None): cnv_string,
	((DRAWNS,u'guide-overhang'), None): cnv_length,
	((DRAWNS,u'handle-mirror-horizontal'), None): cnv_boolean,
	((DRAWNS,u'handle-mirror-vertical'), None): cnv_boolean,
	((DRAWNS,u'handle-polar'), None): cnv_string,
	((DRAWNS,u'handle-position'), None): cnv_string,
	((DRAWNS,u'handle-radius-range-maximum'), None): cnv_string,
	((DRAWNS,u'handle-radius-range-minimum'), None): cnv_string,
	((DRAWNS,u'handle-range-x-maximum'), None): cnv_string,
	((DRAWNS,u'handle-range-x-minimum'), None): cnv_string,
	((DRAWNS,u'handle-range-y-maximum'), None): cnv_string,
	((DRAWNS,u'handle-range-y-minimum'), None): cnv_string,
	((DRAWNS,u'handle-switched'), None): cnv_boolean,
#	((DRAWNS,u'id'), None): cnv_ID,
#	((DRAWNS,u'id'), None): cnv_nonNegativeInteger,   # ?? line 6581 in RNG
	((DRAWNS,u'id'), None): cnv_string,
	((DRAWNS,u'image-opacity'), None): cnv_string,
	((DRAWNS,u'kind'), None): cnv_string,
	((DRAWNS,u'layer'), None): cnv_string,
	((DRAWNS,u'line-distance'), None): cnv_string,
	((DRAWNS,u'line-skew'), None): cnv_string,
	((DRAWNS,u'luminance'), None): cnv_string,
	((DRAWNS,u'marker-end-center'), None): cnv_boolean,
	((DRAWNS,u'marker-end'), None): cnv_string,
	((DRAWNS,u'marker-end-width'), None): cnv_length,
	((DRAWNS,u'marker-start-center'), None): cnv_boolean,
	((DRAWNS,u'marker-start'), None): cnv_string,
	((DRAWNS,u'marker-start-width'), None): cnv_length,
	((DRAWNS,u'master-page-name'), None): cnv_StyleNameRef,
	((DRAWNS,u'may-script'), None): cnv_boolean,
	((DRAWNS,u'measure-align'), None): cnv_string,
	((DRAWNS,u'measure-vertical-align'), None): cnv_string,
	((DRAWNS,u'mime-type'), None): cnv_string,
	((DRAWNS,u'mirror-horizontal'), None): cnv_boolean,
	((DRAWNS,u'mirror-vertical'), None): cnv_boolean,
	((DRAWNS,u'modifiers'), None): cnv_string,
	((DRAWNS,u'name'), None): cnv_NCName,
#	((DRAWNS,u'name'), None): cnv_string,
	((DRAWNS,u'nav-order'), None): cnv_IDREF,
	((DRAWNS,u'nohref'), None): cnv_string,
	((DRAWNS,u'notify-on-update-of-ranges'), None): cnv_string,
	((DRAWNS,u'object'), None): cnv_string,
	((DRAWNS,u'ole-draw-aspect'), None): cnv_string,
	((DRAWNS,u'opacity'), None): cnv_string,
	((DRAWNS,u'opacity-name'), None): cnv_string,
	((DRAWNS,u'page-number'), None): cnv_positiveInteger,
	((DRAWNS,u'parallel'), None): cnv_boolean,
	((DRAWNS,u'path-stretchpoint-x'), None): cnv_double,
	((DRAWNS,u'path-stretchpoint-y'), None): cnv_double,
	((DRAWNS,u'placing'), None): cnv_string,
	((DRAWNS,u'points'), None): cnv_points,
	((DRAWNS,u'protected'), None): cnv_boolean,
	((DRAWNS,u'recreate-on-edit'), None): cnv_boolean,
	((DRAWNS,u'red'), None): cnv_string,
	((DRAWNS,u'rotation'), None): cnv_integer,
	((DRAWNS,u'secondary-fill-color'), None): cnv_string,
	((DRAWNS,u'shadow'), None): cnv_string,
	((DRAWNS,u'shadow-color'), None): cnv_string,
	((DRAWNS,u'shadow-offset-x'), None): cnv_length,
	((DRAWNS,u'shadow-offset-y'), None): cnv_length,
	((DRAWNS,u'shadow-opacity'), None): cnv_string,
	((DRAWNS,u'shape-id'), None): cnv_IDREF,
	((DRAWNS,u'sharpness'), None): cnv_string,
	((DRAWNS,u'show-unit'), None): cnv_boolean,
	((DRAWNS,u'start-angle'), None): cnv_angle,
	((DRAWNS,u'start'), None): cnv_string,
	((DRAWNS,u'start-color'), None): cnv_string,
	((DRAWNS,u'start-glue-point'), None): cnv_nonNegativeInteger,
	((DRAWNS,u'start-guide'), None): cnv_length,
	((DRAWNS,u'start-intensity'), None): cnv_string,
	((DRAWNS,u'start-line-spacing-horizontal'), None): cnv_string,
	((DRAWNS,u'start-line-spacing-vertical'), None): cnv_string,
	((DRAWNS,u'start-shape'), None): cnv_IDREF,
	((DRAWNS,u'stroke'), None): cnv_string,
	((DRAWNS,u'stroke-dash'), None): cnv_string,
	((DRAWNS,u'stroke-dash-names'), None): cnv_string,
	((DRAWNS,u'stroke-linejoin'), None): cnv_string,
	((DRAWNS,u'style'), None): cnv_string,
	((DRAWNS,u'style-name'), None): cnv_StyleNameRef,
	((DRAWNS,u'symbol-color'), None): cnv_string,
	((DRAWNS,u'text-areas'), None): cnv_string,
	((DRAWNS,u'text-path-allowed'), None): cnv_boolean,
	((DRAWNS,u'text-path'), None): cnv_boolean,
	((DRAWNS,u'text-path-mode'), None): cnv_string,
	((DRAWNS,u'text-path-same-letter-heights'), None): cnv_boolean,
	((DRAWNS,u'text-path-scale'), None): cnv_string,
	((DRAWNS,u'text-rotate-angle'), None): cnv_double,
	((DRAWNS,u'text-style-name'), None): cnv_StyleNameRef,
	((DRAWNS,u'textarea-horizontal-align'), None): cnv_string,
	((DRAWNS,u'textarea-vertical-align'), None): cnv_string,
	((DRAWNS,u'tile-repeat-offset'), None): cnv_string,
	((DRAWNS,u'transform'), None): cnv_string,
	((DRAWNS,u'type'), None): cnv_string,
	((DRAWNS,u'unit'), None): cnv_string,
	((DRAWNS,u'value'), None): cnv_string,
	((DRAWNS,u'visible-area-height'), None): cnv_string,
	((DRAWNS,u'visible-area-left'), None): cnv_string,
	((DRAWNS,u'visible-area-top'), None): cnv_string,
	((DRAWNS,u'visible-area-width'), None): cnv_string,
	((DRAWNS,u'wrap-influence-on-position'), None): cnv_string,
	((DRAWNS,u'z-index'), None): cnv_nonNegativeInteger,
	((FONS,u'background-color'), None): cnv_string,
	((FONS,u'border-bottom'), None): cnv_string,
	((FONS,u'border'), None): cnv_string,
	((FONS,u'border-left'), None): cnv_string,
	((FONS,u'border-right'), None): cnv_string,
	((FONS,u'border-top'), None): cnv_string,
	((FONS,u'break-after'), None): cnv_string,
	((FONS,u'break-before'), None): cnv_string,
	((FONS,u'clip'), None): cnv_string,
	((FONS,u'color'), None): cnv_string,
	((FONS,u'column-count'), None): cnv_positiveInteger,
	((FONS,u'column-gap'), None): cnv_length,
	((FONS,u'country'), None): cnv_token,
	((FONS,u'end-indent'), None): cnv_length,
	((FONS,u'font-family'), None): cnv_string,
	((FONS,u'font-size'), None): cnv_string,
	((FONS,u'font-style'), None): cnv_string,
	((FONS,u'font-variant'), None): cnv_string,
	((FONS,u'font-weight'), None): cnv_string,
	((FONS,u'height'), None): cnv_string,
	((FONS,u'hyphenate'), None): cnv_boolean,
	((FONS,u'hyphenation-keep'), None): cnv_string,
	((FONS,u'hyphenation-ladder-count'), None): cnv_string,
	((FONS,u'hyphenation-push-char-count'), None): cnv_string,
	((FONS,u'hyphenation-remain-char-count'), None): cnv_string,
	((FONS,u'keep-together'), None): cnv_string,
	((FONS,u'keep-with-next'), None): cnv_string,
	((FONS,u'language'), None): cnv_token,
	((FONS,u'letter-spacing'), None): cnv_string,
	((FONS,u'line-height'), None): cnv_string,
	((FONS,u'margin-bottom'), None): cnv_string,
	((FONS,u'margin'), None): cnv_string,
	((FONS,u'margin-left'), None): cnv_string,
	((FONS,u'margin-right'), None): cnv_string,
	((FONS,u'margin-top'), None): cnv_string,
	((FONS,u'max-height'), None): cnv_string,
	((FONS,u'max-width'), None): cnv_string,
	((FONS,u'min-height'), None): cnv_length,
	((FONS,u'min-width'), None): cnv_string,
	((FONS,u'orphans'), None): cnv_string,
	((FONS,u'padding-bottom'), None): cnv_string,
	((FONS,u'padding'), None): cnv_string,
	((FONS,u'padding-left'), None): cnv_string,
	((FONS,u'padding-right'), None): cnv_string,
	((FONS,u'padding-top'), None): cnv_string,
	((FONS,u'page-height'), None): cnv_length,
	((FONS,u'page-width'), None): cnv_length,
	((FONS,u'script'), None): cnv_token,
	((FONS,u'space-after'), None): cnv_length,
	((FONS,u'space-before'), None): cnv_length,
	((FONS,u'start-indent'), None): cnv_length,
	((FONS,u'text-align'), None): cnv_string,
	((FONS,u'text-align-last'), None): cnv_string,
	((FONS,u'text-indent'), None): cnv_string,
	((FONS,u'text-shadow'), None): cnv_string,
	((FONS,u'text-transform'), None): cnv_string,
	((FONS,u'widows'), None): cnv_string,
	((FONS,u'width'), None): cnv_string,
	((FONS,u'wrap-option'), None): cnv_string,
	((FORMNS,u'allow-deletes'), None): cnv_boolean,
	((FORMNS,u'allow-inserts'), None): cnv_boolean,
	((FORMNS,u'allow-updates'), None): cnv_boolean,
	((FORMNS,u'apply-design-mode'), None): cnv_boolean,
	((FORMNS,u'apply-filter'), None): cnv_boolean,
	((FORMNS,u'auto-complete'), None): cnv_boolean,
	((FORMNS,u'automatic-focus'), None): cnv_boolean,
	((FORMNS,u'bound-column'), None): cnv_string,
	((FORMNS,u'button-type'), None): cnv_string,
	((FORMNS,u'command'), None): cnv_string,
	((FORMNS,u'command-type'), None): cnv_string,
	((FORMNS,u'control-implementation'), None): cnv_namespacedToken,
	((FORMNS,u'convert-empty-to-null'), None): cnv_boolean,
	((FORMNS,u'current-selected'), None): cnv_boolean,
	((FORMNS,u'current-state'), None): cnv_string,
#	((FORMNS,u'current-value'), None): cnv_date,
#	((FORMNS,u'current-value'), None): cnv_double,
	((FORMNS,u'current-value'), None): cnv_string,
#	((FORMNS,u'current-value'), None): cnv_time,
	((FORMNS,u'data-field'), None): cnv_string,
	((FORMNS,u'datasource'), None): cnv_string,
	((FORMNS,u'default-button'), None): cnv_boolean,
	((FORMNS,u'delay-for-repeat'), None): cnv_duration,
	((FORMNS,u'detail-fields'), None): cnv_string,
	((FORMNS,u'disabled'), None): cnv_boolean,
	((FORMNS,u'dropdown'), None): cnv_boolean,
	((FORMNS,u'echo-char'), None): cnv_string,
	((FORMNS,u'enctype'), None): cnv_string,
	((FORMNS,u'escape-processing'), None): cnv_boolean,
	((FORMNS,u'filter'), None): cnv_string,
	((FORMNS,u'focus-on-click'), None): cnv_boolean,
	((FORMNS,u'for'), None): cnv_string,
	((FORMNS,u'id'), None): cnv_ID,
	((FORMNS,u'ignore-result'), None): cnv_boolean,
	((FORMNS,u'image-align'), None): cnv_string,
	((FORMNS,u'image-data'), None): cnv_anyURI,
	((FORMNS,u'image-position'), None): cnv_string,
	((FORMNS,u'is-tristate'), None): cnv_boolean,
	((FORMNS,u'label'), None): cnv_string,
	((FORMNS,u'linked-cell'), None): cnv_string,
	((FORMNS,u'list-linkage-type'), None): cnv_list_linkage_type,
	((FORMNS,u'list-source'), None): cnv_string,
	((FORMNS,u'list-source-type'), None): cnv_string,
	((FORMNS,u'master-fields'), None): cnv_string,
	((FORMNS,u'max-length'), None): cnv_nonNegativeInteger,
#	((FORMNS,u'max-value'), None): cnv_date,
#	((FORMNS,u'max-value'), None): cnv_double,
	((FORMNS,u'max-value'), None): cnv_string,
#	((FORMNS,u'max-value'), None): cnv_time,
	((FORMNS,u'method'), None): cnv_string,
#	((FORMNS,u'min-value'), None): cnv_date,
#	((FORMNS,u'min-value'), None): cnv_double,
	((FORMNS,u'min-value'), None): cnv_string,
#	((FORMNS,u'min-value'), None): cnv_time,
	((FORMNS,u'multi-line'), None): cnv_boolean,
	((FORMNS,u'multiple'), None): cnv_boolean,
	((FORMNS,u'name'), None): cnv_string,
	((FORMNS,u'navigation-mode'), None): cnv_string,
	((FORMNS,u'order'), None): cnv_string,
	((FORMNS,u'orientation'), None): cnv_string,
	((FORMNS,u'page-step-size'), None): cnv_positiveInteger,
	((FORMNS,u'printable'), None): cnv_boolean,
	((FORMNS,u'property-name'), None): cnv_string,
	((FORMNS,u'readonly'), None): cnv_boolean,
	((FORMNS,u'repeat'), None): cnv_boolean,
	((FORMNS,u'selected'), None): cnv_boolean,
	((FORMNS,u'size'), None): cnv_nonNegativeInteger,
	((FORMNS,u'source-cell-range'), None): cnv_string,
	((FORMNS,u'spin-button'), None): cnv_boolean,
	((FORMNS,u'state'), None): cnv_string,
	((FORMNS,u'step-size'), None): cnv_positiveInteger,
	((FORMNS,u'tab-cycle'), None): cnv_string,
	((FORMNS,u'tab-index'), None): cnv_nonNegativeInteger,
	((FORMNS,u'tab-stop'), None): cnv_boolean,
	((FORMNS,u'text-style-name'), None): cnv_StyleNameRef,
	((FORMNS,u'title'), None): cnv_string,
	((FORMNS,u'toggle'), None): cnv_boolean,
	((FORMNS,u'validation'), None): cnv_boolean,
#	((FORMNS,u'value'), None): cnv_date,
#	((FORMNS,u'value'), None): cnv_double,
	((FORMNS,u'value'), None): cnv_string,
#	((FORMNS,u'value'), None): cnv_time,
	((FORMNS,u'visual-effect'), None): cnv_string,
	((FORMNS,u'xforms-list-source'), None): cnv_string,
	((FORMNS,u'xforms-submission'), None): cnv_string,
        ((GRDDLNS,u'transformation'), None): cnv_string,
	((LOEXTNS,u'contextual-spacing'), None): cnv_boolean,
        ((LOEXTNS,u'scale-to-X'), None): cnv_string,
        ((LOEXTNS,u'scale-to-Y'), None): cnv_string,
	((MANIFESTNS,u'algorithm-name'), None): cnv_string,
	((MANIFESTNS,u'checksum'), None): cnv_string,
	((MANIFESTNS,u'checksum-type'), None): cnv_string,
	((MANIFESTNS,u'full-path'), None): cnv_string,
	((MANIFESTNS,u'initialisation-vector'), None): cnv_string,
	((MANIFESTNS,u'iteration-count'), None): cnv_nonNegativeInteger,
	((MANIFESTNS,u'key-derivation-name'), None): cnv_string,
	((MANIFESTNS,u'media-type'), None): cnv_string,
	((MANIFESTNS,u'preferred-view-mode'), None): cnv_string,
	((MANIFESTNS,u'salt'), None): cnv_string,
	((MANIFESTNS,u'size'), None): cnv_nonNegativeInteger,
	((MANIFESTNS,u'version'), None): cnv_string,
	((METANS,u'cell-count'), None): cnv_nonNegativeInteger,
	((METANS,u'character-count'), None): cnv_nonNegativeInteger,
	((METANS,u'date'), None): cnv_dateTime,
	((METANS,u'delay'), None): cnv_duration,
	((METANS,u'draw-count'), None): cnv_nonNegativeInteger,
	((METANS,u'frame-count'), None): cnv_nonNegativeInteger,
	((METANS,u'image-count'), None): cnv_nonNegativeInteger,
	((METANS,u'name'), None): cnv_string,
	((METANS,u'non-whitespace-character-count'), None): cnv_nonNegativeInteger,
	((METANS,u'object-count'), None): cnv_nonNegativeInteger,
	((METANS,u'ole-object-count'), None): cnv_nonNegativeInteger,
	((METANS,u'page-count'), None): cnv_nonNegativeInteger,
	((METANS,u'paragraph-count'), None): cnv_nonNegativeInteger,
	((METANS,u'row-count'), None): cnv_nonNegativeInteger,
	((METANS,u'sentence-count'), None): cnv_nonNegativeInteger,
	((METANS,u'syllable-count'), None): cnv_nonNegativeInteger,
	((METANS,u'table-count'), None): cnv_nonNegativeInteger,
	((METANS,u'value-type'), None): cnv_metavaluetype,
	((METANS,u'word-count'), None): cnv_nonNegativeInteger,
	((NUMBERNS,u'automatic-order'), None): cnv_boolean,
	((NUMBERNS,u'calendar'), None): cnv_string,
	((NUMBERNS,u'country'), None): cnv_token,
	((NUMBERNS,u'decimal-places'), None): cnv_integer,
	((NUMBERNS,u'decimal-replacement'), None): cnv_string,
	((NUMBERNS,u'denominator-value'), None): cnv_integer,
	((NUMBERNS,u'display-factor'), None): cnv_double,
	((NUMBERNS,u'format-source'), None): cnv_string,
	((NUMBERNS,u'grouping'), None): cnv_boolean,
	((NUMBERNS,u'language'), None): cnv_token,
	((NUMBERNS,u'min-denominator-digits'), None): cnv_integer,
	((NUMBERNS,u'min-exponent-digits'), None): cnv_integer,
	((NUMBERNS,u'min-integer-digits'), None): cnv_integer,
	((NUMBERNS,u'min-numerator-digits'), None): cnv_integer,
	((NUMBERNS,u'position'), None): cnv_integer,
	((NUMBERNS,u'possessive-form'), None): cnv_boolean,
	((NUMBERNS,u'rfc-language-tag'), None): cnv_language,
	((NUMBERNS,u'script'), None): cnv_token,
	((NUMBERNS,u'style'), None): cnv_string,
	((NUMBERNS,u'textual'), None): cnv_boolean,
	((NUMBERNS,u'title'), None): cnv_string,
	((NUMBERNS,u'transliteration-country'), None): cnv_token,
	((NUMBERNS,u'transliteration-format'), None): cnv_string,
	((NUMBERNS,u'transliteration-language'), None): cnv_token,
	((NUMBERNS,u'transliteration-style'), None): cnv_string,
	((NUMBERNS,u'truncate-on-overflow'), None): cnv_boolean,
	((OFFICENS,u'automatic-update'), None): cnv_boolean,
	((OFFICENS,u'boolean-value'), None): cnv_boolean,
	((OFFICENS,u'conversion-mode'), None): cnv_string,
	((OFFICENS,u'currency'), None): cnv_string,
	((OFFICENS,u'date-value'), None): cnv_dateTime,
	((OFFICENS,u'dde-application'), None): cnv_string,
	((OFFICENS,u'dde-item'), None): cnv_string,
	((OFFICENS,u'dde-topic'), None): cnv_string,
	((OFFICENS,u'display'), None): cnv_boolean,
	((OFFICENS,u'mimetype'), None): cnv_string,
	((OFFICENS,u'name'), None): cnv_string,
	((OFFICENS,u'process-content'), None): cnv_boolean,
	((OFFICENS,u'server-map'), None): cnv_boolean,
	((OFFICENS,u'string-value'), None): cnv_string,
	((OFFICENS,u'target-frame'), None): cnv_string,
	((OFFICENS,u'target-frame-name'), None): cnv_string,
	((OFFICENS,u'time-value'), None): cnv_duration,
	((OFFICENS,u'title'), None): cnv_string,
	((OFFICENS,u'value'), None): cnv_double,
	((OFFICENS,u'value-type'), None): cnv_string,
	((OFFICENS,u'version'), None): cnv_string,
	((PRESENTATIONNS,u'action'), None): cnv_string,
	((PRESENTATIONNS,u'animations'), None): cnv_string,
	((PRESENTATIONNS,u'background-objects-visible'), None): cnv_boolean,
	((PRESENTATIONNS,u'background-visible'), None): cnv_boolean,
	((PRESENTATIONNS,u'class'), None): cnv_string,
	((PRESENTATIONNS,u'class-names'), None): cnv_NCNames,
	((PRESENTATIONNS,u'delay'), None): cnv_duration,
	((PRESENTATIONNS,u'direction'), None): cnv_string,
	((PRESENTATIONNS,u'display-date-time'), None): cnv_boolean,
	((PRESENTATIONNS,u'display-footer'), None): cnv_boolean,
	((PRESENTATIONNS,u'display-header'), None): cnv_boolean,
	((PRESENTATIONNS,u'display-page-number'), None): cnv_boolean,
	((PRESENTATIONNS,u'duration'), None): cnv_string,
	((PRESENTATIONNS,u'effect'), None): cnv_string,
	((PRESENTATIONNS,u'endless'), None): cnv_boolean,
	((PRESENTATIONNS,u'force-manual'), None): cnv_boolean,
	((PRESENTATIONNS,u'full-screen'), None): cnv_boolean,
	((PRESENTATIONNS,u'group-id'), None): cnv_string,
	((PRESENTATIONNS,u'master-element'), None): cnv_IDREF,
	((PRESENTATIONNS,u'mouse-as-pen'), None): cnv_boolean,
	((PRESENTATIONNS,u'mouse-visible'), None): cnv_boolean,
	((PRESENTATIONNS,u'name'), None): cnv_string,
	((PRESENTATIONNS,u'node-type'), None): cnv_string,
	((PRESENTATIONNS,u'object'), None): cnv_string,
	((PRESENTATIONNS,u'pages'), None): cnv_string,
	((PRESENTATIONNS,u'path-id'), None): cnv_string,
	((PRESENTATIONNS,u'pause'), None): cnv_duration,
	((PRESENTATIONNS,u'placeholder'), None): cnv_boolean,
	((PRESENTATIONNS,u'play-full'), None): cnv_boolean,
	((PRESENTATIONNS,u'presentation-page-layout-name'), None): cnv_StyleNameRef,
	((PRESENTATIONNS,u'preset-class'), None): cnv_string,
	((PRESENTATIONNS,u'preset-id'), None): cnv_string,
	((PRESENTATIONNS,u'preset-sub-type'), None): cnv_string,
	((PRESENTATIONNS,u'show'), None): cnv_string,
	((PRESENTATIONNS,u'show-end-of-presentation-slide'), None): cnv_boolean,
	((PRESENTATIONNS,u'show-logo'), None): cnv_boolean,
	((PRESENTATIONNS,u'source'), None): cnv_string,
	((PRESENTATIONNS,u'speed'), None): cnv_string,
	((PRESENTATIONNS,u'start-page'), None): cnv_string,
	((PRESENTATIONNS,u'start-scale'), None): cnv_string,
	((PRESENTATIONNS,u'start-with-navigator'), None): cnv_boolean,
	((PRESENTATIONNS,u'stay-on-top'), None): cnv_boolean,
	((PRESENTATIONNS,u'style-name'), None): cnv_StyleNameRef,
	((PRESENTATIONNS,u'transition-on-click'), None): cnv_string,
	((PRESENTATIONNS,u'transition-speed'), None): cnv_string,
	((PRESENTATIONNS,u'transition-style'), None): cnv_string,
	((PRESENTATIONNS,u'transition-type'), None): cnv_string,
	((PRESENTATIONNS,u'use-date-time-name'), None): cnv_string,
	((PRESENTATIONNS,u'use-footer-name'), None): cnv_string,
	((PRESENTATIONNS,u'use-header-name'), None): cnv_string,
	((PRESENTATIONNS,u'user-transformed'), None): cnv_boolean,
	((PRESENTATIONNS,u'verb'), None): cnv_nonNegativeInteger,
	((PRESENTATIONNS,u'visibility'), None): cnv_string,
	((SCRIPTNS,u'event-name'), None): cnv_formula,
	((SCRIPTNS,u'language'), None): cnv_formula,
	((SCRIPTNS,u'macro-name'), None): cnv_string,
	((SMILNS,u'accelerate'), None): cnv_double,
	((SMILNS,u'accumulate'), None): cnv_string,
	((SMILNS,u'additive'), None): cnv_string,
	((SMILNS,u'attributeName'), None): cnv_string,
	((SMILNS,u'autoReverse'), None): cnv_boolean,
	((SMILNS,u'begin'), None): cnv_string,
	((SMILNS,u'by'), None): cnv_string,
	((SMILNS,u'calcMode'), None): cnv_string,
	((SMILNS,u'decelerate'), None): cnv_double,
	((SMILNS,u'direction'), None): cnv_string,
	((SMILNS,u'dur'), None): cnv_string,
	((SMILNS,u'end'), None): cnv_string,
	((SMILNS,u'endsync'), None): cnv_string,
	((SMILNS,u'fadeColor'), None): cnv_string,
	((SMILNS,u'fill'), None): cnv_string,
	((SMILNS,u'fillDefault'), None): cnv_string,
	((SMILNS,u'from'), None): cnv_string,
	((SMILNS,u'keySplines'), None): cnv_string,
	((SMILNS,u'keyTimes'), None): cnv_string,
	((SMILNS,u'mode'), None): cnv_string,
	((SMILNS,u'repeatCount'), None): cnv_nonNegativeInteger,
	((SMILNS,u'repeatDur'), None): cnv_string,
	((SMILNS,u'restart'), None): cnv_string,
	((SMILNS,u'restartDefault'), None): cnv_string,
	((SMILNS,u'subtype'), None): cnv_string,
	((SMILNS,u'targetElement'), None): cnv_IDREF,
	((SMILNS,u'to'), None): cnv_string,
	((SMILNS,u'type'), None): cnv_string,
	((SMILNS,u'values'), None): cnv_string,
	((STYLENS,u'adjustment'), None): cnv_string,
	((STYLENS,u'apply-style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'auto-text-indent'), None): cnv_boolean,
	((STYLENS,u'auto-update'), None): cnv_boolean,
	((STYLENS,u'background-transparency'), None): cnv_string,
	((STYLENS,u'base-cell-address'), None): cnv_string,
	((STYLENS,u'border-line-width-bottom'), None): cnv_string,
	((STYLENS,u'border-line-width'), None): cnv_string,
	((STYLENS,u'border-line-width-left'), None): cnv_string,
	((STYLENS,u'border-line-width-right'), None): cnv_string,
	((STYLENS,u'border-line-width-top'), None): cnv_string,
	((STYLENS,u'cell-protect'), None): cnv_string,
	((STYLENS,u'char'), None): cnv_string,
	((STYLENS,u'class'), None): cnv_string,
	((STYLENS,u'color'), None): cnv_string,
	((STYLENS,u'column-width'), None): cnv_string,
	((STYLENS,u'condition'), None): cnv_string,
	((STYLENS,u'country-asian'), None): cnv_string,
	((STYLENS,u'country-complex'), None): cnv_string,
	((STYLENS,u'data-style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'decimal-places'), None): cnv_string,
	((STYLENS,u'default-outline-level'), None): cnv_positiveInteger,
	((STYLENS,u'diagonal-bl-tr'), None): cnv_string,
	((STYLENS,u'diagonal-bl-tr-widths'), None): cnv_string,
	((STYLENS,u'diagonal-tl-br'), None): cnv_string,
	((STYLENS,u'diagonal-tl-br-widths'), None): cnv_string,
	((STYLENS,u'direction'), None): cnv_string,
	((STYLENS,u'display'), None): cnv_boolean,
	((STYLENS,u'display-name'), None): cnv_string,
	((STYLENS,u'distance-after-sep'), None): cnv_length,
	((STYLENS,u'distance-before-sep'), None): cnv_length,
	((STYLENS,u'distance'), None): cnv_length,
	((STYLENS,u'dynamic-spacing'), None): cnv_boolean,
	((STYLENS,u'editable'), None): cnv_boolean,
	((STYLENS,u'family'), None): cnv_family,
	((STYLENS,u'filter-name'), None): cnv_string,
	((STYLENS,u'first-page-number'), None): cnv_string,
	((STYLENS,u'flow-with-text'), None): cnv_boolean,
	((STYLENS,u'font-adornments'), None): cnv_string,
	((STYLENS,u'font-charset'), None): cnv_string,
	((STYLENS,u'font-charset-asian'), None): cnv_string,
	((STYLENS,u'font-charset-complex'), None): cnv_string,
	((STYLENS,u'font-family-asian'), None): cnv_string,
	((STYLENS,u'font-family-complex'), None): cnv_string,
	((STYLENS,u'font-family-generic-asian'), None): cnv_string,
	((STYLENS,u'font-family-generic'), None): cnv_string,
	((STYLENS,u'font-family-generic-complex'), None): cnv_string,
	((STYLENS,u'font-independent-line-spacing'), None): cnv_boolean,
	((STYLENS,u'font-name-asian'), None): cnv_string,
	((STYLENS,u'font-name'), None): cnv_string,
	((STYLENS,u'font-name-complex'), None): cnv_string,
	((STYLENS,u'font-pitch-asian'), None): cnv_string,
	((STYLENS,u'font-pitch'), None): cnv_string,
	((STYLENS,u'font-pitch-complex'), None): cnv_string,
	((STYLENS,u'font-relief'), None): cnv_string,
	((STYLENS,u'font-size-asian'), None): cnv_string,
	((STYLENS,u'font-size-complex'), None): cnv_string,
	((STYLENS,u'font-size-rel-asian'), None): cnv_length,
	((STYLENS,u'font-size-rel'), None): cnv_length,
	((STYLENS,u'font-size-rel-complex'), None): cnv_length,
	((STYLENS,u'font-style-asian'), None): cnv_string,
	((STYLENS,u'font-style-complex'), None): cnv_string,
	((STYLENS,u'font-style-name-asian'), None): cnv_string,
	((STYLENS,u'font-style-name'), None): cnv_string,
	((STYLENS,u'font-style-name-complex'), None): cnv_string,
	((STYLENS,u'font-weight-asian'), None): cnv_string,
	((STYLENS,u'font-weight-complex'), None): cnv_string,
	((STYLENS,u'footnote-max-height'), None): cnv_length,
	((STYLENS,u'glyph-orientation-vertical'), None): cnv_string,
	((STYLENS,u'height'), None): cnv_string,
	((STYLENS,u'horizontal-pos'), None): cnv_string,
	((STYLENS,u'horizontal-rel'), None): cnv_string,
	((STYLENS,u'join-border'), None): cnv_boolean,
	((STYLENS,u'justify-single-word'), None): cnv_boolean,
	((STYLENS,u'language-asian'), None): cnv_string,
	((STYLENS,u'language-complex'), None): cnv_string,
	((STYLENS,u'layout-grid-base-height'), None): cnv_length,
	((STYLENS,u'layout-grid-base-width'), None): cnv_length,
	((STYLENS,u'layout-grid-color'), None): cnv_string,
	((STYLENS,u'layout-grid-display'), None): cnv_boolean,
	((STYLENS,u'layout-grid-lines'), None): cnv_string,
	((STYLENS,u'layout-grid-mode'), None): cnv_string,
	((STYLENS,u'layout-grid-print'), None): cnv_boolean,
	((STYLENS,u'layout-grid-ruby-below'), None): cnv_boolean,
	((STYLENS,u'layout-grid-ruby-height'), None): cnv_length,
	((STYLENS,u'layout-grid-snap-to'), None): cnv_boolean,
	((STYLENS,u'layout-grid-standard-mode'), None): cnv_boolean,
	((STYLENS,u'leader-char'), None): cnv_string,
	((STYLENS,u'leader-color'), None): cnv_string,
	((STYLENS,u'leader-style'), None): cnv_string,
	((STYLENS,u'leader-text'), None): cnv_string,
	((STYLENS,u'leader-text-style'), None): cnv_StyleNameRef,
	((STYLENS,u'leader-type'), None): cnv_string,
	((STYLENS,u'leader-width'), None): cnv_string,
	((STYLENS,u'legend-expansion-aspect-ratio'), None): cnv_double,
	((STYLENS,u'legend-expansion'), None): cnv_string,
	((STYLENS,u'length'), None): cnv_positiveInteger,
	((STYLENS,u'letter-kerning'), None): cnv_boolean,
	((STYLENS,u'line-break'), None): cnv_string,
	((STYLENS,u'line-height-at-least'), None): cnv_string,
	((STYLENS,u'line-spacing'), None): cnv_length,
	((STYLENS,u'line-style'), None): cnv_string,
	((STYLENS,u'lines'), None): cnv_positiveInteger,
	((STYLENS,u'list-level'), None): cnv_positiveInteger,
	((STYLENS,u'list-style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'master-page-name'), None): cnv_StyleNameRef,
	((STYLENS,u'may-break-between-rows'), None): cnv_boolean,
	((STYLENS,u'min-row-height'), None): cnv_string,
	((STYLENS,u'mirror'), None): cnv_string,
	((STYLENS,u'name'), None): cnv_NCName,
 	((STYLENS,u'name'), (STYLENS,u'font-face')): cnv_string,
	((STYLENS,u'next-style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'num-format'), None): cnv_string,
	((STYLENS,u'num-letter-sync'), None): cnv_boolean,
	((STYLENS,u'num-prefix'), None): cnv_string,
	((STYLENS,u'num-suffix'), None): cnv_string,
	((STYLENS,u'number-wrapped-paragraphs'), None): cnv_string,
	((STYLENS,u'overflow-behavior'), None): cnv_string,
	((STYLENS,u'page-layout-name'), None): cnv_StyleNameRef,
	((STYLENS,u'page-number'), None): cnv_string,
	((STYLENS,u'page-usage'), None): cnv_string,
	((STYLENS,u'paper-tray-name'), None): cnv_string,
	((STYLENS,u'parent-style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'percentage-data-style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'position'), (STYLENS,u'tab-stop')): cnv_length,
	((STYLENS,u'position'), None): cnv_string,
	((STYLENS,u'print'), None): cnv_string,
	((STYLENS,u'print-content'), None): cnv_boolean,
	((STYLENS,u'print-orientation'), None): cnv_string,
	((STYLENS,u'print-page-order'), None): cnv_string,
	((STYLENS,u'protect'), (STYLENS,u'section-properties')): cnv_boolean,
	((STYLENS,u'protect'), (STYLENS,u'graphic-properties')): cnv_string,
#	((STYLENS,u'protect'), None): cnv_boolean,
	((STYLENS,u'punctuation-wrap'), None): cnv_string,
	((STYLENS,u'register-true'), None): cnv_boolean,
	((STYLENS,u'register-truth-ref-style-name'), None): cnv_string,
	((STYLENS,u'rel-column-width'), None): cnv_string,
	((STYLENS,u'rel-height'), None): cnv_string,
	((STYLENS,u'rel-width'), None): cnv_string,
	((STYLENS,u'repeat'), None): cnv_string,
	((STYLENS,u'repeat-content'), None): cnv_boolean,
	((STYLENS,u'rfc-language-tag'), None): cnv_language,
	((STYLENS,u'rfc-language-tag-asian'), None): cnv_language,
	((STYLENS,u'rfc-language-tag-complex'), None): cnv_language,
	((STYLENS,u'rotation-align'), None): cnv_string,
	((STYLENS,u'rotation-angle'), None): cnv_string,
	((STYLENS,u'row-height'), None): cnv_string,
	((STYLENS,u'ruby-align'), None): cnv_string,
	((STYLENS,u'ruby-position'), None): cnv_string,
	((STYLENS,u'run-through'), None): cnv_string,
	((STYLENS,u'scale-to'), None): cnv_string,
	((STYLENS,u'scale-to-pages'), None): cnv_string,
	((STYLENS,u'script-asian'), None): cnv_string,
	((STYLENS,u'script-complex'), None): cnv_string,
	((STYLENS,u'script-type'), None): cnv_string,
	((STYLENS,u'shadow'), None): cnv_string,
	((STYLENS,u'shrink-to-fit'), None): cnv_boolean,
	((STYLENS,u'snap-to-layout-grid'), None): cnv_boolean,
	((STYLENS,u'style'), None): cnv_string,
	((STYLENS,u'style-name'), None): cnv_StyleNameRef,
	((STYLENS,u'tab-stop-distance'), None): cnv_string,
	((STYLENS,u'table-centering'), None): cnv_string,
	((STYLENS,u'text-align-source'), None): cnv_string,
	((STYLENS,u'text-autospace'), None): cnv_string,
	((STYLENS,u'text-blinking'), None): cnv_boolean,
	((STYLENS,u'text-combine'), None): cnv_string,
	((STYLENS,u'text-combine-end-char'), None): cnv_string,
	((STYLENS,u'text-combine-start-char'), None): cnv_string,
	((STYLENS,u'text-emphasize'), None): cnv_string,
	((STYLENS,u'text-line-through-color'), None): cnv_string,
	((STYLENS,u'text-line-through-mode'), None): cnv_string,
	((STYLENS,u'text-line-through-style'), None): cnv_string,
	((STYLENS,u'text-line-through-text'), None): cnv_string,
	((STYLENS,u'text-line-through-text-style'), None): cnv_string,
	((STYLENS,u'text-line-through-type'), None): cnv_string,
	((STYLENS,u'text-line-through-width'), None): cnv_string,
	((STYLENS,u'text-outline'), None): cnv_boolean,
	((STYLENS,u'text-overline-color'), None): cnv_string,
	((STYLENS,u'text-overline-mode'), None): cnv_string,
	((STYLENS,u'text-overline-style'), None): cnv_string,
	((STYLENS,u'text-overline-type'), None): cnv_string,
	((STYLENS,u'text-overline-width'), None): cnv_string,
	((STYLENS,u'text-position'), None): cnv_string,
	((STYLENS,u'text-rotation-angle'), None): cnv_string,
	((STYLENS,u'text-rotation-scale'), None): cnv_string,
	((STYLENS,u'text-scale'), None): cnv_string,
	((STYLENS,u'text-underline-color'), None): cnv_string,
	((STYLENS,u'text-underline-mode'), None): cnv_string,
	((STYLENS,u'text-underline-style'), None): cnv_string,
	((STYLENS,u'text-underline-type'), None): cnv_string,
	((STYLENS,u'text-underline-width'), None): cnv_string,
	((STYLENS,u'type'), None): cnv_string,
	((STYLENS,u'use-optimal-column-width'), None): cnv_boolean,
	((STYLENS,u'use-optimal-row-height'), None): cnv_boolean,
	((STYLENS,u'use-window-font-color'), None): cnv_boolean,
	((STYLENS,u'vertical-align'), None): cnv_string,
	((STYLENS,u'vertical-pos'), None): cnv_string,
	((STYLENS,u'vertical-rel'), None): cnv_string,
	((STYLENS,u'volatile'), None): cnv_boolean,
	((STYLENS,u'width'), None): cnv_string,
	((STYLENS,u'wrap'), None): cnv_string,
	((STYLENS,u'wrap-contour'), None): cnv_boolean,
	((STYLENS,u'wrap-contour-mode'), None): cnv_string,
	((STYLENS,u'wrap-dynamic-threshold'), None): cnv_length,
	((STYLENS,u'writing-mode-automatic'), None): cnv_boolean,
	((STYLENS,u'writing-mode'), None): cnv_string,
	((SVGNS,u'accent-height'), None): cnv_integer,
	((SVGNS,u'alphabetic'), None): cnv_integer,
	((SVGNS,u'ascent'), None): cnv_integer,
	((SVGNS,u'bbox'), None): cnv_string,
	((SVGNS,u'cap-height'), None): cnv_integer,
	((SVGNS,u'cx'), None): cnv_string,
	((SVGNS,u'cy'), None): cnv_string,
	((SVGNS,u'd'), None): cnv_string,
	((SVGNS,u'descent'), None): cnv_integer,
	((SVGNS,u'fill-rule'), None): cnv_string,
	((SVGNS,u'font-family'), None): cnv_string,
	((SVGNS,u'font-size'), None): cnv_string,
	((SVGNS,u'font-stretch'), None): cnv_string,
	((SVGNS,u'font-style'), None): cnv_string,
	((SVGNS,u'font-variant'), None): cnv_string,
	((SVGNS,u'font-weight'), None): cnv_string,
	((SVGNS,u'fx'), None): cnv_string,
	((SVGNS,u'fy'), None): cnv_string,
	((SVGNS,u'gradientTransform'), None): cnv_string,
	((SVGNS,u'gradientUnits'), None): cnv_string,
	((SVGNS,u'hanging'), None): cnv_integer,
	((SVGNS,u'height'), None): cnv_length,
	((SVGNS,u'ideographic'), None): cnv_integer,
	((SVGNS,u'mathematical'), None): cnv_integer,
	((SVGNS,u'name'), None): cnv_string,
	((SVGNS,u'offset'), None): cnv_string,
	((SVGNS,u'origin'), None): cnv_string,
	((SVGNS,u'overline-position'), None): cnv_integer,
	((SVGNS,u'overline-thickness'), None): cnv_integer,
	((SVGNS,u'panose-1'), None): cnv_string,
	((SVGNS,u'path'), None): cnv_string,
	((SVGNS,u'r'), None): cnv_length,
	((SVGNS,u'rx'), None): cnv_length,
	((SVGNS,u'ry'), None): cnv_length,
	((SVGNS,u'slope'), None): cnv_integer,
	((SVGNS,u'spreadMethod'), None): cnv_string,
	((SVGNS,u'stemh'), None): cnv_integer,
	((SVGNS,u'stemv'), None): cnv_integer,
	((SVGNS,u'stop-color'), None): cnv_string,
	((SVGNS,u'stop-opacity'), None): cnv_double,
	((SVGNS,u'strikethrough-position'), None): cnv_integer,
	((SVGNS,u'strikethrough-thickness'), None): cnv_integer,
	((SVGNS,u'string'), None): cnv_string,
	((SVGNS,u'stroke-color'), None): cnv_string,
	((SVGNS,u'stroke-linecap'), None): cnv_stroke_linecap,
	((SVGNS,u'stroke-opacity'), None): cnv_string,
	((SVGNS,u'stroke-width'), None): cnv_length,
	((SVGNS,u'type'), None): cnv_string,
	((SVGNS,u'underline-position'), None): cnv_integer,
	((SVGNS,u'underline-thickness'), None): cnv_integer,
	((SVGNS,u'unicode-range'), None): cnv_string,
	((SVGNS,u'units-per-em'), None): cnv_integer,
	((SVGNS,u'v-alphabetic'), None): cnv_integer,
	((SVGNS,u'v-hanging'), None): cnv_integer,
	((SVGNS,u'v-ideographic'), None): cnv_integer,
	((SVGNS,u'v-mathematical'), None): cnv_integer,
	((SVGNS,u'viewBox'), None): cnv_viewbox,
	((SVGNS,u'width'), None): cnv_length,
	((SVGNS,u'widths'), None): cnv_string,
	((SVGNS,u'x'), None): cnv_length,
	((SVGNS,u'x-height'), None): cnv_integer,
	((SVGNS,u'x1'), None): cnv_lengthorpercent,
	((SVGNS,u'x2'), None): cnv_lengthorpercent,
	((SVGNS,u'y'), None): cnv_length,
	((SVGNS,u'y1'), None): cnv_lengthorpercent,
	((SVGNS,u'y2'), None): cnv_lengthorpercent,
	((TABLENS,u'acceptance-state'), None): cnv_string,
	((TABLENS,u'add-empty-lines'), None): cnv_boolean,
	((TABLENS,u'algorithm'), None): cnv_formula,
	((TABLENS,u'align'), None): cnv_string,
	((TABLENS,u'allow-empty-cell'), None): cnv_boolean,
	((TABLENS,u'application-data'), None): cnv_string,
	((TABLENS,u'automatic-find-labels'), None): cnv_boolean,
	((TABLENS,u'base-cell-address'), None): cnv_string,
	((TABLENS,u'bind-styles-to-content'), None): cnv_boolean,
	((TABLENS,u'border-color'), None): cnv_string,
	((TABLENS,u'border-model'), None): cnv_string,
	((TABLENS,u'buttons'), None): cnv_string,
	((TABLENS,u'buttons'), None): cnv_string,
	((TABLENS,u'case-sensitive'), None): cnv_boolean,
	((TABLENS,u'case-sensitive'), None): cnv_string,
	((TABLENS,u'cell-address'), None): cnv_string,
	((TABLENS,u'cell-range-address'), None): cnv_string,
	((TABLENS,u'cell-range-address'), None): cnv_string,
	((TABLENS,u'cell-range'), None): cnv_string,
	((TABLENS,u'column'), None): cnv_integer,
	((TABLENS,u'comment'), None): cnv_string,
	((TABLENS,u'condition'), None): cnv_formula,
	((TABLENS,u'condition-source'), None): cnv_string,
	((TABLENS,u'condition-source-range-address'), None): cnv_string,
	((TABLENS,u'contains-error'), None): cnv_boolean,
	((TABLENS,u'contains-header'), None): cnv_boolean,
	((TABLENS,u'content-validation-name'), None): cnv_string,
	((TABLENS,u'copy-back'), None): cnv_boolean,
	((TABLENS,u'copy-formulas'), None): cnv_boolean,
	((TABLENS,u'copy-styles'), None): cnv_boolean,
	((TABLENS,u'count'), None): cnv_positiveInteger,
	((TABLENS,u'country'), None): cnv_token,
	((TABLENS,u'data-cell-range-address'), None): cnv_string,
	((TABLENS,u'data-field'), None): cnv_string,
	((TABLENS,u'data-type'), None): cnv_string,
	((TABLENS,u'database-name'), None): cnv_string,
	((TABLENS,u'database-table-name'), None): cnv_string,
	((TABLENS,u'date-end'), None): cnv_string,
	((TABLENS,u'date-start'), None): cnv_string,
	((TABLENS,u'date-value'), None): cnv_date,
	((TABLENS,u'default-cell-style-name'), None): cnv_StyleNameRef,
	((TABLENS,u'direction'), None): cnv_string,
	((TABLENS,u'display-border'), None): cnv_boolean,
	((TABLENS,u'display'), None): cnv_boolean,
	((TABLENS,u'display-duplicates'), None): cnv_boolean,
	((TABLENS,u'display-filter-buttons'), None): cnv_boolean,
	((TABLENS,u'display-list'), None): cnv_string,
	((TABLENS,u'display-member-mode'), None): cnv_string,
	((TABLENS,u'drill-down-on-double-click'), None): cnv_boolean,
	((TABLENS,u'embedded-number-behavior'), None): cnv_string,
	((TABLENS,u'enabled'), None): cnv_boolean,
	((TABLENS,u'end-cell-address'), None): cnv_string,
	((TABLENS,u'end'), None): cnv_string,
	((TABLENS,u'end-column'), None): cnv_integer,
	((TABLENS,u'end-position'), None): cnv_integer,
	((TABLENS,u'end-row'), None): cnv_integer,
	((TABLENS,u'end-table'), None): cnv_integer,
	((TABLENS,u'end-x'), None): cnv_length,
	((TABLENS,u'end-y'), None): cnv_length,
	((TABLENS,u'execute'), None): cnv_boolean,
	((TABLENS,u'expression'), None): cnv_formula,
	((TABLENS,u'field-name'), None): cnv_string,
	((TABLENS,u'field-number'), None): cnv_nonNegativeInteger,
	((TABLENS,u'field-number'), None): cnv_string,
	((TABLENS,u'filter-name'), None): cnv_string,
	((TABLENS,u'filter-options'), None): cnv_string,
	((TABLENS,u'first-row-end-column'), None): cnv_rowOrCol,
	((TABLENS,u'first-row-start-column'), None): cnv_rowOrCol,
	((TABLENS,u'formula'), None): cnv_formula,
	((TABLENS,u'function'), None): cnv_string,
	((TABLENS,u'function'), None): cnv_string,
	((TABLENS,u'grand-total'), None): cnv_string,
	((TABLENS,u'group-by-field-number'), None): cnv_nonNegativeInteger,
	((TABLENS,u'grouped-by'), None): cnv_string,
	((TABLENS,u'has-persistent-data'), None): cnv_boolean,
	((TABLENS,u'id'), None): cnv_string,
	((TABLENS,u'identify-categories'), None): cnv_boolean,
	((TABLENS,u'ignore-empty-rows'), None): cnv_boolean,
	((TABLENS,u'index'), None): cnv_nonNegativeInteger,
	((TABLENS,u'is-active'), None): cnv_boolean,
	((TABLENS,u'is-data-layout-field'), None): cnv_string,
	((TABLENS,u'is-selection'), None): cnv_boolean,
	((TABLENS,u'is-sub-table'), None): cnv_boolean,
	((TABLENS,u'label-cell-range-address'), None): cnv_string,
	((TABLENS,u'language'), None): cnv_token,
	((TABLENS,u'language'), None): cnv_token,
	((TABLENS,u'last-column-spanned'), None): cnv_positiveInteger,
	((TABLENS,u'last-row-end-column'), None): cnv_rowOrCol,
	((TABLENS,u'last-row-spanned'), None): cnv_positiveInteger,
	((TABLENS,u'last-row-start-column'), None): cnv_rowOrCol,
	((TABLENS,u'layout-mode'), None): cnv_string,
	((TABLENS,u'link-to-source-data'), None): cnv_boolean,
	((TABLENS,u'marked-invalid'), None): cnv_boolean,
	((TABLENS,u'matrix-covered'), None): cnv_boolean,
	((TABLENS,u'maximum-difference'), None): cnv_double,
	((TABLENS,u'member-count'), None): cnv_nonNegativeInteger,
	((TABLENS,u'member-name'), None): cnv_string,
	((TABLENS,u'member-type'), None): cnv_string,
	((TABLENS,u'message-type'), None): cnv_string,
	((TABLENS,u'mode'), None): cnv_string,
	((TABLENS,u'multi-deletion-spanned'), None): cnv_integer,
	((TABLENS,u'name'), None): cnv_string,
	((TABLENS,u'name'), None): cnv_string,
	((TABLENS,u'null-year'), None): cnv_positiveInteger,
	((TABLENS,u'number-columns-repeated'), None): cnv_positiveInteger,
	((TABLENS,u'number-columns-spanned'), None): cnv_positiveInteger,
	((TABLENS,u'number-matrix-columns-spanned'), None): cnv_positiveInteger,
	((TABLENS,u'number-matrix-rows-spanned'), None): cnv_positiveInteger,
	((TABLENS,u'number-rows-repeated'), None): cnv_positiveInteger,
	((TABLENS,u'number-rows-spanned'), None): cnv_positiveInteger,
	((TABLENS,u'object-name'), None): cnv_string,
	((TABLENS,u'on-update-keep-size'), None): cnv_boolean,
	((TABLENS,u'on-update-keep-styles'), None): cnv_boolean,
	((TABLENS,u'operator'), None): cnv_string,
	((TABLENS,u'operator'), None): cnv_string,
	((TABLENS,u'order'), None): cnv_string,
	((TABLENS,u'orientation'), None): cnv_string,
	((TABLENS,u'orientation'), None): cnv_string,
	((TABLENS,u'page-breaks-on-group-change'), None): cnv_boolean,
	((TABLENS,u'paragraph-style-name'), None): cnv_StyleNameRef,
	((TABLENS,u'parse-sql-statement'), None): cnv_boolean,
	((TABLENS,u'password'), None): cnv_string,
	((TABLENS,u'position'), None): cnv_integer,
	((TABLENS,u'precision-as-shown'), None): cnv_boolean,
	((TABLENS,u'print'), None): cnv_boolean,
	((TABLENS,u'print-ranges'), None): cnv_string,
	((TABLENS,u'protect'), None): cnv_boolean,
	((TABLENS,u'protected'), None): cnv_boolean,
	((TABLENS,u'protection-key'), None): cnv_string,
	((TABLENS,u'protection-key-digest-algorithm'), None): cnv_anyURI,
	((TABLENS,u'query-name'), None): cnv_string,
	((TABLENS,u'range-usable-as'), None): cnv_string,
	((TABLENS,u'rfc-language-tag'), None): cnv_language,
	((TABLENS,u'refresh-delay'), None): cnv_boolean,
	((TABLENS,u'refresh-delay'), None): cnv_duration,
	((TABLENS,u'rejecting-change-id'), None): cnv_string,
	((TABLENS,u'row'), None): cnv_integer,
	((TABLENS,u'scenario-ranges'), None): cnv_string,
	((TABLENS,u'script'), None): cnv_string,
	((TABLENS,u'search-criteria-must-apply-to-whole-cell'), None): cnv_boolean,
	((TABLENS,u'selected-page'), None): cnv_string,
	((TABLENS,u'show-details'), None): cnv_boolean,
	((TABLENS,u'show-empty'), None): cnv_boolean,
	((TABLENS,u'show-empty'), None): cnv_string,
	((TABLENS,u'show-filter-button'), None): cnv_boolean,
	((TABLENS,u'sort-mode'), None): cnv_string,
	((TABLENS,u'source-cell-range-addresses'), None): cnv_string,
	((TABLENS,u'source-cell-range-addresses'), None): cnv_string,
	((TABLENS,u'source-field-name'), None): cnv_string,
	((TABLENS,u'source-field-name'), None): cnv_string,
	((TABLENS,u'source-name'), None): cnv_string,
	((TABLENS,u'sql-statement'), None): cnv_string,
	((TABLENS,u'start'), None): cnv_string,
	((TABLENS,u'start-column'), None): cnv_integer,
	((TABLENS,u'start-position'), None): cnv_integer,
	((TABLENS,u'start-row'), None): cnv_integer,
	((TABLENS,u'start-table'), None): cnv_integer,
	((TABLENS,u'status'), None): cnv_string,
	((TABLENS,u'step'), None): cnv_double,
	((TABLENS,u'steps'), None): cnv_positiveInteger,
	((TABLENS,u'structure-protected'), None): cnv_boolean,
	((TABLENS,u'style-name'), None): cnv_StyleNameRef,
	((TABLENS,u'table-background'), None): cnv_boolean,
	((TABLENS,u'table'), None): cnv_integer,
	((TABLENS,u'table-name'), None): cnv_string,
	((TABLENS,u'target-cell-address'), None): cnv_string,
	((TABLENS,u'target-cell-address'), None): cnv_string,
	((TABLENS,u'target-range-address'), None): cnv_string,
	((TABLENS,u'target-range-address'), None): cnv_string,
	((TABLENS,u'template-name'), None): cnv_string,
	((TABLENS,u'title'), None): cnv_string,
	((TABLENS,u'track-changes'), None): cnv_boolean,
	((TABLENS,u'type'), None): cnv_string,
	((TABLENS,u'use-banding-columns-styles'), None): cnv_boolean,
	((TABLENS,u'use-banding-rows-styles'), None): cnv_boolean,
	((TABLENS,u'use-first-column-styles'), None): cnv_boolean,
	((TABLENS,u'use-first-row-styles'), None): cnv_boolean,
	((TABLENS,u'use-labels'), None): cnv_string,
	((TABLENS,u'use-last-column-styles'), None): cnv_boolean,
	((TABLENS,u'use-last-row-styles'), None): cnv_boolean,
	((TABLENS,u'use-regular-expressions'), None): cnv_boolean,
	((TABLENS,u'use-wildcards'), None): cnv_boolean,
	((TABLENS,u'used-hierarchy'), None): cnv_integer,
	((TABLENS,u'user-name'), None): cnv_string,
	((TABLENS,u'value'), None): cnv_string,
	((TABLENS,u'value'), None): cnv_string,
	((TABLENS,u'value-type'), None): cnv_string,
	((TABLENS,u'visibility'), None): cnv_string,
	((TEXTNS,u'active'), None): cnv_boolean,
	((TEXTNS,u'address'), None): cnv_string,
	((TEXTNS,u'alphabetical-separators'), None): cnv_boolean,
	((TEXTNS,u'anchor-page-number'), None): cnv_positiveInteger,
	((TEXTNS,u'anchor-type'), None): cnv_string,
	((TEXTNS,u'animation'), None): cnv_string,
	((TEXTNS,u'animation-delay'), None): cnv_string,
	((TEXTNS,u'animation-direction'), None): cnv_string,
	((TEXTNS,u'animation-repeat'), None): cnv_string,
	((TEXTNS,u'animation-start-inside'), None): cnv_boolean,
	((TEXTNS,u'animation-steps'), None): cnv_length,
	((TEXTNS,u'animation-stop-inside'), None): cnv_boolean,
	((TEXTNS,u'annote'), None): cnv_string,
	((TEXTNS,u'author'), None): cnv_string,
	((TEXTNS,u'bibliography-data-field'), None): cnv_string,
	((TEXTNS,u'bibliography-type'), None): cnv_string,
	((TEXTNS,u'booktitle'), None): cnv_string,
	((TEXTNS,u'bullet-char'), None): cnv_string,
	((TEXTNS,u'bullet-relative-size'), None): cnv_string,
	((TEXTNS,u'c'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'capitalize-entries'), None): cnv_boolean,
	((TEXTNS,u'caption-sequence-format'), None): cnv_string,
	((TEXTNS,u'caption-sequence-name'), None): cnv_string,
	((TEXTNS,u'change-id'), None): cnv_IDREF,
	((TEXTNS,u'chapter'), None): cnv_string,
	((TEXTNS,u'citation-body-style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'citation-style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'class-names'), None): cnv_NCNames,
	((TEXTNS,u'column-name'), None): cnv_string,
	((TEXTNS,u'combine-entries'), None): cnv_boolean,
	((TEXTNS,u'combine-entries-with-dash'), None): cnv_boolean,
	((TEXTNS,u'combine-entries-with-pp'), None): cnv_boolean,
	((TEXTNS,u'comma-separated'), None): cnv_boolean,
	((TEXTNS,u'cond-style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'condition'), None): cnv_formula,
	((TEXTNS,u'connection-name'), None): cnv_string,
	((TEXTNS,u'consecutive-numbering'), None): cnv_boolean,
	((TEXTNS,u'continue-list'), None): cnv_IDREF,
	((TEXTNS,u'continue-numbering'), None): cnv_boolean,
	((TEXTNS,u'copy-outline-levels'), None): cnv_boolean,
	((TEXTNS,u'count-empty-lines'), None): cnv_boolean,
	((TEXTNS,u'count-in-text-boxes'), None): cnv_boolean,
	((TEXTNS,u'current-value'), None): cnv_boolean,
	((TEXTNS,u'custom1'), None): cnv_string,
	((TEXTNS,u'custom2'), None): cnv_string,
	((TEXTNS,u'custom3'), None): cnv_string,
	((TEXTNS,u'custom4'), None): cnv_string,
	((TEXTNS,u'custom5'), None): cnv_string,
	((TEXTNS,u'database-name'), None): cnv_string,
	((TEXTNS,u'date-adjust'), None): cnv_duration,
	((TEXTNS,u'date-value'), None): cnv_date,
#	((TEXTNS,u'date-value'), None): cnv_dateTime,
	((TEXTNS,u'default-style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'description'), None): cnv_string,
	((TEXTNS,u'display'), None): cnv_string,
	((TEXTNS,u'display-levels'), None): cnv_positiveInteger,
	((TEXTNS,u'display-outline-level'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'dont-balance-text-columns'), None): cnv_boolean,
	((TEXTNS,u'duration'), None): cnv_duration,
	((TEXTNS,u'edition'), None): cnv_string,
	((TEXTNS,u'editor'), None): cnv_string,
	((TEXTNS,u'filter-name'), None): cnv_string,
	((TEXTNS,u'fixed'), None): cnv_boolean,
	((TEXTNS,u'footnotes-position'), None): cnv_string,
	((TEXTNS,u'formula'), None): cnv_formula,
	((TEXTNS,u'global'), None): cnv_boolean,
	((TEXTNS,u'howpublished'), None): cnv_string,
	((TEXTNS,u'id'), None): cnv_ID,
#	((TEXTNS,u'id'), None): cnv_string,
	((TEXTNS,u'identifier'), None): cnv_string,
	((TEXTNS,u'ignore-case'), None): cnv_boolean,
	((TEXTNS,u'increment'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'index-name'), None): cnv_string,
	((TEXTNS,u'index-scope'), None): cnv_string,
	((TEXTNS,u'institution'), None): cnv_string,
	((TEXTNS,u'is-hidden'), None): cnv_boolean,
	((TEXTNS,u'is-list-header'), None): cnv_boolean,
	((TEXTNS,u'isbn'), None): cnv_string,
	((TEXTNS,u'issn'), None): cnv_string,
	((TEXTNS,u'issn'), None): cnv_string,
	((TEXTNS,u'journal'), None): cnv_string,
	((TEXTNS,u'key'), None): cnv_string,
	((TEXTNS,u'key1'), None): cnv_string,
	((TEXTNS,u'key1-phonetic'), None): cnv_string,
	((TEXTNS,u'key2'), None): cnv_string,
	((TEXTNS,u'key2-phonetic'), None): cnv_string,
	((TEXTNS,u'kind'), None): cnv_string,
	((TEXTNS,u'label'), None): cnv_string,
	((TEXTNS,u'label-followed-by'), None): cnv_string,
	((TEXTNS,u'level'), None): cnv_positiveInteger,
	((TEXTNS,u'line-break'), None): cnv_boolean,
	((TEXTNS,u'line-number'), None): cnv_string,
	((TEXTNS,u'list-id'), None): cnv_NCName,
	((TEXTNS,u'list-level-position-and-space-mode'), None): cnv_string,
	((TEXTNS,u'list-tab-stop-position'), None): cnv_length,
	((TEXTNS,u'main-entry'), None): cnv_boolean,
	((TEXTNS,u'main-entry-style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'master-page-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'min-label-distance'), None): cnv_string,
	((TEXTNS,u'min-label-width'), None): cnv_string,
	((TEXTNS,u'month'), None): cnv_string,
	((TEXTNS,u'name'), None): cnv_string,
	((TEXTNS,u'note-class'), None): cnv_textnoteclass,
	((TEXTNS,u'note'), None): cnv_string,
	((TEXTNS,u'number'), None): cnv_string,
	((TEXTNS,u'number-lines'), None): cnv_boolean,
	((TEXTNS,u'number-position'), None): cnv_string,
	((TEXTNS,u'numbered-entries'), None): cnv_boolean,
	((TEXTNS,u'offset'), None): cnv_string,
	((TEXTNS,u'organizations'), None): cnv_string,
	((TEXTNS,u'outline-level'), None): cnv_string,
	((TEXTNS,u'page-adjust'), None): cnv_integer,
	((TEXTNS,u'pages'), None): cnv_string,
	((TEXTNS,u'placeholder-type'), None): cnv_string,
	((TEXTNS,u'prefix'), None): cnv_string,
	((TEXTNS,u'protected'), None): cnv_boolean,
	((TEXTNS,u'protection-key'), None): cnv_string,
	((TEXTNS,u'protection-key-digest-algorithm'), None): cnv_anyURI,
	((TEXTNS,u'publisher'), None): cnv_string,
	((TEXTNS,u'ref-name'), None): cnv_string,
	((TEXTNS,u'reference-format'), None): cnv_string,
	((TEXTNS,u'relative-tab-stop-position'), None): cnv_boolean,
	((TEXTNS,u'report-type'), None): cnv_string,
	((TEXTNS,u'restart-numbering'), None): cnv_boolean,
	((TEXTNS,u'restart-on-page'), None): cnv_boolean,
	((TEXTNS,u'row-number'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'school'), None): cnv_string,
	((TEXTNS,u'section-name'), None): cnv_string,
	((TEXTNS,u'select-page'), None): cnv_string,
	((TEXTNS,u'separation-character'), None): cnv_string,
	((TEXTNS,u'series'), None): cnv_string,
	((TEXTNS,u'sort-algorithm'), None): cnv_string,
	((TEXTNS,u'sort-ascending'), None): cnv_boolean,
	((TEXTNS,u'sort-by-position'), None): cnv_boolean,
	((TEXTNS,u'space-before'), None): cnv_string,
	((TEXTNS,u'start-numbering-at'), None): cnv_string,
	((TEXTNS,u'start-value'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'start-value'), None): cnv_positiveInteger,
	((TEXTNS,u'string-value'), None): cnv_string,
	((TEXTNS,u'string-value-if-false'), None): cnv_string,
	((TEXTNS,u'string-value-if-true'), None): cnv_string,
	((TEXTNS,u'string-value-phonetic'), None): cnv_string,
	((TEXTNS,u'style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'style-override'), None): cnv_StyleNameRef,
	((TEXTNS,u'suffix'), None): cnv_string,
	((TEXTNS,u'tab-ref'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'table-name'), None): cnv_string,
	((TEXTNS,u'table-type'), None): cnv_string,
	((TEXTNS,u'time-adjust'), None): cnv_duration,
	((TEXTNS,u'time-value'), None): cnv_dateTime,
	((TEXTNS,u'time-value'), None): cnv_time,
	((TEXTNS,u'title'), None): cnv_string,
	((TEXTNS,u'track-changes'), None): cnv_boolean,
	((TEXTNS,u'url'), None): cnv_string,
	((TEXTNS,u'use-caption'), None): cnv_boolean,
	((TEXTNS,u'use-chart-objects'), None): cnv_boolean,
	((TEXTNS,u'use-draw-objects'), None): cnv_boolean,
	((TEXTNS,u'use-floating-frames'), None): cnv_boolean,
	((TEXTNS,u'use-graphics'), None): cnv_boolean,
	((TEXTNS,u'use-index-marks'), None): cnv_boolean,
	((TEXTNS,u'use-index-source-styles'), None): cnv_boolean,
	((TEXTNS,u'use-keys-as-entries'), None): cnv_boolean,
	((TEXTNS,u'use-math-objects'), None): cnv_boolean,
	((TEXTNS,u'use-objects'), None): cnv_boolean,
	((TEXTNS,u'use-other-objects'), None): cnv_boolean,
	((TEXTNS,u'use-outline-level'), None): cnv_boolean,
	((TEXTNS,u'use-soft-page-breaks'), None): cnv_boolean,
	((TEXTNS,u'use-spreadsheet-objects'), None): cnv_boolean,
	((TEXTNS,u'use-tables'), None): cnv_boolean,
	((TEXTNS,u'value'), None): cnv_nonNegativeInteger,
	((TEXTNS,u'visited-style-name'), None): cnv_StyleNameRef,
	((TEXTNS,u'volume'), None): cnv_string,
	((TEXTNS,u'year'), None): cnv_string,
	((XFORMSNS,u'bind'), None): cnv_string,
	((XHTMLNS,u'about'), None): cnv_anyURI,
	((XHTMLNS,u'content'), None): cnv_string,
	((XHTMLNS,u'datatype'), None): cnv_anyURI,
	((XHTMLNS,u'property'), None): cnv_anyURI,
	((XLINKNS,u'actuate'), None): cnv_string,
	((XLINKNS,u'href'), None): cnv_anyURI,
	((XLINKNS,u'show'), None): cnv_xlinkshow,
	((XLINKNS,u'title'), None): cnv_string,
	((XLINKNS,u'type'), None): cnv_xlinktype,
	((XMLNS,u'id'), None): cnv_NCName,
}

class AttrConverters:
    def convert(self, attribute, value, element):
        """ Based on the element, figures out how to check/convert the attribute value
            All values are converted to string
        """
        conversion = attrconverters.get((attribute, element.qname), None)
        if conversion is not None:
            return conversion(attribute, value, element)
        else:
            conversion = attrconverters.get((attribute, None), None)
            if conversion is not None:
                return conversion(attribute, value, element)
        if sys.version_info[0]==2:
            return unicode(value)
        else:
            return str(value)

''',
    },
    'odf.chart': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#
from __future__ import absolute_import
from odf.namespaces import CHARTNS
from odf.element import Element

# Autogenerated
def Axis(**args):
    return Element(qname = (CHARTNS,'axis'), **args)

def Categories(**args):
    return Element(qname = (CHARTNS,'categories'), **args)

def Chart(**args):
    return Element(qname = (CHARTNS,'chart'), **args)

def DataLabel(**args):
    return Element(qname = (CHARTNS,'data-label'), **args)

def DataPoint(**args):
    return Element(qname = (CHARTNS,'data-point'), **args)

def Domain(**args):
    return Element(qname = (CHARTNS,'domain'), **args)

def Equation(**args):
    return Element(qname = (CHARTNS,'equation'), **args)

def ErrorIndicator(**args):
    return Element(qname = (CHARTNS,'error-indicator'), **args)

def Floor(**args):
    return Element(qname = (CHARTNS,'floor'), **args)

def Footer(**args):
    return Element(qname = (CHARTNS,'footer'), **args)

def Grid(**args):
    return Element(qname = (CHARTNS,'grid'), **args)

def LabelSeparator(**args):
    return Element(qname = (CHARTNS,'label-separator'), **args)

def Legend(**args):
    return Element(qname = (CHARTNS,'legend'), **args)

def MeanValue(**args):
    return Element(qname = (CHARTNS,'mean-value'), **args)

def PlotArea(**args):
    return Element(qname = (CHARTNS,'plot-area'), **args)

def RegressionCurve(**args):
    return Element(qname = (CHARTNS,'regression-curve'), **args)

def Series(**args):
    return Element(qname = (CHARTNS,'series'), **args)

def StockGainMarker(**args):
    return Element(qname = (CHARTNS,'stock-gain-marker'), **args)

def StockLossMarker(**args):
    return Element(qname = (CHARTNS,'stock-loss-marker'), **args)

def StockRangeLine(**args):
    return Element(qname = (CHARTNS,'stock-range-line'), **args)

def Subtitle(**args):
    return Element(qname = (CHARTNS,'subtitle'), **args)

def SymbolImage(**args):
    return Element(qname = (CHARTNS,'symbol-image'), **args)

def Title(**args):
    return Element(qname = (CHARTNS,'title'), **args)

def Wall(**args):
    return Element(qname = (CHARTNS,'wall'), **args)


""",
    },
    'odf.config': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import CONFIGNS
from odf.element import Element

# Autogenerated
def ConfigItem(**args):
    return Element(qname = (CONFIGNS, 'config-item'), **args)

def ConfigItemMapEntry(**args):
    return Element(qname = (CONFIGNS,'config-item-map-entry'), **args)

def ConfigItemMapIndexed(**args):
    return Element(qname = (CONFIGNS,'config-item-map-indexed'), **args)

def ConfigItemMapNamed(**args):
    return Element(qname = (CONFIGNS,'config-item-map-named'), **args)

def ConfigItemSet(**args):
    return Element(qname = (CONFIGNS, 'config-item-set'), **args)


""",
    },
    'odf.dc': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import DCNS
from odf.element import Element

# Autogenerated
def Creator(**args):
    return Element(qname = (DCNS,'creator'), **args)

def Date(**args):
    return Element(qname = (DCNS,'date'), **args)

def Description(**args):
    return Element(qname = (DCNS,'description'), **args)

def Language(**args):
    return Element(qname = (DCNS,'language'), **args)

def Subject(**args):
    return Element(qname = (DCNS,'subject'), **args)

def Title(**args):
    return Element(qname = (DCNS,'title'), **args)

# The following complete the Dublin Core elements, but there is no
# guarantee a compliant implementation of OpenDocument will preserve
# these elements

#def Contributor(**args):
#    return Element(qname = (DCNS,'contributor'), **args)

#def Coverage(**args):
#    return Element(qname = (DCNS,'coverage'), **args)

#def Format(**args):
#    return Element(qname = (DCNS,'format'), **args)

#def Identifier(**args):
#    return Element(qname = (DCNS,'identifier'), **args)

#def Publisher(**args):
#    return Element(qname = (DCNS,'publisher'), **args)

#def Relation(**args):
#    return Element(qname = (DCNS,'relation'), **args)

#def Rights(**args):
#    return Element(qname = (DCNS,'rights'), **args)

#def Source(**args):
#    return Element(qname = (DCNS,'source'), **args)

#def Type(**args):
#    return Element(qname = (DCNS,'type'), **args)

""",
    },
    'odf.dr3d': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#
import sys, os.path
sys.path.append(os.path.dirname(__file__))

from odf.namespaces import DR3DNS
from odf.element import Element
from odf.draw import StyleRefElement

# Autogenerated
def Cube(**args):
    return StyleRefElement(qname = (DR3DNS,'cube'), **args)

def Extrude(**args):
    return StyleRefElement(qname = (DR3DNS,'extrude'), **args)

def Light(Element):
    return StyleRefElement(qname = (DR3DNS,'light'), **args)

def Rotate(**args):
    return StyleRefElement(qname = (DR3DNS,'rotate'), **args)

def Scene(**args):
    return StyleRefElement(qname = (DR3DNS,'scene'), **args)

def Sphere(**args):
    return StyleRefElement(qname = (DR3DNS,'sphere'), **args)


""",
    },
    'odf.draw': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

import sys, os.path
sys.path.append(os.path.dirname(__file__))
from odf.namespaces import DRAWNS, STYLENS, PRESENTATIONNS
from odf.element import Element

def StyleRefElement(stylename=None, classnames=None, **args):
    qattrs = {}
    if stylename is not None:
        f = stylename.getAttrNS(STYLENS, 'family')
        if f == 'graphic':
            qattrs[(DRAWNS,u'style-name')]= stylename
        elif f == 'presentation':
            qattrs[(PRESENTATIONNS,u'style-name')]= stylename
        else:
            raise ValueError( "Style's family must be either 'graphic' or 'presentation'")
    if classnames is not None:
        f = classnames[0].getAttrNS(STYLENS, 'family')
        if f == 'graphic':
            qattrs[(DRAWNS,u'class-names')]= classnames
        elif f == 'presentation':
            qattrs[(PRESENTATIONNS,u'class-names')]= classnames
        else:
            raise ValueError( "Style's family must be either 'graphic' or 'presentation'")
    return Element(qattributes=qattrs, **args)

def DrawElement(name=None, **args):
    e = Element(name=name, **args)
    if 'displayname' not in args:
        e.setAttrNS(DRAWNS,'display-name', name)
    return e

# Autogenerated
def A(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (DRAWNS,'a'), **args)

def Applet(**args):
    return Element(qname = (DRAWNS,'applet'), **args)

def AreaCircle(**args):
    return Element(qname = (DRAWNS,'area-circle'), **args)

def AreaPolygon(**args):
    return Element(qname = (DRAWNS,'area-polygon'), **args)

def AreaRectangle(**args):
    return Element(qname = (DRAWNS,'area-rectangle'), **args)

def Caption(**args):
    return StyleRefElement(qname = (DRAWNS,'caption'), **args)

def Circle(**args):
    return StyleRefElement(qname = (DRAWNS,'circle'), **args)

def Connector(**args):
    return StyleRefElement(qname = (DRAWNS,'connector'), **args)

def ContourPath(**args):
    return Element(qname = (DRAWNS,'contour-path'), **args)

def ContourPolygon(**args):
    return Element(qname = (DRAWNS,'contour-polygon'), **args)

def Control(**args):
    return StyleRefElement(qname = (DRAWNS,'control'), **args)

def CustomShape(**args):
    return StyleRefElement(qname = (DRAWNS,'custom-shape'), **args)

def Ellipse(**args):
    return StyleRefElement(qname = (DRAWNS,'ellipse'), **args)

def EnhancedGeometry(**args):
    return Element(qname = (DRAWNS,'enhanced-geometry'), **args)

def Equation(**args):
    return Element(qname = (DRAWNS,'equation'), **args)

def FillImage(**args):
    args.setdefault('type', 'simple')
    return DrawElement(qname = (DRAWNS,'fill-image'), **args)

def FloatingFrame(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (DRAWNS,'floating-frame'), **args)

def Frame(**args):
    return StyleRefElement(qname = (DRAWNS,'frame'), **args)

def G(**args):
    return StyleRefElement(qname = (DRAWNS,'g'), **args)

def GluePoint(**args):
    return Element(qname = (DRAWNS,'glue-point'), **args)

def Gradient(**args):
    return DrawElement(qname = (DRAWNS,'gradient'), **args)

def Handle(**args):
    return Element(qname = (DRAWNS,'handle'), **args)

def Hatch(**args):
    return DrawElement(qname = (DRAWNS,'hatch'), **args)

def Image(**args):
    return Element(qname = (DRAWNS,'image'), **args)

def ImageMap(**args):
    return Element(qname = (DRAWNS,'image-map'), **args)

def Layer(**args):
    return Element(qname = (DRAWNS,'layer'), **args)

def LayerSet(**args):
    return Element(qname = (DRAWNS,'layer-set'), **args)

def Line(**args):
    return StyleRefElement(qname = (DRAWNS,'line'), **args)

def Marker(**args):
    return DrawElement(qname = (DRAWNS,'marker'), **args)

def Measure(**args):
    return StyleRefElement(qname = (DRAWNS,'measure'), **args)

def Object(**args):
    return Element(qname = (DRAWNS,'object'), **args)

def ObjectOle(**args):
    return Element(qname = (DRAWNS,'object-ole'), **args)

def Opacity(**args):
    return DrawElement(qname = (DRAWNS,'opacity'), **args)

def Page(**args):
    return Element(qname = (DRAWNS,'page'), **args)

def PageThumbnail(**args):
    return StyleRefElement(qname = (DRAWNS,'page-thumbnail'), **args)

def Param(**args):
    return Element(qname = (DRAWNS,'param'), **args)

def Path(**args):
    return StyleRefElement(qname = (DRAWNS,'path'), **args)

def Plugin(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (DRAWNS,'plugin'), **args)

def Polygon(**args):
    return StyleRefElement(qname = (DRAWNS,'polygon'), **args)

def Polyline(**args):
    return StyleRefElement(qname = (DRAWNS,'polyline'), **args)

def Rect(**args):
    return StyleRefElement(qname = (DRAWNS,'rect'), **args)

def RegularPolygon(**args):
    return StyleRefElement(qname = (DRAWNS,'regular-polygon'), **args)

def StrokeDash(**args):
    return DrawElement(qname = (DRAWNS,'stroke-dash'), **args)

def TextBox(**args):
    return Element(qname = (DRAWNS,'text-box'), **args)


""",
    },
    'odf.easyliststyle': {
        'is_package': False,
        'source': r'''
# -*- coding: utf-8 -*-
#   Create a <text:list-style> element from a text string.
#   Copyright (C) 2008 J. David Eisenberg
#
#   This program is free software; you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation; either version 2 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License along
#   with this program; if not, write to the Free Software Foundation, Inc.,
#   51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
#
# Contributor(s):
#

import re, sys, os.path
sys.path.append(os.path.dirname(__file__))
from odf.style import Style, TextProperties, ListLevelProperties
from odf.text import ListStyle,ListLevelStyleNumber,ListLevelStyleBullet

"""
Create a <text:list-style> element from a string or array.

List styles require a lot of code to create one level at a time.
These routines take a string and delimiter, or a list of
strings, and creates a <text:list-style> element for you.
Each item in the string (or array) represents a list level
 * style for levels 1-10.</p>
 *
 * <p>If an item contains <code>1</code>, <code>I</code>,
 * <code>i</code>, <code>A</code>, or <code>a</code>, then it is presumed
 * to be a numbering style; otherwise it is a bulleted style based on the
 * first character in the item.</p>
"""

_MAX_LIST_LEVEL = 10
SHOW_ALL_LEVELS = True
SHOW_ONE_LEVEL = False

def styleFromString(name, specifiers, delim, spacing, showAllLevels):
    specArray = specifiers.split(delim)
    return styleFromList( name, specArray, spacing, showAllLevels )

def styleFromList( styleName, specArray, spacing, showAllLevels):
    bullet = ""
    numPrefix = ""
    numSuffix = ""
    numberFormat = ""
    cssLengthNum = 0
    cssLengthUnits = ""
    numbered = False
    displayLevels = 0
    listStyle = ListStyle(name=styleName)
    numFormatPattern = re.compile("([1IiAa])")
    cssLengthPattern = re.compile("([^a-z]+)\\s*([a-z]+)?")
    m = cssLengthPattern.search( spacing )
    if (m != None):
        cssLengthNum = float(m.group(1))
        if (m.lastindex == 2):
            cssLengthUnits = m.group(2)
    i = 0
    while i < len(specArray):
        specification = specArray[i]
        m = numFormatPattern.search(specification)
        if (m != None):
            numberFormat = m.group(1)
            numPrefix = specification[0:m.start(1)]
            numSuffix = specification[m.end(1):]
            bullet = ""
            numbered = True
            if (showAllLevels):
                displayLevels = i + 1
            else:
                displayLevels = 1
        else:    # it's a bullet style
            bullet = specification
            numPrefix = ""
            numSuffix = ""
            numberFormat = ""
            displayLevels = 1
            numbered = False
        if (numbered):
            lls = ListLevelStyleNumber(level=(i+1))
            if (numPrefix != ''):
                lls.setAttribute('numprefix', numPrefix)
            if (numSuffix != ''):
                lls.setAttribute('numsuffix', numSuffix)
            lls.setAttribute('displaylevels', displayLevels)
        else:
            lls = ListLevelStyleBullet(level=(i+1),bulletchar=bullet[0])
        llp = ListLevelProperties()
        llp.setAttribute('spacebefore', str(cssLengthNum * (i+1)) + cssLengthUnits)
        llp.setAttribute('minlabelwidth', str(cssLengthNum) + cssLengthUnits)
        lls.addElement( llp )
        listStyle.addElement(lls)
        i += 1
    return listStyle

# vim: set expandtab sw=4 :

''',
    },
    'odf.element': {
        'is_package': False,
        'source': r'''
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright (C) 2007-2010 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

# Note: This script has copied a lot of text from xml.dom.minidom.
# Whatever license applies to that file also applies to this file.
#
import sys, os.path
sys.path.append(os.path.dirname(__file__))
import re
import xml.dom
from xml.dom.minicompat import *
from odf.namespaces import nsdict
import odf.grammar as grammar
from odf.attrconverters import AttrConverters

if sys.version_info[0] == 3:
    unicode=str # unicode function does not exist
    unichr=chr  # unichr does not exist

_xml11_illegal_ranges = (
    (0x0, 0x0,),
    (0xd800, 0xdfff,),
    (0xfffe, 0xffff,),
)

_xml10_illegal_ranges = _xml11_illegal_ranges + (
    (0x01, 0x08,),
    (0x0b, 0x0c,),
    (0x0e, 0x1f,),
)

_xml_discouraged_ranges = (
    (0x7f, 0x84,),
    (0x86, 0x9f,),
)

if sys.maxunicode >= 0x10000:
    # modern or "wide" python build
    _xml_discouraged_ranges = _xml_discouraged_ranges + (
        (0x1fffe, 0x1ffff,),
        (0x2fffe, 0x2ffff,),
        (0x3fffe, 0x3ffff,),
        (0x4fffe, 0x4ffff,),
        (0x5fffe, 0x5ffff,),
        (0x6fffe, 0x6ffff,),
        (0x7fffe, 0x7ffff,),
        (0x8fffe, 0x8ffff,),
        (0x9fffe, 0x9ffff,),
        (0xafffe, 0xaffff,),
        (0xbfffe, 0xbffff,),
        (0xcfffe, 0xcffff,),
        (0xdfffe, 0xdffff,),
        (0xefffe, 0xeffff,),
        (0xffffe, 0xfffff,),
        (0x10fffe, 0x10ffff,),
    )
# else "narrow" python build - only possible with old versions

def _range_seq_to_re(range_seq):
    # range pairs are specified as closed intervals
    return re.compile(u"[{}]".format(
        u"".join(
            u"{}-{}".format(re.escape(unichr(lo)), re.escape(unichr(hi)))
            for lo, hi in range_seq
        )
    ), flags=re.UNICODE)

_xml_filtered_chars_re = _range_seq_to_re(_xml10_illegal_ranges + _xml_discouraged_ranges)

def _handle_unrepresentable(data):
    return _xml_filtered_chars_re.sub(u"\ufffd", data)

# The following code is pasted form xml.sax.saxutils
# Tt makes it possible to run the code without the xml sax package installed
# To make it possible to have <rubbish> in your text elements, it is necessary to escape the texts
def _escape(data, entities={}):
    """ Escape &, <, and > in a string of data.

        You can escape other strings of data by passing a dictionary as
        the optional entities parameter.  The keys and values must all be
        strings; each key will be replaced with its corresponding value.
    """
    data = data.replace("&", "&amp;")
    data = data.replace("<", "&lt;")
    data = data.replace(">", "&gt;")
    for chars, entity in entities.items():
        data = data.replace(chars, entity)
    return data

def _sanitize(data, entities={}):
    return _escape(_handle_unrepresentable(data), entities=entities)

def _quoteattr(data, entities={}):
    """ Escape and quote an attribute value.

        Escape &, <, and > in a string of data, then quote it for use as
        an attribute value.  The \" character will be escaped as well, if
        necessary.

        You can escape other strings of data by passing a dictionary as
        the optional entities parameter.  The keys and values must all be
        strings; each key will be replaced with its corresponding value.
    """
    entities['\n']='&#10;'
    entities['\r']='&#12;'
    data = _sanitize(data, entities)
    if '"' in data:
        if "'" in data:
            data = '"%s"' % data.replace('"', "&quot;")
        else:
            data = "'%s'" % data
    else:
        data = '"%s"' % data
    return data

def _nssplit(qualifiedName):
    """ Split a qualified name into namespace part and local part.  """
    fields = qualifiedName.split(':', 1)
    if len(fields) == 2:
        return fields
    else:
        return (None, fields[0])

def _nsassign(namespace):
    return nsdict.setdefault(namespace,"ns" + str(len(nsdict)))


# Exceptions
class IllegalChild(Exception):
    """ Complains if you add an element to a parent where it is not allowed """
class IllegalText(Exception):
    """ Complains if you add text or cdata to an element where it is not allowed """

class Node(xml.dom.Node):
    """ super class for more specific nodes """
    parentNode = None
    nextSibling = None
    previousSibling = None

    def hasChildNodes(self):
        """ Tells whether this element has any children; text nodes,
            subelements, whatever.
        """
        if self.childNodes:
            return True
        else:
            return False

    def _get_childNodes(self):
        return self.childNodes

    def _get_firstChild(self):
        if self.childNodes:
            return self.childNodes[0]

    def _get_lastChild(self):
        if self.childNodes:
            return self.childNodes[-1]

    def insertBefore(self, newChild, refChild):
        """ Inserts the node newChild before the existing child node refChild.
            If refChild is null, insert newChild at the end of the list of children.
        """
        if newChild.nodeType not in self._child_node_types:
            raise IllegalChild( "%s cannot be child of %s" % (newChild.tagName, self.tagName))
        if newChild.parentNode is not None:
            newChild.parentNode.removeChild(newChild)
        if refChild is None:
            self.appendChild(newChild)
        else:
            try:
                index = self.childNodes.index(refChild)
            except ValueError:
                raise xml.dom.NotFoundErr()
            self.childNodes.insert(index, newChild)
            newChild.nextSibling = refChild
            refChild.previousSibling = newChild
            if index:
                node = self.childNodes[index-1]
                node.nextSibling = newChild
                newChild.previousSibling = node
            else:
                newChild.previousSibling = None
            newChild.parentNode = self
        return newChild

    def appendChild(self, newChild):
        """ Adds the node newChild to the end of the list of children of this node.
            If the newChild is already in the tree, it is first removed.
        """
        if newChild.nodeType == self.DOCUMENT_FRAGMENT_NODE:
            for c in tuple(newChild.childNodes):
                self.appendChild(c)
            ### The DOM does not clearly specify what to return in this case
            return newChild
        if newChild.nodeType not in self._child_node_types:
            raise IllegalChild( "<%s> is not allowed in %s" % ( newChild.tagName, self.tagName))
        if newChild.parentNode is not None:
            newChild.parentNode.removeChild(newChild)
        _append_child(self, newChild)
        newChild.nextSibling = None
        return newChild

    def removeChild(self, oldChild):
        """ Removes the child node indicated by oldChild from the list of children, and returns it.
        """
        #FIXME: update ownerDocument.element_dict or find other solution
        try:
            self.childNodes.remove(oldChild)
        except ValueError:
            raise xml.dom.NotFoundErr()
        if oldChild.nextSibling is not None:
            oldChild.nextSibling.previousSibling = oldChild.previousSibling
        if oldChild.previousSibling is not None:
            oldChild.previousSibling.nextSibling = oldChild.nextSibling
        oldChild.nextSibling = oldChild.previousSibling = None
        if self.ownerDocument:
            self.ownerDocument.remove_from_caches(oldChild)
        oldChild.parentNode = None
        return oldChild

    def __str__(self):
        val = []
        for c in self.childNodes:
            val.append(str(c))
        return ''.join(val)

    def __unicode__(self):
        val = []
        for c in self.childNodes:
            val.append(unicode(c))
        return u''.join(val)

defproperty(Node, "firstChild", doc="First child node, or None.")
defproperty(Node, "lastChild",  doc="Last child node, or None.")

def _append_child(self, node):
    # fast path with less checks; usable by DOM builders if careful
    childNodes = self.childNodes
    if childNodes:
        last = childNodes[-1]
        node.__dict__["previousSibling"] = last
        last.__dict__["nextSibling"] = node
    childNodes.append(node)
    node.__dict__["parentNode"] = self

class Childless:
    """ Mixin that makes childless-ness easy to implement and avoids
        the complexity of the Node methods that deal with children.
    """

    attributes = None
    childNodes = EmptyNodeList()
    firstChild = None
    lastChild = None

    def _get_firstChild(self):
        return None

    def _get_lastChild(self):
        return None

    def appendChild(self, node):
        """ Raises an error """
        raise xml.dom.HierarchyRequestErr(
            self.tagName + " nodes cannot have children")

    def hasChildNodes(self):
        return False

    def insertBefore(self, newChild, refChild):
        """ Raises an error """
        raise xml.dom.HierarchyRequestErr(
            self.tagName + " nodes do not have children")

    def removeChild(self, oldChild):
        """ Raises an error """
        raise xml.dom.NotFoundErr(
            self.tagName + " nodes do not have children")

    def replaceChild(self, newChild, oldChild):
        """ Raises an error """
        raise xml.dom.HierarchyRequestErr(
            self.tagName + " nodes do not have children")

class Text(Childless, Node):
    nodeType = Node.TEXT_NODE
    tagName = "Text"

    def __init__(self, data):
        self.data = data

    def __str__(self):
        return self.data

    def __unicode__(self):
        return self.data

    def toXml(self,level,f):
        """ Write XML in UTF-8 """
        if self.data:
            f.write(_sanitize(unicode(self.data)))
    
class CDATASection(Text, Childless):
    nodeType = Node.CDATA_SECTION_NODE

    def toXml(self,level,f):
        """ Generate XML output of the node. If the text contains "]]>", then
            escape it by going out of CDATA mode (]]>), then write the string
            and then go into CDATA mode again. (<![CDATA[)
        """
        if self.data:
            f.write('<![CDATA[%s]]>' % self.data.replace(']]>',']]>]]><![CDATA['))

class Element(Node):
    """ Creates a arbitrary element and is intended to be subclassed not used on its own.
        This element is the base of every element it defines a class which resembles
        a xml-element. The main advantage of this kind of implementation is that you don't
        have to create a toXML method for every different object. Every element
        consists of an attribute, optional subelements, optional text and optional cdata.
    """

    nodeType = Node.ELEMENT_NODE
    namespaces = {}  # Due to shallow copy this is a static variable

    _child_node_types = (Node.ELEMENT_NODE,
                         Node.PROCESSING_INSTRUCTION_NODE,
                         Node.COMMENT_NODE,
                         Node.TEXT_NODE,
                         Node.CDATA_SECTION_NODE,
                         Node.ENTITY_REFERENCE_NODE)

    def __init__(self, attributes=None, text=None, cdata=None, qname=None, qattributes=None, check_grammar=True, **args):
        if qname is not None:
            self.qname = qname
        assert(hasattr(self, 'qname'))
        self.ownerDocument = None
        self.childNodes=[]
        self.allowed_children = grammar.allowed_children.get(self.qname)
        prefix = self.get_nsprefix(self.qname[0])
        self.tagName = prefix + ":" + self.qname[1]
        if text is not None:
            self.addText(text)
        if cdata is not None:
            self.addCDATA(cdata)

        allowed_attrs = self.allowed_attributes()
        if allowed_attrs is not None:
            allowed_args = [ a[1].lower().replace('-','') for a in allowed_attrs]
        self.attributes={}
        # Load the attributes from the 'attributes' argument
        if attributes:
            for attr, value in attributes.items():
                self.setAttribute(attr, value)
        # Load the qualified attributes
        if qattributes:
            for attr, value in qattributes.items():
                self.setAttrNS(attr[0], attr[1], value)
        if allowed_attrs is not None:
            # Load the attributes from the 'args' argument
            for arg in args.keys():
                self.setAttribute(arg, args[arg])
        else:
            for arg in args.keys():  # If any attribute is allowed
                self.attributes[arg]=args[arg]
        if not check_grammar:
            return
        # Test that all mandatory attributes have been added.
        required = grammar.required_attributes.get(self.qname)
        if required:
            for r in required:
                if self.getAttrNS(r[0],r[1]) is None:
                    raise AttributeError( "Required attribute missing: %s in <%s>" % (r[1].lower().replace('-',''), self.tagName))

    def get_knownns(self, prefix):
        """ Odfpy maintains a list of known namespaces. In some cases a prefix is used, and
            we need to know which namespace it resolves to.
        """
        global nsdict
        for ns,p in nsdict.items():
            if p == prefix: return ns
        return None

    def get_nsprefix(self, namespace):
        """ Odfpy maintains a list of known namespaces. In some cases we have a namespace URL,
            and needs to look up or assign the prefix for it.
        """
        if namespace is None: namespace = ""
        prefix = _nsassign(namespace)
        if not namespace in self.namespaces:
            self.namespaces[namespace] = prefix
        return prefix

    def allowed_attributes(self):
        return grammar.allowed_attributes.get(self.qname)

    def _setOwnerDoc(self, element):
        element.ownerDocument = self.ownerDocument
        for child in element.childNodes:
            self._setOwnerDoc(child)

    def addElement(self, element, check_grammar=True):
        """ adds an element to an Element

            Element.addElement(Element)
        """
        if check_grammar and self.allowed_children is not None:
            if element.qname not in self.allowed_children:
                raise IllegalChild( "<%s> is not allowed in <%s>" % ( element.tagName, self.tagName))
        self.appendChild(element)
        self._setOwnerDoc(element)
        if self.ownerDocument:
            self.ownerDocument.rebuild_caches(element)

    def addText(self, text, check_grammar=True):
        """ Adds text to an element
            Setting check_grammar=False turns off grammar checking
        """
        if check_grammar and self.qname not in grammar.allows_text:
            raise IllegalText( "The <%s> element does not allow text" % self.tagName)
        else:
            if text != '':
                self.appendChild(Text(text))

    def addCDATA(self, cdata, check_grammar=True):
        """ Adds CDATA to an element
            Setting check_grammar=False turns off grammar checking
        """
        if check_grammar and self.qname not in grammar.allows_text:
            raise IllegalText( "The <%s> element does not allow text" % self.tagName)
        else:
            self.appendChild(CDATASection(cdata))

    def removeAttribute(self, attr, check_grammar=True):
        """ Removes an attribute by name. """
        allowed_attrs = self.allowed_attributes()
        if allowed_attrs is None:
            if type(attr) == type(()):
                prefix, localname = attr
                self.removeAttrNS(prefix, localname)
            else:
                raise AttributeError( "Unable to add simple attribute - use (namespace, localpart)")
        else:
            # Construct a list of allowed arguments
            allowed_args = [ a[1].lower().replace('-','') for a in allowed_attrs]
            if check_grammar and attr not in allowed_args:
                raise AttributeError( "Attribute %s is not allowed in <%s>" % ( attr, self.tagName))
            i = allowed_args.index(attr)
            self.removeAttrNS(allowed_attrs[i][0], allowed_attrs[i][1])

    def setAttribute(self, attr, value, check_grammar=True):
        """ Add an attribute to the element
            This is sort of a convenience method. All attributes in ODF have
            namespaces. The library knows what attributes are legal and then allows
            the user to provide the attribute as a keyword argument and the
            library will add the correct namespace.
            Must overwrite, If attribute already exists.
        """
        if attr == 'parent' and value is not None:
            value.addElement(self)
        else:
            allowed_attrs = self.allowed_attributes()
            if allowed_attrs is None:
                if type(attr) == type(()):
                    prefix, localname = attr
                    self.setAttrNS(prefix, localname, value)
                else:
                    raise AttributeError( "Unable to add simple attribute - use (namespace, localpart)")
            else:
                # Construct a list of allowed arguments
                allowed_args = [ a[1].lower().replace('-','') for a in allowed_attrs]
                if check_grammar and attr not in allowed_args:
                    raise AttributeError( "Attribute %s is not allowed in <%s>" % ( attr, self.tagName))
                i = allowed_args.index(attr)
                self.setAttrNS(allowed_attrs[i][0], allowed_attrs[i][1], value)

    def setAttrNS(self, namespace, localpart, value):
        """ Add an attribute to the element
            In case you need to add an attribute the library doesn't know about
            then you must provide the full qualified name
            It will not check that the attribute is legal according to the schema.
            Must overwrite, If attribute already exists.
        """
        allowed_attrs = self.allowed_attributes()
        prefix = self.get_nsprefix(namespace)
#       if allowed_attrs and (namespace, localpart) not in allowed_attrs:
#           raise AttributeError( "Attribute %s:%s is not allowed in element <%s>" % ( prefix, localpart, self.tagName))
        c = AttrConverters()
        self.attributes[(namespace, localpart)] = c.convert((namespace, localpart), value, self)

    def getAttrNS(self, namespace, localpart):
        """
        gets an attribute, given a namespace and a key
        @param namespace a unicode string or a bytes: the namespace
        @param localpart a unicode string or a bytes:
        the key to get the attribute
        @return an attribute as a unicode string or a bytes: if both paramters
        are byte strings, it will be a bytes; if both attributes are
        unicode strings, it will be a unicode string
        """
        prefix = self.get_nsprefix(namespace)
        result = self.attributes.get((namespace, localpart))

        assert(
            (type(namespace), type(namespace), type(namespace) == \
                 type(b""), type(b""), type(b"")) or
            (type(namespace), type(namespace), type(namespace) == \
                 type(u""), type(u""), type(u""))
            )

        return result

    def removeAttrNS(self, namespace, localpart):
        del self.attributes[(namespace, localpart)]

    def getAttribute(self, attr):
        """ Get an attribute value. The method knows which namespace the attribute is in
        """
        allowed_attrs = self.allowed_attributes()
        if allowed_attrs is None:
            if type(attr) == type(()):
                prefix, localname = attr
                return self.getAttrNS(prefix, localname)
            else:
                raise AttributeError( "Unable to get simple attribute - use (namespace, localpart)")
        else:
            # Construct a list of allowed arguments
            allowed_args = [ a[1].lower().replace('-','') for a in allowed_attrs]
            i = allowed_args.index(attr)
            return self.getAttrNS(allowed_attrs[i][0], allowed_attrs[i][1])

    def write_open_tag(self, level, f):
        f.write(('<'+self.tagName))
        if level == 0:
            for namespace, prefix in self.namespaces.items():
                f.write(u' xmlns:' + prefix + u'="'+ _sanitize(str(namespace))+'"')
        for qname in self.attributes.keys():
            prefix = self.get_nsprefix(qname[0])
            f.write(u' '+_sanitize(str(prefix+u':'+qname[1]))+u'='+_quoteattr(unicode(self.attributes[qname])))
        f.write(u'>')

    def write_close_tag(self, level, f):
        f.write('</'+self.tagName+'>')

    def toXml(self, level, f):
        """
        Generate an XML stream out of the tree structure
        @param level integer: level in the XML tree; zero at root of the tree
        @param f an open writable file able to accept unicode strings
        """
        f.write(u'<'+self.tagName)
        if level == 0:
            for namespace, prefix in self.namespaces.items():
                f.write(u' xmlns:' + prefix + u'="'+ _sanitize(str(namespace))+u'"')
        for qname in self.attributes.keys():
            prefix = self.get_nsprefix(qname[0])
            f.write(u' '+_sanitize(unicode(prefix+':'+qname[1]))+u'='+_quoteattr(unicode(self.attributes[qname])))
        if self.childNodes:
            f.write(u'>')
            for element in self.childNodes:
                element.toXml(level+1,f)
            f.write(u'</'+self.tagName+'>')
        else:
            f.write(u'/>')

    def _getElementsByObj(self, obj, accumulator):
        if self.qname == obj.qname:
            accumulator.append(self)
        for e in self.childNodes:
            if e.nodeType == Node.ELEMENT_NODE:
                accumulator = e._getElementsByObj(obj, accumulator)
        return accumulator

    def getElementsByType(self, element):
        """ Gets elements based on the type, which is function from text.py, draw.py etc. """
        obj = element(check_grammar=False)
        return self._getElementsByObj(obj,[])

    def isInstanceOf(self, element):
        """ This is a check to see if the object is an instance of a type """
        obj = element(check_grammar=False)
        return self.qname == obj.qname

''',
    },
    'odf.elementtypes': {
        'is_package': False,
        'source': r"""
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright (C) 2008 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import *

# Inline element don't cause a box
# They are analogous to the HTML elements SPAN, B, I etc.
inline_elements = (
    (TEXTNS,u'a'),
    (TEXTNS,u'author-initials'),
    (TEXTNS,u'author-name'),
    (TEXTNS,u'bibliography-mark'),
    (TEXTNS,u'bookmark-ref'),
    (TEXTNS,u'chapter'),
    (TEXTNS,u'character-count'),
    (TEXTNS,u'conditional-text'),
    (TEXTNS,u'creation-date'),
    (TEXTNS,u'creation-time'),
    (TEXTNS,u'creator'),
    (TEXTNS,u'database-display'),
    (TEXTNS,u'database-name'),
    (TEXTNS,u'database-next'),
    (TEXTNS,u'database-row-number'),
    (TEXTNS,u'database-row-select'),
    (TEXTNS,u'date'),
    (TEXTNS,u'dde-connection'),
    (TEXTNS,u'description'),
    (TEXTNS,u'editing-cycles'),
    (TEXTNS,u'editing-duration'),
    (TEXTNS,u'execute-macro'),
    (TEXTNS,u'expression'),
    (TEXTNS,u'file-name'),
    (TEXTNS,u'hidden-paragraph'),
    (TEXTNS,u'hidden-text'),
    (TEXTNS,u'image-count'),
    (TEXTNS,u'initial-creator'),
    (TEXTNS,u'keywords'),
    (TEXTNS,u'measure'),
    (TEXTNS,u'modification-date'),
    (TEXTNS,u'modification-time'),
    (TEXTNS,u'note-ref'),
    (TEXTNS,u'object-count'),
    (TEXTNS,u'page-continuation'),
    (TEXTNS,u'page-count'),
    (TEXTNS,u'page-number'),
    (TEXTNS,u'page-variable-get'),
    (TEXTNS,u'page-variable-set'),
    (TEXTNS,u'paragraph-count'),
    (TEXTNS,u'placeholder'),
    (TEXTNS,u'print-date'),
    (TEXTNS,u'printed-by'),
    (TEXTNS,u'print-time'),
    (TEXTNS,u'reference-ref'),
    (TEXTNS,u'ruby'),
    (TEXTNS,u'ruby-base'),
    (TEXTNS,u'ruby-text'),
    (TEXTNS,u'script'),
    (TEXTNS,u'sender-city'),
    (TEXTNS,u'sender-company'),
    (TEXTNS,u'sender-country'),
    (TEXTNS,u'sender-email'),
    (TEXTNS,u'sender-fax'),
    (TEXTNS,u'sender-firstname'),
    (TEXTNS,u'sender-initials'),
    (TEXTNS,u'sender-lastname'),
    (TEXTNS,u'sender-phone-private'),
    (TEXTNS,u'sender-phone-work'),
    (TEXTNS,u'sender-position'),
    (TEXTNS,u'sender-postal-code'),
    (TEXTNS,u'sender-state-or-province'),
    (TEXTNS,u'sender-street'),
    (TEXTNS,u'sender-title'),
    (TEXTNS,u'sequence'),
    (TEXTNS,u'sequence-ref'),
    (TEXTNS,u'sheet-name'),
    (TEXTNS,u'span'),
    (TEXTNS,u'subject'),
    (TEXTNS,u'table-count'),
    (TEXTNS,u'table-formula'),
    (TEXTNS,u'template-name'),
    (TEXTNS,u'text-input'),
    (TEXTNS,u'time'),
    (TEXTNS,u'title'),
    (TEXTNS,u'user-defined'),
    (TEXTNS,u'user-field-get'),
    (TEXTNS,u'user-field-input'),
    (TEXTNS,u'variable-get'),
    (TEXTNS,u'variable-input'),
    (TEXTNS,u'variable-set'),
    (TEXTNS,u'word-count'),
)


# It is almost impossible to determine what elements are block elements.
# There are so many that don't fit the form
block_elements = (
    (TEXTNS,u'h'),
    (TEXTNS,u'p'),
    (TEXTNS,u'list'),
    (TEXTNS,u'list-item'),
    (TEXTNS,u'section'),
)

declarative_elements = (
    (OFFICENS,u'font-face-decls'),
    (PRESENTATIONNS,u'date-time-decl'),
    (PRESENTATIONNS,u'footer-decl'),
    (PRESENTATIONNS,u'header-decl'),
    (TABLENS,u'table-template'),
    (TEXTNS,u'alphabetical-index-entry-template'),
    (TEXTNS,u'alphabetical-index-source'),
    (TEXTNS,u'bibliography-entry-template'),
    (TEXTNS,u'bibliography-source'),
    (TEXTNS,u'dde-connection-decls'),
    (TEXTNS,u'illustration-index-entry-template'),
    (TEXTNS,u'illustration-index-source'),
    (TEXTNS,u'index-source-styles'),
    (TEXTNS,u'index-title-template'),
    (TEXTNS,u'note-continuation-notice-backward'),
    (TEXTNS,u'note-continuation-notice-forward'),
    (TEXTNS,u'notes-configuration'),
    (TEXTNS,u'object-index-entry-template'),
    (TEXTNS,u'object-index-source'),
    (TEXTNS,u'sequence-decls'),
    (TEXTNS,u'table-index-entry-template'),
    (TEXTNS,u'table-index-source'),
    (TEXTNS,u'table-of-content-entry-template'),
    (TEXTNS,u'table-of-content-source'),
    (TEXTNS,u'user-field-decls'),
    (TEXTNS,u'user-index-entry-template'),
    (TEXTNS,u'user-index-source'),
    (TEXTNS,u'variable-decls'),
)

empty_elements = (
    (ANIMNS,u'animate'),
    (ANIMNS,u'animateColor'),
    (ANIMNS,u'animateMotion'),
    (ANIMNS,u'animateTransform'),
    (ANIMNS,u'audio'),
    (ANIMNS,u'param'),
    (ANIMNS,u'set'),
    (ANIMNS,u'transitionFilter'),
    (CHARTNS,u'categories'),
    (CHARTNS,u'data-point'),
    (CHARTNS,u'domain'),
    (CHARTNS,u'error-indicator'),
    (CHARTNS,u'floor'),
    (CHARTNS,u'grid'),
    (CHARTNS,u'legend'),
    (CHARTNS,u'mean-value'),
    (CHARTNS,u'regression-curve'),
    (CHARTNS,u'stock-gain-marker'),
    (CHARTNS,u'stock-loss-marker'),
    (CHARTNS,u'stock-range-line'),
    (CHARTNS,u'symbol-image'),
    (CHARTNS,u'wall'),
    (DR3DNS,u'cube'),
    (DR3DNS,u'extrude'),
    (DR3DNS,u'light'),
    (DR3DNS,u'rotate'),
    (DR3DNS,u'sphere'),
    (DRAWNS,u'contour-path'),
    (DRAWNS,u'contour-polygon'),
    (DRAWNS,u'equation'),
    (DRAWNS,u'fill-image'),
    (DRAWNS,u'floating-frame'),
    (DRAWNS,u'glue-point'),
    (DRAWNS,u'gradient'),
    (DRAWNS,u'handle'),
    (DRAWNS,u'hatch'),
    (DRAWNS,u'layer'),
    (DRAWNS,u'marker'),
    (DRAWNS,u'opacity'),
    (DRAWNS,u'page-thumbnail'),
    (DRAWNS,u'param'),
    (DRAWNS,u'stroke-dash'),
    (FORMNS,u'connection-resource'),
    (FORMNS,u'list-value'),
    (FORMNS,u'property'),
    (MANIFESTNS,u'algorithm'),
    (MANIFESTNS,u'key-derivation'),
    (METANS,u'auto-reload'),
    (METANS,u'document-statistic'),
    (METANS,u'hyperlink-behaviour'),
    (METANS,u'template'),
    (NUMBERNS,u'am-pm'),
    (NUMBERNS,u'boolean'),
    (NUMBERNS,u'day'),
    (NUMBERNS,u'day-of-week'),
    (NUMBERNS,u'era'),
    (NUMBERNS,u'fraction'),
    (NUMBERNS,u'hours'),
    (NUMBERNS,u'minutes'),
    (NUMBERNS,u'month'),
    (NUMBERNS,u'quarter'),
    (NUMBERNS,u'scientific-number'),
    (NUMBERNS,u'seconds'),
    (NUMBERNS,u'text-content'),
    (NUMBERNS,u'week-of-year'),
    (NUMBERNS,u'year'),
    (OFFICENS,u'dde-source'),
    (PRESENTATIONNS,u'date-time'),
    (PRESENTATIONNS,u'footer'),
    (PRESENTATIONNS,u'header'),
    (PRESENTATIONNS,u'placeholder'),
    (PRESENTATIONNS,u'play'),
    (PRESENTATIONNS,u'show'),
    (PRESENTATIONNS,u'sound'),
    (SCRIPTNS,u'event-listener'),
    (STYLENS,u'column'),
    (STYLENS,u'column-sep'),
    (STYLENS,u'drop-cap'),
    (STYLENS,u'footnote-sep'),
    (STYLENS,u'list-level-properties'),
    (STYLENS,u'map'),
    (STYLENS,u'ruby-properties'),
    (STYLENS,u'table-column-properties'),
    (STYLENS,u'tab-stop'),
    (STYLENS,u'text-properties'),
    (SVGNS,u'definition-src'),
    (SVGNS,u'font-face-format'),
    (SVGNS,u'font-face-name'),
    (SVGNS,u'stop'),
    (TABLENS,u'body'),
    (TABLENS,u'cell-address'),
    (TABLENS,u'cell-range-source'),
    (TABLENS,u'change-deletion'),
    (TABLENS,u'consolidation'),
    (TABLENS,u'database-source-query'),
    (TABLENS,u'database-source-sql'),
    (TABLENS,u'database-source-table'),
    (TABLENS,u'data-pilot-display-info'),
    (TABLENS,u'data-pilot-field-reference'),
    (TABLENS,u'data-pilot-group-member'),
    (TABLENS,u'data-pilot-layout-info'),
    (TABLENS,u'data-pilot-member'),
    (TABLENS,u'data-pilot-sort-info'),
    (TABLENS,u'data-pilot-subtotal'),
    (TABLENS,u'dependency'),
    (TABLENS,u'error-macro'),
    (TABLENS,u'even-columns'),
    (TABLENS,u'even-rows'),
    (TABLENS,u'filter-condition'),
    (TABLENS,u'first-column'),
    (TABLENS,u'first-row'),
    (TABLENS,u'highlighted-range'),
    (TABLENS,u'insertion-cut-off'),
    (TABLENS,u'iteration'),
    (TABLENS,u'label-range'),
    (TABLENS,u'last-column'),
    (TABLENS,u'last-row'),
    (TABLENS,u'movement-cut-off'),
    (TABLENS,u'named-expression'),
    (TABLENS,u'named-range'),
    (TABLENS,u'null-date'),
    (TABLENS,u'odd-columns'),
    (TABLENS,u'odd-rows'),
    (TABLENS,u'operation'),
    (TABLENS,u'scenario'),
    (TABLENS,u'sort-by'),
    (TABLENS,u'sort-groups'),
    (TABLENS,u'source-range-address'),
    (TABLENS,u'source-service'),
    (TABLENS,u'subtotal-field'),
    (TABLENS,u'table-column'),
    (TABLENS,u'table-source'),
    (TABLENS,u'target-range-address'),
    (TEXTNS,u'alphabetical-index-auto-mark-file'),
    (TEXTNS,u'alphabetical-index-mark'),
    (TEXTNS,u'alphabetical-index-mark-end'),
    (TEXTNS,u'alphabetical-index-mark-start'),
    (TEXTNS,u'bookmark'),
    (TEXTNS,u'bookmark-end'),
    (TEXTNS,u'bookmark-start'),
    (TEXTNS,u'change'),
    (TEXTNS,u'change-end'),
    (TEXTNS,u'change-start'),
    (TEXTNS,u'dde-connection-decl'),
    (TEXTNS,u'index-entry-bibliography'),
    (TEXTNS,u'index-entry-chapter'),
    (TEXTNS,u'index-entry-link-end'),
    (TEXTNS,u'index-entry-link-start'),
    (TEXTNS,u'index-entry-page-number'),
    (TEXTNS,u'index-entry-tab-stop'),
    (TEXTNS,u'index-entry-text'),
    (TEXTNS,u'index-source-style'),
    (TEXTNS,u'line-break'),
    (TEXTNS,u'page'),
    (TEXTNS,u'reference-mark'),
    (TEXTNS,u'reference-mark-end'),
    (TEXTNS,u'reference-mark-start'),
    (TEXTNS,u's'),
    (TEXTNS,u'section-source'),
    (TEXTNS,u'sequence-decl'),
    (TEXTNS,u'soft-page-break'),
    (TEXTNS,u'sort-key'),
    (TEXTNS,u'tab'),
    (TEXTNS,u'toc-mark'),
    (TEXTNS,u'toc-mark-end'),
    (TEXTNS,u'toc-mark-start'),
    (TEXTNS,u'user-field-decl'),
    (TEXTNS,u'user-index-mark'),
    (TEXTNS,u'user-index-mark-end'),
    (TEXTNS,u'user-index-mark-start'),
    (TEXTNS,u'variable-decl')
)

""",
    },
    'odf.form': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import FORMNS
from odf.element import Element


# Autogenerated
def Button(**args):
    return Element(qname = (FORMNS,'button'), **args)

def Checkbox(**args):
    return Element(qname = (FORMNS,'checkbox'), **args)

def Column(**args):
    return Element(qname = (FORMNS,'column'), **args)

def Combobox(**args):
    return Element(qname = (FORMNS,'combobox'), **args)

def ConnectionResource(**args):
    return Element(qname = (FORMNS,'connection-resource'), **args)

def Date(**args):
    return Element(qname = (FORMNS,'date'), **args)

def File(**args):
    return Element(qname = (FORMNS,'file'), **args)

def FixedText(**args):
    return Element(qname = (FORMNS,'fixed-text'), **args)

def Form(**args):
    return Element(qname = (FORMNS,'form'), **args)

def FormattedText(**args):
    return Element(qname = (FORMNS,'formatted-text'), **args)

def Frame(**args):
    return Element(qname = (FORMNS,'frame'), **args)

def GenericControl(**args):
    return Element(qname = (FORMNS,'generic-control'), **args)

def Grid(**args):
    return Element(qname = (FORMNS,'grid'), **args)

def Hidden(**args):
    return Element(qname = (FORMNS,'hidden'), **args)

def Image(**args):
    return Element(qname = (FORMNS,'image'), **args)

def ImageFrame(**args):
    return Element(qname = (FORMNS,'image-frame'), **args)

def Item(**args):
    return Element(qname = (FORMNS,'item'), **args)

def ListProperty(**args):
    return Element(qname = (FORMNS,'list-property'), **args)

def ListValue(**args):
    return Element(qname = (FORMNS,'list-value'), **args)

def Listbox(**args):
    return Element(qname = (FORMNS,'listbox'), **args)

def Number(**args):
    return Element(qname = (FORMNS,'number'), **args)

def Option(**args):
    return Element(qname = (FORMNS,'option'), **args)

def Password(**args):
    return Element(qname = (FORMNS,'password'), **args)

def Properties(**args):
    return Element(qname = (FORMNS,'properties'), **args)

def Property(**args):
    return Element(qname = (FORMNS,'property'), **args)

def Radio(**args):
    return Element(qname = (FORMNS,'radio'), **args)

def Text(**args):
    return Element(qname = (FORMNS,'text'), **args)

def Textarea(**args):
    return Element(qname = (FORMNS,'textarea'), **args)

def Time(**args):
    return Element(qname = (FORMNS,'time'), **args)

def ValueRange(**args):
    return Element(qname = (FORMNS,'value-range'), **args)


""",
    },
    'odf.grammar': {
        'is_package': False,
        'source': r'''
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

__doc__=""" In principle the OpenDocument schema converted to python structures.
Currently it contains the legal child elements of a given element.
To be used for validation check in the API
"""
import sys, os.path
sys.path.append(os.path.dirname(__file__))
from odf.namespaces import *

# The following code is generated from the RelaxNG schema with this notice:

#       Open Document Format for Office Applications (OpenDocument) Version 1.2
#       OASIS Standard, 29 September 2011
#       Relax-NG Schema
#       Source: http://docs.oasis-open.org/office/v1.2/os/
#       Copyright (c) OASIS Open 2002-2011. All Rights Reserved.

#       All capitalized terms in the following text have the meanings assigned to them
#       in the OASIS Intellectual Property Rights Policy (the "OASIS IPR Policy"). The
#       full Policy may be found at the OASIS website.

#       This document and translations of it may be copied and furnished to others, and
#       derivative works that comment on or otherwise explain it or assist in its
#       implementation may be prepared, copied, published, and distributed, in whole or
#       in part, without restriction of any kind, provided that the above copyright
#       notice and this section are included on all such copies and derivative works.
#       However, this document itself may not be modified in any way, including by
#       removing the copyright notice or references to OASIS, except as needed for the
#       purpose of developing any document or deliverable produced by an OASIS
#       Technical Committee (in which case the rules applicable to copyrights, as set
#       forth in the OASIS IPR Policy, must be followed) or as required to translate it
#       into languages other than English.

#       The limited permissions granted above are perpetual and will not be revoked by
#       OASIS or its successors or assigns.

#       This document and the information contained herein is provided on an "AS IS"
#       basis and OASIS DISCLAIMS ALL WARRANTIES, EXPRESS OR IMPLIED, INCLUDING BUT NOT
#       LIMITED TO ANY WARRANTY THAT THE USE OF THE INFORMATION HEREIN WILL NOT
#       INFRINGE ANY OWNERSHIP RIGHTS OR ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR
#       FITNESS FOR A PARTICULAR PURPOSE.

allowed_children = {
	(DCNS,u'creator') : (
	),
	(DCNS,u'date') : (
	),
	(DCNS,u'description') : (
	),
	(DCNS,u'language') : (
	),
	(DCNS,u'subject') : (
	),
	(DCNS,u'title') : (
	),
# Completes Dublin Core start
#	(DCNS,'contributor') : (
#	),
#	(DCNS,'coverage') : (
#	),
#	(DCNS,'format') : (
#	),
#	(DCNS,'identifier') : (
#	),
#	(DCNS,'publisher') : (
#	),
#	(DCNS,'relation') : (
#	),
#	(DCNS,'rights') : (
#	),
#	(DCNS,'source') : (
#	),
#	(DCNS,'type') : (
#	),
# Completes Dublin Core end
	(MATHNS,u'math') : None,

	(XFORMSNS,u'model') : None,

	(ANIMNS,u'animate') : (
	),
	(ANIMNS,u'animateColor') : (
	),
	(ANIMNS,u'animateMotion') : (
	),
	(ANIMNS,u'animateTransform') : (
	),
	(ANIMNS,u'audio') : (
	),
	(ANIMNS,u'command') : (
		(ANIMNS,u'param'),
	),
# allowed_children
	(ANIMNS,u'iterate') : (
		(ANIMNS,u'animate'),
		(ANIMNS,u'animateColor'),
		(ANIMNS,u'animateMotion'),
		(ANIMNS,u'animateTransform'),
		(ANIMNS,u'audio'),
		(ANIMNS,u'command'),
		(ANIMNS,u'iterate'),
		(ANIMNS,u'par'),
		(ANIMNS,u'seq'),
		(ANIMNS,u'set'),
		(ANIMNS,u'transitionFilter'),
	),
	(ANIMNS,u'par') : (
		(ANIMNS,u'animate'),
		(ANIMNS,u'animateColor'),
		(ANIMNS,u'animateMotion'),
		(ANIMNS,u'animateTransform'),
		(ANIMNS,u'audio'),
		(ANIMNS,u'command'),
		(ANIMNS,u'iterate'),
		(ANIMNS,u'par'),
		(ANIMNS,u'seq'),
		(ANIMNS,u'set'),
		(ANIMNS,u'transitionFilter'),
	),
# allowed_children
	(ANIMNS,u'param') : (
	),
	(ANIMNS,u'seq') : (
		(ANIMNS,u'animate'),
		(ANIMNS,u'animateColor'),
		(ANIMNS,u'animateMotion'),
		(ANIMNS,u'animateTransform'),
		(ANIMNS,u'audio'),
		(ANIMNS,u'command'),
		(ANIMNS,u'iterate'),
		(ANIMNS,u'par'),
		(ANIMNS,u'seq'),
		(ANIMNS,u'set'),
		(ANIMNS,u'transitionFilter'),
	),
	(ANIMNS,u'set') : (
	),
	(ANIMNS,u'transitionFilter') : (
	),
	(CHARTNS,u'axis') : (
		(CHARTNS,u'categories'),
		(CHARTNS,u'grid'),
		(CHARTNS,u'title'),
	),
# allowed_children
	(CHARTNS,u'categories') : (
	),
	(CHARTNS,u'chart') : (
		(CHARTNS,u'footer'),
		(CHARTNS,u'legend'),
		(CHARTNS,u'plot-area'),
		(CHARTNS,u'subtitle'),
		(CHARTNS,u'title'),
		(TABLENS,u'table'),
	),
	(CHARTNS,u'data-label') : (
		(TEXTNS,u'p'),
	),
	(CHARTNS,u'data-point') : (
		(CHARTNS,u'data-label'),
	),
	(CHARTNS,u'domain') : (
	),
# allowed_children
	(CHARTNS,u'equation') : (
		(TEXTNS,u'p'),
	),
	(CHARTNS,u'error-indicator') : (
	),
	(CHARTNS,u'floor') : (
	),
	(CHARTNS,u'footer') : (
		(TEXTNS,u'p'),
	),
	(CHARTNS,u'grid') : (
	),
	(CHARTNS,u'label-separator') : (
		(TEXTNS,u'p'),
	),
	(CHARTNS,u'legend') : (
		(TEXTNS,u'p'),
	),
# allowed_children
	(CHARTNS,u'mean-value') : (
	),
	(CHARTNS,u'plot-area') : (
		(CHARTNS,u'axis'),
		(CHARTNS,u'floor'),
		(CHARTNS,u'series'),
		(CHARTNS,u'stock-gain-marker'),
		(CHARTNS,u'stock-loss-marker'),
		(CHARTNS,u'stock-range-line'),
		(CHARTNS,u'wall'),
		(DR3DNS,u'light'),
	),
	(CHARTNS,u'regression-curve') : (
		(CHARTNS,u'equation'),
	),
	(CHARTNS,u'series') : (
		(CHARTNS,u'data-label'),
		(CHARTNS,u'data-point'),
		(CHARTNS,u'domain'),
		(CHARTNS,u'error-indicator'),
		(CHARTNS,u'mean-value'),
		(CHARTNS,u'regression-curve'),
	),
	(CHARTNS,u'stock-gain-marker') : (
	),
	(CHARTNS,u'stock-loss-marker') : (
	),
# allowed_children
	(CHARTNS,u'stock-range-line') : (
	),
	(CHARTNS,u'subtitle') : (
		(TEXTNS,u'p'),
	),
	(CHARTNS,u'symbol-image') : (
	),
	(CHARTNS,u'title') : (
		(TEXTNS,u'p'),
	),
	(CHARTNS,u'wall') : (
	),
	(CONFIGNS,u'config-item') : (
	),
	(CONFIGNS,u'config-item-map-entry') : (
		(CONFIGNS,u'config-item'),
		(CONFIGNS,u'config-item-map-indexed'),
		(CONFIGNS,u'config-item-map-named'),
		(CONFIGNS,u'config-item-set'),
	),
	(CONFIGNS,u'config-item-map-indexed') : (
		(CONFIGNS,u'config-item-map-entry'),
	),
	(CONFIGNS,u'config-item-map-named') : (
		(CONFIGNS,u'config-item-map-entry'),
	),
# allowed_children
	(CONFIGNS,u'config-item-set') : (
		(CONFIGNS,u'config-item'),
		(CONFIGNS,u'config-item-map-indexed'),
		(CONFIGNS,u'config-item-map-named'),
		(CONFIGNS,u'config-item-set'),
	),
	(MANIFESTNS,u'algorithm') : (
	),
	(MANIFESTNS,u'encryption-data') : (
		(MANIFESTNS,u'algorithm'),
		(MANIFESTNS,u'key-derivation'),
	),
	(MANIFESTNS,u'file-entry') : (
		(MANIFESTNS,u'encryption-data'),
	),
	(MANIFESTNS,u'key-derivation') : (
	),
	(MANIFESTNS,u'manifest') : (
		(MANIFESTNS,u'file-entry'),
	),
	(NUMBERNS,u'am-pm') : (
	),
	(NUMBERNS,u'boolean') : (
	),
# allowed_children
	(NUMBERNS,u'boolean-style') : (
		(NUMBERNS,u'boolean'),
		(NUMBERNS,u'text'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
	(NUMBERNS,u'currency-style') : (
		(NUMBERNS,u'currency-symbol'),
		(NUMBERNS,u'number'),
		(NUMBERNS,u'text'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
	(NUMBERNS,u'currency-symbol') : (
	),
	(NUMBERNS,u'date-style') : (
		(NUMBERNS,u'am-pm'),
		(NUMBERNS,u'day'),
		(NUMBERNS,u'day-of-week'),
		(NUMBERNS,u'era'),
		(NUMBERNS,u'hours'),
		(NUMBERNS,u'minutes'),
		(NUMBERNS,u'month'),
		(NUMBERNS,u'quarter'),
		(NUMBERNS,u'seconds'),
		(NUMBERNS,u'text'),
		(NUMBERNS,u'week-of-year'),
		(NUMBERNS,u'year'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
# allowed_children
	(NUMBERNS,u'day') : (
	),
	(NUMBERNS,u'day-of-week') : (
	),
	(NUMBERNS,u'embedded-text') : (
	),
	(NUMBERNS,u'era') : (
	),
	(NUMBERNS,u'fraction') : (
	),
	(NUMBERNS,u'hours') : (
	),
	(NUMBERNS,u'minutes') : (
	),
	(NUMBERNS,u'month') : (
	),
	(NUMBERNS,u'number') : (
		(NUMBERNS,u'embedded-text'),
	),
	(NUMBERNS,u'number-style') : (
		(NUMBERNS,u'fraction'),
		(NUMBERNS,u'number'),
		(NUMBERNS,u'scientific-number'),
		(NUMBERNS,u'text'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
# allowed_children
	(NUMBERNS,u'percentage-style') : (
		(NUMBERNS,u'number'),
		(NUMBERNS,u'text'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
	(NUMBERNS,u'quarter') : (
	),
	(NUMBERNS,u'scientific-number') : (
	),
	(NUMBERNS,u'seconds') : (
	),
	(NUMBERNS,u'text') : (
	),
	(NUMBERNS,u'text-content') : (
	),
	(NUMBERNS,u'text-style') : (
		(NUMBERNS,u'text'),
		(NUMBERNS,u'text-content'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
# allowed_children
	(NUMBERNS,u'time-style') : (
		(NUMBERNS,u'am-pm'),
		(NUMBERNS,u'hours'),
		(NUMBERNS,u'minutes'),
		(NUMBERNS,u'seconds'),
		(NUMBERNS,u'text'),
		(STYLENS,u'map'),
		(STYLENS,u'text-properties'),
	),
# allowed_children
	(NUMBERNS,u'week-of-year') : (
	),
	(NUMBERNS,u'year') : (
	),
	(DR3DNS,u'cube') : (
	),
	(DR3DNS,u'extrude') : (
	),
	(DR3DNS,u'light') : (
	),
	(DR3DNS,u'rotate') : (
	),
	(DR3DNS,u'scene') : (
		(DR3DNS,u'cube'),
		(DR3DNS,u'extrude'),
		(DR3DNS,u'light'),
		(DR3DNS,u'rotate'),
		(DR3DNS,u'scene'),
		(DR3DNS,u'sphere'),
		(DRAWNS,u'glue-point'),
		(SVGNS,u'title'),
		(SVGNS,u'desc'),
	),
	(DR3DNS,u'sphere') : (
	),
	(DRAWNS,u'a') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
	),
# allowed_children
	(DRAWNS,u'applet') : (
		(DRAWNS,u'param'),
	),
	(DRAWNS,u'area-circle') : (
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
	(DRAWNS,u'area-polygon') : (
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
	(DRAWNS,u'area-rectangle') : (
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
	(DRAWNS,u'caption') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'circle') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
# allowed_children
	(DRAWNS,u'connector') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'contour-path') : (
	),
	(DRAWNS,u'contour-polygon') : (
	),
	(DRAWNS,u'control') : (
		(DRAWNS,u'glue-point'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
	(DRAWNS,u'custom-shape') : (
		(DRAWNS,u'enhanced-geometry'),
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
# allowed_children
	(DRAWNS,u'ellipse') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'enhanced-geometry') : (
		(DRAWNS,u'equation'),
		(DRAWNS,u'handle'),
	),
	(DRAWNS,u'equation') : (
	),
# allowed_children
	(DRAWNS,u'fill-image') : (
	),
	(DRAWNS,u'floating-frame') : (
	),
	(DRAWNS,u'frame') : (
		(DRAWNS,u'applet'),
		(DRAWNS,u'contour-path'),
		(DRAWNS,u'contour-polygon'),
		(DRAWNS,u'floating-frame'),
		(DRAWNS,u'glue-point'),
		(DRAWNS,u'image'),
		(DRAWNS,u'image-map'),
		(DRAWNS,u'object'),
		(DRAWNS,u'object-ole'),
		(DRAWNS,u'plugin'),
		(DRAWNS,u'text-box'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TABLENS,u'table'),
	),
# allowed_children
	(DRAWNS,u'g') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'glue-point'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(OFFICENS,u'event-listeners'),
	),
	(DRAWNS,u'glue-point') : (
	),
	(DRAWNS,u'gradient') : (
	),
	(DRAWNS,u'handle') : (
	),
	(DRAWNS,u'hatch') : (
	),
# allowed_children
	(DRAWNS,u'image') : (
		(OFFICENS,u'binary-data'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'image-map') : (
		(DRAWNS,u'area-circle'),
		(DRAWNS,u'area-polygon'),
		(DRAWNS,u'area-rectangle'),
	),
	(DRAWNS,u'layer') : (
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
	(DRAWNS,u'layer-set') : (
		(DRAWNS,u'layer'),
	),
	(DRAWNS,u'line') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'marker') : (
	),
	(DRAWNS,u'measure') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
		(SVGNS,u'title'),
		(SVGNS,u'desc'),
	),
	(DRAWNS,u'object') : (
		(MATHNS,u'math'),
		(OFFICENS,u'document'),
	),
# allowed_children
	(DRAWNS,u'object-ole') : (
		(OFFICENS,u'binary-data'),
	),
	(DRAWNS,u'opacity') : (
	),
	(DRAWNS,u'page') : (
		(ANIMNS,u'animate'),
		(ANIMNS,u'animateColor'),
		(ANIMNS,u'animateMotion'),
		(ANIMNS,u'animateTransform'),
		(ANIMNS,u'audio'),
		(ANIMNS,u'command'),
		(ANIMNS,u'iterate'),
		(ANIMNS,u'par'),
		(ANIMNS,u'seq'),
		(ANIMNS,u'set'),
		(ANIMNS,u'transitionFilter'),
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'layer-set'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'forms'),
		(PRESENTATIONNS,u'animations'),
		(PRESENTATIONNS,u'notes'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
# allowed_children
	(DRAWNS,u'page-thumbnail') : (
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
	),
	(DRAWNS,u'param') : (
	),
	(DRAWNS,u'path') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'plugin') : (
		(DRAWNS,u'param'),
	),
	(DRAWNS,u'polygon') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'title'),
		(SVGNS,u'desc'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'polyline') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
# allowed_children
	(DRAWNS,u'rect') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'regular-polygon') : (
		(DRAWNS,u'glue-point'),
		(OFFICENS,u'event-listeners'),
		(SVGNS,u'desc'),
		(SVGNS,u'title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
	(DRAWNS,u'stroke-dash') : (
	),
	(DRAWNS,u'text-box') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
# allowed_children
	(FORMNS,u'button') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'checkbox') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'column') : (
		(FORMNS,u'checkbox'),
		(FORMNS,u'combobox'),
		(FORMNS,u'date'),
		(FORMNS,u'formatted-text'),
		(FORMNS,u'listbox'),
		(FORMNS,u'number'),
		(FORMNS,u'text'),
		(FORMNS,u'textarea'),
		(FORMNS,u'time'),
	),
	(FORMNS,u'combobox') : (
		(FORMNS,u'item'),
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'connection-resource') : (
	),
	(FORMNS,u'date') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'file') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'fixed-text') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
# allowed_children
	(FORMNS,u'form') : (
		(FORMNS,u'button'),
		(FORMNS,u'checkbox'),
		(FORMNS,u'combobox'),
		(FORMNS,u'connection-resource'),
		(FORMNS,u'date'),
		(FORMNS,u'file'),
		(FORMNS,u'fixed-text'),
		(FORMNS,u'form'),
		(FORMNS,u'formatted-text'),
		(FORMNS,u'frame'),
		(FORMNS,u'generic-control'),
		(FORMNS,u'grid'),
		(FORMNS,u'hidden'),
		(FORMNS,u'image'),
		(FORMNS,u'image-frame'),
		(FORMNS,u'listbox'),
		(FORMNS,u'number'),
		(FORMNS,u'password'),
		(FORMNS,u'properties'),
		(FORMNS,u'radio'),
		(FORMNS,u'text'),
		(FORMNS,u'textarea'),
		(FORMNS,u'time'),
		(FORMNS,u'value-range'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'formatted-text') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'frame') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'generic-control') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'grid') : (
		(FORMNS,u'column'),
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'hidden') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'image') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
# allowed_children
	(FORMNS,u'image-frame') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'item') : (
	),
	(FORMNS,u'list-property') : (
		(FORMNS,u'list-value'),
		(FORMNS,u'list-value'),
		(FORMNS,u'list-value'),
		(FORMNS,u'list-value'),
		(FORMNS,u'list-value'),
		(FORMNS,u'list-value'),
		(FORMNS,u'list-value'),
	),
	(FORMNS,u'list-value') : (
	),
	(FORMNS,u'listbox') : (
		(FORMNS,u'option'),
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
# allowed_children
	(FORMNS,u'number') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'option') : (
	),
	(FORMNS,u'password') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'properties') : (
		(FORMNS,u'list-property'),
		(FORMNS,u'property'),
	),
	(FORMNS,u'property') : (
	),
	(FORMNS,u'radio') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'text') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'textarea') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
		(TEXTNS,u'p'),
	),
# allowed_children
	(FORMNS,u'time') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(FORMNS,u'value-range') : (
		(FORMNS,u'properties'),
		(OFFICENS,u'event-listeners'),
	),
	(METANS,u'auto-reload') : (
	),
	(METANS,u'creation-date') : (
	),
	(METANS,u'date-string') : (
	),
	(METANS,u'document-statistic') : (
	),
	(METANS,u'editing-cycles') : (
	),
	(METANS,u'editing-duration') : (
	),
	(METANS,u'generator') : (
	),
	(METANS,u'hyperlink-behaviour') : (
	),
	(METANS,u'initial-creator') : (
	),
	(METANS,u'keyword') : (
	),
	(METANS,u'print-date') : (
	),
	(METANS,u'printed-by') : (
	),
	(METANS,u'template') : (
	),
	(METANS,u'user-defined') : (
	),
# allowed_children
	(OFFICENS,u'annotation') : (
		(DCNS,u'creator'),
		(DCNS,u'date'),
		(METANS,u'date-string'),
		(TEXTNS,u'list'),
		(TEXTNS,u'p'),
	),
# allowed_children
	(OFFICENS,u'annotation-end') : (
	),
	(OFFICENS,u'automatic-styles') : (
		(NUMBERNS,u'boolean-style'),
		(NUMBERNS,u'currency-style'),
		(NUMBERNS,u'date-style'),
		(NUMBERNS,u'number-style'),
		(NUMBERNS,u'percentage-style'),
		(NUMBERNS,u'text-style'),
		(NUMBERNS,u'time-style'),
		(STYLENS,u'page-layout'),
		(STYLENS,u'style'),
		(TEXTNS,u'list-style'),
	),
	(OFFICENS,u'binary-data') : (
	),
	(OFFICENS,u'body') : (
		(OFFICENS,u'chart'),
		(OFFICENS,u'drawing'),
		(OFFICENS,u'image'),
		(OFFICENS,u'presentation'),
		(OFFICENS,u'spreadsheet'),
		(OFFICENS,u'text'),
	),
	(OFFICENS,u'change-info') : (
		(DCNS,u'creator'),
		(DCNS,u'date'),
		(TEXTNS,u'p'),
	),
	(OFFICENS,u'chart') : (
		(CHARTNS,u'chart'),
		(TABLENS,u'calculation-settings'),
		(TABLENS,u'consolidation'),
		(TABLENS,u'content-validations'),
		(TABLENS,u'data-pilot-tables'),
		(TABLENS,u'database-ranges'),
		(TABLENS,u'dde-links'),
		(TABLENS,u'label-ranges'),
		(TABLENS,u'named-expressions'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'variable-decls'),
	),
	(OFFICENS,u'dde-source') : (
	),
	(OFFICENS,u'document') : (
		(OFFICENS,u'automatic-styles'),
		(OFFICENS,u'body'),
		(OFFICENS,u'font-face-decls'),
		(OFFICENS,u'master-styles'),
		(OFFICENS,u'meta'),
		(OFFICENS,u'scripts'),
		(OFFICENS,u'settings'),
		(OFFICENS,u'styles'),
	),
	(OFFICENS,u'document-content') : (
		(OFFICENS,u'automatic-styles'),
		(OFFICENS,u'body'),
		(OFFICENS,u'font-face-decls'),
		(OFFICENS,u'scripts'),
	),
	(OFFICENS,u'document-meta') : (
		(OFFICENS,u'meta'),
	),
	(OFFICENS,u'document-settings') : (
		(OFFICENS,u'settings'),
	),
	(OFFICENS,u'document-styles') : (
		(OFFICENS,u'automatic-styles'),
		(OFFICENS,u'font-face-decls'),
		(OFFICENS,u'master-styles'),
		(OFFICENS,u'styles'),
	),
	(OFFICENS,u'drawing') : (
		(DRAWNS,u'page'),
		(TABLENS,u'calculation-settings'),
		(TABLENS,u'consolidation'),
		(TABLENS,u'content-validations'),
		(TABLENS,u'data-pilot-tables'),
		(TABLENS,u'database-ranges'),
		(TABLENS,u'dde-links'),
		(TABLENS,u'label-ranges'),
		(TABLENS,u'named-expressions'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'variable-decls'),
	),
	(OFFICENS,u'event-listeners') : (
		(PRESENTATIONNS,u'event-listener'),
		(SCRIPTNS,u'event-listener'),
	),
	(OFFICENS,u'font-face-decls') : (
		(STYLENS,u'font-face'),
	),
# allowed_children
	(OFFICENS,u'forms') : (
		(XFORMSNS,u'model'),
		(FORMNS,u'form'),
	),
	(OFFICENS,u'image') : (
		(DRAWNS,u'frame'),
	),
	(OFFICENS,u'master-styles') : (
		(DRAWNS,u'layer-set'),
		(STYLENS,u'handout-master'),
		(STYLENS,u'master-page'),
	),
	(OFFICENS,u'meta') : (
		(DCNS,u'creator'),
		(DCNS,u'date'),
		(DCNS,u'description'),
		(DCNS,u'language'),
		(DCNS,u'subject'),
		(DCNS,u'title'),
# Completes Dublin Core start
#		(DCNS,'contributor'),
#		(DCNS,'coverage'),
#		(DCNS,'format'),
#		(DCNS,'identifier'),
#		(DCNS,'publisher'),
#		(DCNS,'relation'),
#		(DCNS,'rights'),
#		(DCNS,'source'),
#		(DCNS,'type'),
# Completes Dublin Core end
		(METANS,u'auto-reload'),
		(METANS,u'creation-date'),
		(METANS,u'document-statistic'),
		(METANS,u'editing-cycles'),
		(METANS,u'editing-duration'),
		(METANS,u'generator'),
		(METANS,u'hyperlink-behaviour'),
		(METANS,u'initial-creator'),
		(METANS,u'keyword'),
		(METANS,u'print-date'),
		(METANS,u'printed-by'),
		(METANS,u'template'),
		(METANS,u'user-defined'),
	),
	(OFFICENS,u'presentation') : (
		(DRAWNS,u'page'),
		(PRESENTATIONNS,u'date-time-decl'),
		(PRESENTATIONNS,u'footer-decl'),
		(PRESENTATIONNS,u'header-decl'),
		(PRESENTATIONNS,u'settings'),
		(TABLENS,u'calculation-settings'),
		(TABLENS,u'consolidation'),
		(TABLENS,u'content-validations'),
		(TABLENS,u'data-pilot-tables'),
		(TABLENS,u'database-ranges'),
		(TABLENS,u'dde-links'),
		(TABLENS,u'label-ranges'),
		(TABLENS,u'named-expressions'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'variable-decls'),
	),
# allowed_children
	(OFFICENS,u'script') : None,

	(OFFICENS,u'scripts') : (
		(OFFICENS,u'event-listeners'),
		(OFFICENS,u'script'),
	),
	(OFFICENS,u'settings') : (
		(CONFIGNS,u'config-item-set'),
	),
	(OFFICENS,u'spreadsheet') : (
		(TABLENS,u'calculation-settings'),
		(TABLENS,u'consolidation'),
		(TABLENS,u'content-validations'),
		(TABLENS,u'data-pilot-tables'),
		(TABLENS,u'database-ranges'),
		(TABLENS,u'dde-links'),
		(TABLENS,u'label-ranges'),
		(TABLENS,u'named-expressions'),
		(TABLENS,u'table'),
		(TABLENS,u'tracked-changes'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'variable-decls'),
	),
	(OFFICENS,u'styles') : (
		(NUMBERNS,u'boolean-style'),
		(NUMBERNS,u'currency-style'),
		(NUMBERNS,u'date-style'),
		(NUMBERNS,u'number-style'),
		(NUMBERNS,u'percentage-style'),
		(NUMBERNS,u'text-style'),
		(NUMBERNS,u'time-style'),
		(DRAWNS,u'fill-image'),
		(DRAWNS,u'gradient'),
		(DRAWNS,u'hatch'),
		(DRAWNS,u'marker'),
		(DRAWNS,u'opacity'),
		(DRAWNS,u'stroke-dash'),
		(STYLENS,u'default-page-layout'),
		(STYLENS,u'default-style'),
		(STYLENS,u'presentation-page-layout'),
		(STYLENS,u'style'),
		(SVGNS,u'linearGradient'),
		(SVGNS,u'radialGradient'),
		(TABLENS,u'table-template'),
		(TEXTNS,u'bibliography-configuration'),
		(TEXTNS,u'linenumbering-configuration'),
		(TEXTNS,u'list-style'),
		(TEXTNS,u'notes-configuration'),
		(TEXTNS,u'outline-style'),
	),
	(OFFICENS,u'text') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'forms'),
		(TABLENS,u'calculation-settings'),
		(TABLENS,u'consolidation'),
		(TABLENS,u'content-validations'),
		(TABLENS,u'data-pilot-tables'),
		(TABLENS,u'database-ranges'),
		(TABLENS,u'dde-links'),
		(TABLENS,u'label-ranges'),
		(TABLENS,u'named-expressions'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'page-sequence'),
		(TEXTNS,u'section'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'tracked-changes'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'user-index'),
		(TEXTNS,u'variable-decls'),
	),
	(PRESENTATIONNS,u'animation-group') : (
		(PRESENTATIONNS,u'dim'),
		(PRESENTATIONNS,u'hide-shape'),
		(PRESENTATIONNS,u'hide-text'),
		(PRESENTATIONNS,u'play'),
		(PRESENTATIONNS,u'show-shape'),
		(PRESENTATIONNS,u'show-text'),
	),
	(PRESENTATIONNS,u'animations') : (
		(PRESENTATIONNS,u'animation-group'),
		(PRESENTATIONNS,u'dim'),
		(PRESENTATIONNS,u'hide-shape'),
		(PRESENTATIONNS,u'hide-text'),
		(PRESENTATIONNS,u'play'),
		(PRESENTATIONNS,u'show-shape'),
		(PRESENTATIONNS,u'show-text'),
	),
	(PRESENTATIONNS,u'date-time') : (
	),
	(PRESENTATIONNS,u'date-time-decl') : (
	),
	(PRESENTATIONNS,u'dim') : (
		(PRESENTATIONNS,u'sound'),
	),
	(PRESENTATIONNS,u'event-listener') : (
		(PRESENTATIONNS,u'sound'),
	),
	(PRESENTATIONNS,u'footer') : (
	),
	(PRESENTATIONNS,u'footer-decl') : (
	),
	(PRESENTATIONNS,u'header') : (
	),
	(PRESENTATIONNS,u'header-decl') : (
	),
	(PRESENTATIONNS,u'hide-shape') : (
		(PRESENTATIONNS,u'sound'),
	),
	(PRESENTATIONNS,u'hide-text') : (
		(PRESENTATIONNS,u'sound'),
	),
# allowed_children
	(PRESENTATIONNS,u'notes') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'forms'),
	),
	(PRESENTATIONNS,u'placeholder') : (
	),
	(PRESENTATIONNS,u'play') : (
	),
	(PRESENTATIONNS,u'settings') : (
		(PRESENTATIONNS,u'show'),
	),
	(PRESENTATIONNS,u'show') : (
	),
	(PRESENTATIONNS,u'show-shape') : (
		(PRESENTATIONNS,u'sound'),
	),
	(PRESENTATIONNS,u'show-text') : (
		(PRESENTATIONNS,u'sound'),
	),
	(PRESENTATIONNS,u'sound') : (
	),
	(SCRIPTNS,u'event-listener') : (
	),
	(STYLENS,u'background-image') : (
		(OFFICENS,u'binary-data'),
	),
# allowed_children
	(STYLENS,u'chart-properties') : (
		(CHARTNS,u'label-separator'),
		(CHARTNS,u'symbol-image'),
	),
	(STYLENS,u'column') : (
	),
	(STYLENS,u'column-sep') : (
	),
	(STYLENS,u'columns') : (
		(STYLENS,u'column'),
		(STYLENS,u'column-sep'),
	),
# allowed_children
	(STYLENS,u'default-page-layout') : (
		(STYLENS,u'page-layout-properties'),
		(STYLENS,u'header-style'),
		(STYLENS,u'footer-style'),
	),
	(STYLENS,u'default-style') : (
		(STYLENS,u'chart-properties'),
		(STYLENS,u'drawing-page-properties'),
		(STYLENS,u'graphic-properties'),
		(STYLENS,u'paragraph-properties'),
		(STYLENS,u'ruby-properties'),
		(STYLENS,u'section-properties'),
		(STYLENS,u'table-cell-properties'),
		(STYLENS,u'table-column-properties'),
		(STYLENS,u'table-properties'),
		(STYLENS,u'table-row-properties'),
		(STYLENS,u'text-properties'),
	),
	(STYLENS,u'drawing-page-properties') : (
		(PRESENTATIONNS,u'sound'),
	),
	(STYLENS,u'drop-cap') : (
	),
	(STYLENS,u'font-face') : (
		(SVGNS,u'definition-src'),
		(SVGNS,u'font-face-src'),
	),
	(STYLENS,u'footer') : (
		(STYLENS,u'region-center'),
		(STYLENS,u'region-left'),
		(STYLENS,u'region-right'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'index-title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'tracked-changes'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'user-index'),
		(TEXTNS,u'variable-decls'),
	),
# allowed_children
	(STYLENS,u'footer-left') : (
		(STYLENS,u'region-center'),
		(STYLENS,u'region-left'),
		(STYLENS,u'region-right'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'index-title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'tracked-changes'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'user-index'),
		(TEXTNS,u'variable-decls'),
	),
	(STYLENS,u'footer-style') : (
		(STYLENS,u'header-footer-properties'),
	),
	(STYLENS,u'footnote-sep') : (
	),
	(STYLENS,u'graphic-properties') : (
		(STYLENS,u'background-image'),
		(STYLENS,u'columns'),
		(TEXTNS,u'list-style'),
	),
# allowed_children
	(STYLENS,u'handout-master') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
	),
	(STYLENS,u'header') : (
		(STYLENS,u'region-center'),
		(STYLENS,u'region-left'),
		(STYLENS,u'region-right'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'index-title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'tracked-changes'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'user-index'),
		(TEXTNS,u'variable-decls'),
	),
# allowed_children
	(STYLENS,u'header-footer-properties') : (
		(STYLENS,u'background-image'),
	),
	(STYLENS,u'header-left') : (
		(STYLENS,u'region-center'),
		(STYLENS,u'region-left'),
		(STYLENS,u'region-right'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'alphabetical-index-auto-mark-file'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'dde-connection-decls'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'index-title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'sequence-decls'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'tracked-changes'),
		(TEXTNS,u'user-field-decls'),
		(TEXTNS,u'user-index'),
		(TEXTNS,u'variable-decls'),
	),
	(STYLENS,u'header-style') : (
		(STYLENS,u'header-footer-properties'),
	),
# allowed_children
	(STYLENS,u'list-level-label-alignment') : (
	),
	(STYLENS,u'list-level-properties') : (
		(STYLENS,u'list-level-label-alignment'),
	),
	(STYLENS,u'map') : (
	),
	(STYLENS,u'master-page') : (
		(ANIMNS,u'animate'),
		(ANIMNS,u'animateColor'),
		(ANIMNS,u'animateMotion'),
		(ANIMNS,u'animateTransform'),
		(ANIMNS,u'audio'),
		(ANIMNS,u'command'),
		(ANIMNS,u'iterate'),
		(ANIMNS,u'par'),
		(ANIMNS,u'seq'),
		(ANIMNS,u'set'),
		(ANIMNS,u'transitionFilter'),
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'layer-set'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'forms'),
		(PRESENTATIONNS,u'notes'),
		(STYLENS,u'footer'),
		(STYLENS,u'footer-left'),
		(STYLENS,u'header'),
		(STYLENS,u'header-left'),
	),
	(STYLENS,u'page-layout') : (
		(STYLENS,u'footer-style'),
		(STYLENS,u'header-style'),
		(STYLENS,u'page-layout-properties'),
	),
	(STYLENS,u'page-layout-properties') : (
		(STYLENS,u'background-image'),
		(STYLENS,u'columns'),
		(STYLENS,u'footnote-sep'),
	),
# allowed_children
	(STYLENS,u'paragraph-properties') : (
		(STYLENS,u'background-image'),
		(STYLENS,u'drop-cap'),
		(STYLENS,u'tab-stops'),
	),
	(STYLENS,u'presentation-page-layout') : (
		(PRESENTATIONNS,u'placeholder'),
	),
	(STYLENS,u'region-center') : (
		(TEXTNS,u'p'),
	),
	(STYLENS,u'region-left') : (
		(TEXTNS,u'p'),
	),
	(STYLENS,u'region-right') : (
		(TEXTNS,u'p'),
	),
	(STYLENS,u'ruby-properties') : (
	),
	(STYLENS,u'section-properties') : (
		(STYLENS,u'background-image'),
		(STYLENS,u'columns'),
		(TEXTNS,u'notes-configuration'),
	),
	(STYLENS,u'style') : (
		(STYLENS,u'chart-properties'),
		(STYLENS,u'drawing-page-properties'),
		(STYLENS,u'graphic-properties'),
		(STYLENS,u'map'),
		(STYLENS,u'paragraph-properties'),
		(STYLENS,u'ruby-properties'),
		(STYLENS,u'section-properties'),
		(STYLENS,u'table-cell-properties'),
		(STYLENS,u'table-column-properties'),
		(STYLENS,u'table-properties'),
		(STYLENS,u'table-row-properties'),
		(STYLENS,u'text-properties'),
	),
	(STYLENS,u'tab-stop') : (
	),
	(STYLENS,u'tab-stops') : (
		(STYLENS,u'tab-stop'),
	),
# allowed_children
	(STYLENS,u'table-cell-properties') : (
		(STYLENS,u'background-image'),
	),
	(STYLENS,u'table-column-properties') : (
	),
	(STYLENS,u'table-properties') : (
		(STYLENS,u'background-image'),
	),
	(STYLENS,u'table-row-properties') : (
		(STYLENS,u'background-image'),
	),
	(STYLENS,u'text-properties') : (
	),
	(SVGNS,u'definition-src') : (
	),
	(SVGNS,u'desc') : (
	),
	(SVGNS,u'font-face-format') : (
	),
	(SVGNS,u'font-face-name') : (
	),
	(SVGNS,u'font-face-src') : (
		(SVGNS,u'font-face-name'),
		(SVGNS,u'font-face-uri'),
	),
	(SVGNS,u'font-face-uri') : (
		(SVGNS,u'font-face-format'),
	),
	(SVGNS,u'linearGradient') : (
		(SVGNS,u'stop'),
	),
	(SVGNS,u'radialGradient') : (
		(SVGNS,u'stop'),
	),
	(SVGNS,u'stop') : (
	),
	(SVGNS,u'title') : (
	),
# allowed_children
	(TABLENS,u'background') : (
	),
	(TABLENS,u'body') : (
	),
	(TABLENS,u'calculation-settings') : (
		(TABLENS,u'iteration'),
		(TABLENS,u'null-date'),
	),
# allowed_children
	(TABLENS,u'cell-address') : (
	),
	(TABLENS,u'cell-content-change') : (
		(OFFICENS,u'change-info'),
		(TABLENS,u'cell-address'),
		(TABLENS,u'deletions'),
		(TABLENS,u'dependencies'),
		(TABLENS,u'previous'),
	),
	(TABLENS,u'cell-content-deletion') : (
		(TABLENS,u'cell-address'),
		(TABLENS,u'change-track-table-cell'),
	),
	(TABLENS,u'cell-range-source') : (
	),
	(TABLENS,u'change-deletion') : (
	),
	(TABLENS,u'change-track-table-cell') : (
		(TEXTNS,u'p'),
	),
	(TABLENS,u'consolidation') : (
	),
	(TABLENS,u'content-validation') : (
		(OFFICENS,u'event-listeners'),
		(TABLENS,u'error-macro'),
		(TABLENS,u'error-message'),
		(TABLENS,u'help-message'),
	),
# allowed_children
	(TABLENS,u'content-validations') : (
		(TABLENS,u'content-validation'),
	),
	(TABLENS,u'covered-table-cell') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(TABLENS,u'cell-range-source'),
		(TABLENS,u'detective'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
# allowed_children
	(TABLENS,u'cut-offs') : (
		(TABLENS,u'insertion-cut-off'),
		(TABLENS,u'movement-cut-off'),
	),
	(TABLENS,u'data-pilot-display-info') : (
	),
	(TABLENS,u'data-pilot-field') : (
		(TABLENS,u'data-pilot-field-reference'),
		(TABLENS,u'data-pilot-groups'),
		(TABLENS,u'data-pilot-level'),
	),
	(TABLENS,u'data-pilot-field-reference') : (
	),
	(TABLENS,u'data-pilot-group') : (
		(TABLENS,u'data-pilot-group-member'),
	),
	(TABLENS,u'data-pilot-group-member') : (
	),
	(TABLENS,u'data-pilot-groups') : (
		(TABLENS,u'data-pilot-group'),
	),
	(TABLENS,u'data-pilot-layout-info') : (
	),
	(TABLENS,u'data-pilot-level') : (
		(TABLENS,u'data-pilot-display-info'),
		(TABLENS,u'data-pilot-layout-info'),
		(TABLENS,u'data-pilot-members'),
		(TABLENS,u'data-pilot-sort-info'),
		(TABLENS,u'data-pilot-subtotals'),
	),
	(TABLENS,u'data-pilot-member') : (
	),
	(TABLENS,u'data-pilot-members') : (
		(TABLENS,u'data-pilot-member'),
	),
	(TABLENS,u'data-pilot-sort-info') : (
	),
	(TABLENS,u'data-pilot-subtotal') : (
	),
	(TABLENS,u'data-pilot-subtotals') : (
		(TABLENS,u'data-pilot-subtotal'),
	),
# allowed_children
	(TABLENS,u'data-pilot-table') : (
		(TABLENS,u'data-pilot-field'),
		(TABLENS,u'database-source-query'),
		(TABLENS,u'database-source-sql'),
		(TABLENS,u'database-source-table'),
		(TABLENS,u'source-cell-range'),
		(TABLENS,u'source-service'),
	),
	(TABLENS,u'data-pilot-tables') : (
		(TABLENS,u'data-pilot-table'),
	),
	(TABLENS,u'database-range') : (
		(TABLENS,u'database-source-query'),
		(TABLENS,u'database-source-sql'),
		(TABLENS,u'database-source-table'),
		(TABLENS,u'filter'),
		(TABLENS,u'sort'),
		(TABLENS,u'subtotal-rules'),
	),
	(TABLENS,u'database-ranges') : (
		(TABLENS,u'database-range'),
	),
	(TABLENS,u'database-source-query') : (
	),
	(TABLENS,u'database-source-sql') : (
	),
	(TABLENS,u'database-source-table') : (
	),
# allowed_children
	(TABLENS,u'dde-link') : (
		(OFFICENS,u'dde-source'),
		(TABLENS,u'table'),
	),
	(TABLENS,u'dde-links') : (
		(TABLENS,u'dde-link'),
	),
	(TABLENS,u'deletion') : (
		(OFFICENS,u'change-info'),
		(TABLENS,u'cut-offs'),
		(TABLENS,u'deletions'),
		(TABLENS,u'dependencies'),
	),
	(TABLENS,u'deletions') : (
		(TABLENS,u'cell-content-deletion'),
		(TABLENS,u'change-deletion'),
	),
	(TABLENS,u'dependencies') : (
		(TABLENS,u'dependency'),
	),
	(TABLENS,u'dependency') : (
	),
# allowed_children
	(TABLENS,u'desc') : (
	),
	(TABLENS,u'detective') : (
		(TABLENS,u'highlighted-range'),
		(TABLENS,u'operation'),
	),
# allowed_children
	(TABLENS,u'error-macro') : (
	),
	(TABLENS,u'error-message') : (
		(TEXTNS,u'p'),
	),
	(TABLENS,u'even-columns') : (
	),
	(TABLENS,u'even-rows') : (
	),
	(TABLENS,u'filter') : (
		(TABLENS,u'filter-and'),
		(TABLENS,u'filter-condition'),
		(TABLENS,u'filter-or'),
	),
	(TABLENS,u'filter-and') : (
		(TABLENS,u'filter-condition'),
		(TABLENS,u'filter-or'),
	),
	(TABLENS,u'filter-condition') : (
		(TABLENS,u'filter-set-item'),
	),
	(TABLENS,u'filter-or') : (
		(TABLENS,u'filter-and'),
		(TABLENS,u'filter-condition'),
	),
# allowed_children
	(TABLENS,u'filter-set-item') : (
	),
# allowed_children
	(TABLENS,u'first-column') : (
	),
	(TABLENS,u'first-row') : (
	),
	(TABLENS,u'help-message') : (
		(TEXTNS,u'p'),
	),
	(TABLENS,u'highlighted-range') : (
	),
	(TABLENS,u'insertion') : (
		(OFFICENS,u'change-info'),
		(TABLENS,u'deletions'),
		(TABLENS,u'dependencies'),
	),
	(TABLENS,u'insertion-cut-off') : (
	),
	(TABLENS,u'iteration') : (
	),
	(TABLENS,u'label-range') : (
	),
	(TABLENS,u'label-ranges') : (
		(TABLENS,u'label-range'),
	),
	(TABLENS,u'last-column') : (
	),
	(TABLENS,u'last-row') : (
	),
	(TABLENS,u'movement') : (
		(OFFICENS,u'change-info'),
		(TABLENS,u'deletions'),
		(TABLENS,u'dependencies'),
		(TABLENS,u'source-range-address'),
		(TABLENS,u'target-range-address'),
	),
	(TABLENS,u'movement-cut-off') : (
	),
	(TABLENS,u'named-expression') : (
	),
	(TABLENS,u'named-expressions') : (
		(TABLENS,u'named-expression'),
		(TABLENS,u'named-range'),
	),
# allowed_children
	(TABLENS,u'named-range') : (
	),
	(TABLENS,u'null-date') : (
	),
	(TABLENS,u'odd-columns') : (
	),
	(TABLENS,u'odd-rows') : (
	),
	(TABLENS,u'operation') : (
	),
	(TABLENS,u'previous') : (
		(TABLENS,u'change-track-table-cell'),
	),
	(TABLENS,u'scenario') : (
	),
	(TABLENS,u'shapes') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
	),
# allowed_children
	(TABLENS,u'sort') : (
		(TABLENS,u'sort-by'),
	),
	(TABLENS,u'sort-by') : (
	),
	(TABLENS,u'sort-groups') : (
	),
	(TABLENS,u'source-cell-range') : (
		(TABLENS,u'filter'),
	),
	(TABLENS,u'source-range-address') : (
	),
	(TABLENS,u'source-service') : (
	),
	(TABLENS,u'subtotal-field') : (
	),
	(TABLENS,u'subtotal-rule') : (
		(TABLENS,u'subtotal-field'),
	),
	(TABLENS,u'subtotal-rules') : (
		(TABLENS,u'sort-groups'),
		(TABLENS,u'subtotal-rule'),
	),
# allowed_children
	(TABLENS,u'table') : (
		(OFFICENS,u'dde-source'),
		(OFFICENS,u'forms'),
		(TABLENS,u'desc'),
		(TABLENS,u'named-expressions'),
		(TABLENS,u'scenario'),
		(TABLENS,u'shapes'),
		(TABLENS,u'table-column'),
		(TABLENS,u'table-column-group'),
		(TABLENS,u'table-columns'),
		(TABLENS,u'table-header-columns'),
		(TABLENS,u'table-header-rows'),
		(TABLENS,u'table-row'),
		(TABLENS,u'table-row-group'),
		(TABLENS,u'table-rows'),
		(TABLENS,u'table-source'),
		(TABLENS,u'title'),
		(TEXTNS,u'soft-page-break'),
	),
	(TABLENS,u'table-cell') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(TABLENS,u'cell-range-source'),
		(TABLENS,u'detective'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
# allowed_children
	(TABLENS,u'table-column') : (
	),
	(TABLENS,u'table-column-group') : (
		(TABLENS,u'table-column'),
		(TABLENS,u'table-column-group'),
		(TABLENS,u'table-columns'),
		(TABLENS,u'table-header-columns'),
	),
	(TABLENS,u'table-columns') : (
		(TABLENS,u'table-column'),
	),
	(TABLENS,u'table-header-columns') : (
		(TABLENS,u'table-column'),
	),
	(TABLENS,u'table-header-rows') : (
		(TABLENS,u'table-row'),
		(TEXTNS,u'soft-page-break'),
	),
	(TABLENS,u'table-row') : (
		(TABLENS,u'covered-table-cell'),
		(TABLENS,u'table-cell'),
	),
	(TABLENS,u'table-row-group') : (
		(TABLENS,u'table-header-rows'),
		(TABLENS,u'table-row'),
		(TABLENS,u'table-row-group'),
		(TABLENS,u'table-rows'),
		(TEXTNS,u'soft-page-break'),
	),
	(TABLENS,u'table-rows') : (
		(TABLENS,u'table-row'),
		(TEXTNS,u'soft-page-break'),
	),
# allowed_children
	(TABLENS,u'table-source') : (
	),
	(TABLENS,u'table-template') : (
		(TABLENS,u'background'),
		(TABLENS,u'body'),
		(TABLENS,u'even-columns'),
		(TABLENS,u'even-rows'),
		(TABLENS,u'first-column'),
		(TABLENS,u'first-row'),
		(TABLENS,u'last-column'),
		(TABLENS,u'last-row'),
		(TABLENS,u'odd-columns'),
		(TABLENS,u'odd-rows'),
	),
	(TABLENS,u'target-range-address') : (
	),
# allowed_children
	(TABLENS,u'title') : (
	),
	(TABLENS,u'tracked-changes') : (
		(TABLENS,u'cell-content-change'),
		(TABLENS,u'deletion'),
		(TABLENS,u'insertion'),
		(TABLENS,u'movement'),
	),
# allowed_children
	(TEXTNS,u'a') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(OFFICENS,u'event-listeners'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
		(TEXTNS,u'word-count'),

	),
# allowed_children
	(TEXTNS,u'alphabetical-index') : (
		(TEXTNS,u'alphabetical-index-source'),
		(TEXTNS,u'index-body'),
	),
	(TEXTNS,u'alphabetical-index-auto-mark-file') : (
	),
	(TEXTNS,u'alphabetical-index-entry-template') : (
		(TEXTNS,u'index-entry-chapter'),
		(TEXTNS,u'index-entry-page-number'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
		(TEXTNS,u'index-entry-text'),
	),
	(TEXTNS,u'alphabetical-index-mark') : (
	),
	(TEXTNS,u'alphabetical-index-mark-end') : (
	),
	(TEXTNS,u'alphabetical-index-mark-start') : (
	),
	(TEXTNS,u'alphabetical-index-source') : (
		(TEXTNS,u'alphabetical-index-entry-template'),
		(TEXTNS,u'index-title-template'),
	),
	(TEXTNS,u'author-initials') : (
	),
	(TEXTNS,u'author-name') : (
	),
	(TEXTNS,u'bibliography') : (
		(TEXTNS,u'bibliography-source'),
		(TEXTNS,u'index-body'),
	),
	(TEXTNS,u'bibliography-configuration') : (
		(TEXTNS,u'sort-key'),
	),
	(TEXTNS,u'bibliography-entry-template') : (
		(TEXTNS,u'index-entry-bibliography'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
	),
# allowed_children
	(TEXTNS,u'bibliography-mark') : (
	),
	(TEXTNS,u'bibliography-source') : (
		(TEXTNS,u'bibliography-entry-template'),
		(TEXTNS,u'index-title-template'),
	),
	(TEXTNS,u'bookmark') : (
	),
	(TEXTNS,u'bookmark-end') : (
	),
	(TEXTNS,u'bookmark-ref') : (
	),
	(TEXTNS,u'bookmark-start') : (
	),
	(TEXTNS,u'change') : (
	),
	(TEXTNS,u'change-end') : (
	),
	(TEXTNS,u'change-start') : (
	),
	(TEXTNS,u'changed-region') : (
		(TEXTNS,u'deletion'),
		(TEXTNS,u'format-change'),
		(TEXTNS,u'insertion'),
	),
	(TEXTNS,u'chapter') : (
	),
	(TEXTNS,u'character-count') : (
	),
	(TEXTNS,u'conditional-text') : (
	),
	(TEXTNS,u'creation-date') : (
	),
	(TEXTNS,u'creation-time') : (
	),
	(TEXTNS,u'creator') : (
	),
	(TEXTNS,u'database-display') : (
		(FORMNS,u'connection-resource'),
	),
	(TEXTNS,u'database-name') : (
		(FORMNS,u'connection-resource'),
	),
	(TEXTNS,u'database-next') : (
		(FORMNS,u'connection-resource'),
	),
	(TEXTNS,u'database-row-number') : (
		(FORMNS,u'connection-resource'),
	),
	(TEXTNS,u'database-row-select') : (
		(FORMNS,u'connection-resource'),
	),
	(TEXTNS,u'date') : (
	),
	(TEXTNS,u'dde-connection') : (
	),
	(TEXTNS,u'dde-connection-decl') : (
	),
	(TEXTNS,u'dde-connection-decls') : (
		(TEXTNS,u'dde-connection-decl'),
	),
# allowed_children
	(TEXTNS,u'deletion') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'change-info'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
	(TEXTNS,u'description') : (
	),
	(TEXTNS,u'editing-cycles') : (
	),
	(TEXTNS,u'editing-duration') : (
	),
	(TEXTNS,u'execute-macro') : (
		(OFFICENS,u'event-listeners'),
	),
	(TEXTNS,u'expression') : (
	),
	(TEXTNS,u'file-name') : (
	),
	(TEXTNS,u'format-change') : (
		(OFFICENS,u'change-info'),
	),
# allowed_children
	(TEXTNS,u'h') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'a'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'number'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
		(TEXTNS,u'word-count'),

	),
# allowed_children
	(TEXTNS,u'hidden-paragraph') : (
	),
	(TEXTNS,u'hidden-text') : (
	),
	(TEXTNS,u'illustration-index') : (
		(TEXTNS,u'illustration-index-source'),
		(TEXTNS,u'index-body'),
	),
	(TEXTNS,u'illustration-index-entry-template') : (
		(TEXTNS,u'index-entry-chapter'),
		(TEXTNS,u'index-entry-page-number'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
		(TEXTNS,u'index-entry-text'),
	),
	(TEXTNS,u'illustration-index-source') : (
		(TEXTNS,u'illustration-index-entry-template'),
		(TEXTNS,u'index-title-template'),
	),
	(TEXTNS,u'image-count') : (
	),
# allowed_children
	(TEXTNS,u'index-body') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'index-title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
	(TEXTNS,u'index-entry-bibliography') : (
	),
	(TEXTNS,u'index-entry-chapter') : (
	),
	(TEXTNS,u'index-entry-link-end') : (
	),
	(TEXTNS,u'index-entry-link-start') : (
	),
	(TEXTNS,u'index-entry-page-number') : (
	),
	(TEXTNS,u'index-entry-span') : (
	),
	(TEXTNS,u'index-entry-tab-stop') : (
	),
	(TEXTNS,u'index-entry-text') : (
	),
	(TEXTNS,u'index-source-style') : (
	),
	(TEXTNS,u'index-source-styles') : (
		(TEXTNS,u'index-source-style'),
	),
# allowed_children
	(TEXTNS,u'index-title') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'index-title'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
	(TEXTNS,u'index-title-template') : (
	),
	(TEXTNS,u'initial-creator') : (
	),
	(TEXTNS,u'insertion') : (
		(OFFICENS,u'change-info'),
	),
	(TEXTNS,u'keywords') : (
	),
	(TEXTNS,u'line-break') : (
	),
	(TEXTNS,u'linenumbering-configuration') : (
		(TEXTNS,u'linenumbering-separator'),
	),
	(TEXTNS,u'linenumbering-separator') : (
	),
	(TEXTNS,u'list') : (
		(TEXTNS,u'list-header'),
		(TEXTNS,u'list-item'),
	),
	(TEXTNS,u'list-header') : (
		(TEXTNS,u'h'),
		(TEXTNS,u'list'),
		(TEXTNS,u'number'),
		(TEXTNS,u'p'),
		(TEXTNS,u'soft-page-break'),
	),
	(TEXTNS,u'list-item') : (
		(TEXTNS,u'h'),
		(TEXTNS,u'list'),
		(TEXTNS,u'number'),
		(TEXTNS,u'p'),
		(TEXTNS,u'soft-page-break'),
	),
	(TEXTNS,u'list-level-style-bullet') : (
		(STYLENS,u'list-level-properties'),
		(STYLENS,u'text-properties'),
	),
	(TEXTNS,u'list-level-style-image') : (
		(OFFICENS,u'binary-data'),
		(STYLENS,u'list-level-properties'),
	),
	(TEXTNS,u'list-level-style-number') : (
		(STYLENS,u'list-level-properties'),
		(STYLENS,u'text-properties'),
	),
	(TEXTNS,u'list-style') : (
		(TEXTNS,u'list-level-style-bullet'),
		(TEXTNS,u'list-level-style-image'),
		(TEXTNS,u'list-level-style-number'),
	),
	(TEXTNS,u'measure') : (
	),
# allowed_children
	(TEXTNS,u'meta') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'a'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
		(TEXTNS,u'word-count'),
	),
# allowed_children
	(TEXTNS,u'meta-field') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'a'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
		(TEXTNS,u'word-count'),
	),
# allowed_children
	(TEXTNS,u'modification-date') : (
	),
	(TEXTNS,u'modification-time') : (
	),
	(TEXTNS,u'note') : (
		(TEXTNS,u'note-body'),
		(TEXTNS,u'note-citation'),
	),
# allowed_children
	(TEXTNS,u'note-body') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
	(TEXTNS,u'note-citation') : (
	),
	(TEXTNS,u'note-continuation-notice-backward') : (
	),
	(TEXTNS,u'note-continuation-notice-forward') : (
	),
	(TEXTNS,u'note-ref') : (
	),
	(TEXTNS,u'notes-configuration') : (
		(TEXTNS,u'note-continuation-notice-backward'),
		(TEXTNS,u'note-continuation-notice-forward'),
	),
	(TEXTNS,u'number') : (
	),
	(TEXTNS,u'numbered-paragraph') : (
		(TEXTNS,u'h'),
		(TEXTNS,u'number'),
		(TEXTNS,u'p'),
	),
	(TEXTNS,u'object-count') : (
	),
	(TEXTNS,u'object-index') : (
		(TEXTNS,u'index-body'),
		(TEXTNS,u'object-index-source'),
	),
	(TEXTNS,u'object-index-entry-template') : (
		(TEXTNS,u'index-entry-chapter'),
		(TEXTNS,u'index-entry-page-number'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
		(TEXTNS,u'index-entry-text'),
	),
	(TEXTNS,u'object-index-source') : (
		(TEXTNS,u'index-title-template'),
		(TEXTNS,u'object-index-entry-template'),
	),
	(TEXTNS,u'outline-level-style') : (
		(STYLENS,u'list-level-properties'),
		(STYLENS,u'text-properties'),
	),
	(TEXTNS,u'outline-style') : (
		(TEXTNS,u'outline-level-style'),
	),
# allowed_children
	(TEXTNS,u'p') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'a'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
		(TEXTNS,u'word-count'),
	),
	(TEXTNS,u'page') : (
	),
	(TEXTNS,u'page-count') : (
	),
	(TEXTNS,u'page-continuation') : (
	),
	(TEXTNS,u'page-number') : (
	),
	(TEXTNS,u'page-sequence') : (
		(TEXTNS,u'page'),
	),
	(TEXTNS,u'page-variable-get') : (
	),
	(TEXTNS,u'page-variable-set') : (
	),
	(TEXTNS,u'paragraph-count') : (
	),
	(TEXTNS,u'placeholder') : (
	),
	(TEXTNS,u'print-date') : (
	),
	(TEXTNS,u'print-time') : (
	),
	(TEXTNS,u'printed-by') : (
	),
	(TEXTNS,u'reference-mark') : (
	),
	(TEXTNS,u'reference-mark-end') : (
	),
# allowed_children
	(TEXTNS,u'reference-mark-start') : (
	),
	(TEXTNS,u'reference-ref') : (
	),
	(TEXTNS,u'ruby') : (
		(TEXTNS,u'ruby-base'),
		(TEXTNS,u'ruby-text'),
	),
	(TEXTNS,u'ruby-base') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'a'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
		(TEXTNS,u'word-count'),

	),
# allowed_children
	(TEXTNS,u'ruby-text') : (
	),
	(TEXTNS,u's') : (
	),
	(TEXTNS,u'script') : (
	),
	(TEXTNS,u'section') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'dde-source'),
		(TABLENS,u'table'),
		(TEXTNS,u'alphabetical-index'),
		(TEXTNS,u'bibliography'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'h'),
		(TEXTNS,u'illustration-index'),
		(TEXTNS,u'list'),
		(TEXTNS,u'numbered-paragraph'),
		(TEXTNS,u'object-index'),
		(TEXTNS,u'p'),
		(TEXTNS,u'section'),
		(TEXTNS,u'section-source'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'table-index'),
		(TEXTNS,u'table-of-content'),
		(TEXTNS,u'user-index'),
	),
	(TEXTNS,u'section-source') : (
	),
	(TEXTNS,u'sender-city') : (
	),
	(TEXTNS,u'sender-company') : (
	),
	(TEXTNS,u'sender-country') : (
	),
# allowed_children
	(TEXTNS,u'sender-email') : (
	),
	(TEXTNS,u'sender-fax') : (
	),
	(TEXTNS,u'sender-firstname') : (
	),
	(TEXTNS,u'sender-initials') : (
	),
	(TEXTNS,u'sender-lastname') : (
	),
	(TEXTNS,u'sender-phone-private') : (
	),
	(TEXTNS,u'sender-phone-work') : (
	),
	(TEXTNS,u'sender-position') : (
	),
	(TEXTNS,u'sender-postal-code') : (
	),
	(TEXTNS,u'sender-state-or-province') : (
	),
	(TEXTNS,u'sender-street') : (
	),
	(TEXTNS,u'sender-title') : (
	),
	(TEXTNS,u'sequence') : (
	),
	(TEXTNS,u'sequence-decl') : (
	),
	(TEXTNS,u'sequence-decls') : (
		(TEXTNS,u'sequence-decl'),
	),
	(TEXTNS,u'sequence-ref') : (
	),
	(TEXTNS,u'sheet-name') : (
	),
	(TEXTNS,u'soft-page-break') : (
	),
	(TEXTNS,u'sort-key') : (
	),
# allowed_children
	(TEXTNS,u'span') : (
		(DR3DNS,u'scene'),
		(DRAWNS,u'a'),
		(DRAWNS,u'caption'),
		(DRAWNS,u'circle'),
		(DRAWNS,u'connector'),
		(DRAWNS,u'control'),
		(DRAWNS,u'custom-shape'),
		(DRAWNS,u'ellipse'),
		(DRAWNS,u'frame'),
		(DRAWNS,u'g'),
		(DRAWNS,u'line'),
		(DRAWNS,u'measure'),
		(DRAWNS,u'page-thumbnail'),
		(DRAWNS,u'path'),
		(DRAWNS,u'polygon'),
		(DRAWNS,u'polyline'),
		(DRAWNS,u'rect'),
		(DRAWNS,u'regular-polygon'),
		(OFFICENS,u'annotation'),
		(OFFICENS,u'annotation-end'),
		(PRESENTATIONNS,u'date-time'),
		(PRESENTATIONNS,u'footer'),
		(PRESENTATIONNS,u'header'),
		(TEXTNS,u'a'),
		(TEXTNS,u'alphabetical-index-mark'),
		(TEXTNS,u'alphabetical-index-mark-end'),
		(TEXTNS,u'alphabetical-index-mark-start'),
		(TEXTNS,u'author-initials'),
		(TEXTNS,u'author-name'),
		(TEXTNS,u'bibliography-mark'),
		(TEXTNS,u'bookmark'),
		(TEXTNS,u'bookmark-end'),
		(TEXTNS,u'bookmark-ref'),
		(TEXTNS,u'bookmark-start'),
		(TEXTNS,u'change'),
		(TEXTNS,u'change-end'),
		(TEXTNS,u'change-start'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'conditional-text'),
		(TEXTNS,u'creation-date'),
		(TEXTNS,u'creation-time'),
		(TEXTNS,u'creator'),
		(TEXTNS,u'database-display'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'database-next'),
		(TEXTNS,u'database-row-number'),
		(TEXTNS,u'database-row-select'),
		(TEXTNS,u'date'),
		(TEXTNS,u'dde-connection'),
		(TEXTNS,u'description'),
		(TEXTNS,u'editing-cycles'),
		(TEXTNS,u'editing-duration'),
		(TEXTNS,u'execute-macro'),
		(TEXTNS,u'expression'),
		(TEXTNS,u'file-name'),
		(TEXTNS,u'hidden-paragraph'),
		(TEXTNS,u'hidden-text'),
		(TEXTNS,u'initial-creator'),
		(TEXTNS,u'keywords'),
		(TEXTNS,u'line-break'),
		(TEXTNS,u'measure'),
		(TEXTNS,u'meta'),
		(TEXTNS,u'meta-field'),
		(TEXTNS,u'modification-date'),
		(TEXTNS,u'modification-time'),
		(TEXTNS,u'note'),
		(TEXTNS,u'note-ref'),
		(TEXTNS,u'page-count'),
		(TEXTNS,u'paragraph-count'),
		(TEXTNS,u'word-count'),
		(TEXTNS,u'character-count'),
		(TEXTNS,u'table-count'),
		(TEXTNS,u'image-count'),
		(TEXTNS,u'object-count'),
		(TEXTNS,u'page-continuation'),
		(TEXTNS,u'page-number'),
		(TEXTNS,u'page-variable-get'),
		(TEXTNS,u'page-variable-set'),
		(TEXTNS,u'placeholder'),
		(TEXTNS,u'print-date'),
		(TEXTNS,u'print-time'),
		(TEXTNS,u'printed-by'),
		(TEXTNS,u'reference-mark'),
		(TEXTNS,u'reference-mark-end'),
		(TEXTNS,u'reference-mark-start'),
		(TEXTNS,u'reference-ref'),
		(TEXTNS,u'ruby'),
		(TEXTNS,u's'),
		(TEXTNS,u'script'),
		(TEXTNS,u'sender-city'),
		(TEXTNS,u'sender-company'),
		(TEXTNS,u'sender-country'),
		(TEXTNS,u'sender-email'),
		(TEXTNS,u'sender-fax'),
		(TEXTNS,u'sender-firstname'),
		(TEXTNS,u'sender-initials'),
		(TEXTNS,u'sender-lastname'),
		(TEXTNS,u'sender-phone-private'),
		(TEXTNS,u'sender-phone-work'),
		(TEXTNS,u'sender-position'),
		(TEXTNS,u'sender-postal-code'),
		(TEXTNS,u'sender-state-or-province'),
		(TEXTNS,u'sender-street'),
		(TEXTNS,u'sender-title'),
		(TEXTNS,u'sequence'),
		(TEXTNS,u'sequence-ref'),
		(TEXTNS,u'sheet-name'),
		(TEXTNS,u'soft-page-break'),
		(TEXTNS,u'span'),
		(TEXTNS,u'subject'),
		(TEXTNS,u'tab'),
		(TEXTNS,u'table-formula'),
		(TEXTNS,u'template-name'),
		(TEXTNS,u'text-input'),
		(TEXTNS,u'time'),
		(TEXTNS,u'title'),
		(TEXTNS,u'toc-mark'),
		(TEXTNS,u'toc-mark-end'),
		(TEXTNS,u'toc-mark-start'),
		(TEXTNS,u'user-defined'),
		(TEXTNS,u'user-field-get'),
		(TEXTNS,u'user-field-input'),
		(TEXTNS,u'user-index-mark'),
		(TEXTNS,u'user-index-mark-end'),
		(TEXTNS,u'user-index-mark-start'),
		(TEXTNS,u'variable-get'),
		(TEXTNS,u'variable-input'),
		(TEXTNS,u'variable-set'),
	),
# allowed_children
	(TEXTNS,u'subject') : (
	),
	(TEXTNS,u'tab') : (
	),
	(TEXTNS,u'table-count') : (
	),
	(TEXTNS,u'table-formula') : (
	),
	(TEXTNS,u'table-index') : (
		(TEXTNS,u'index-body'),
		(TEXTNS,u'table-index-source'),
	),
	(TEXTNS,u'table-index-entry-template') : (
		(TEXTNS,u'index-entry-chapter'),
		(TEXTNS,u'index-entry-page-number'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
		(TEXTNS,u'index-entry-text'),
	),
	(TEXTNS,u'table-index-source') : (
		(TEXTNS,u'index-title-template'),
		(TEXTNS,u'table-index-entry-template'),
	),
	(TEXTNS,u'table-of-content') : (
		(TEXTNS,u'index-body'),
		(TEXTNS,u'table-of-content-source'),
	),
	(TEXTNS,u'table-of-content-entry-template') : (
		(TEXTNS,u'index-entry-chapter'),
		(TEXTNS,u'index-entry-link-end'),
		(TEXTNS,u'index-entry-link-start'),
		(TEXTNS,u'index-entry-page-number'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
		(TEXTNS,u'index-entry-text'),
	),
	(TEXTNS,u'table-of-content-source') : (
		(TEXTNS,u'index-source-styles'),
		(TEXTNS,u'index-title-template'),
		(TEXTNS,u'table-of-content-entry-template'),
	),
	(TEXTNS,u'template-name') : (
	),
	(TEXTNS,u'text-input') : (
	),
	(TEXTNS,u'time') : (
	),
	(TEXTNS,u'title') : (
	),
	(TEXTNS,u'toc-mark') : (
	),
	(TEXTNS,u'toc-mark-end') : (
	),
	(TEXTNS,u'toc-mark-start') : (
	),
# allowed_children
	(TEXTNS,u'tracked-changes') : (
		(TEXTNS,u'changed-region'),
	),
	(TEXTNS,u'user-defined') : (
	),
	(TEXTNS,u'user-field-decl') : (
	),
	(TEXTNS,u'user-field-decls') : (
		(TEXTNS,u'user-field-decl'),
	),
	(TEXTNS,u'user-field-get') : (
	),
	(TEXTNS,u'user-field-input') : (
	),
	(TEXTNS,u'user-index') : (
		(TEXTNS,u'index-body'),
		(TEXTNS,u'user-index-source'),
	),
	(TEXTNS,u'user-index-entry-template') : (
		(TEXTNS,u'index-entry-chapter'),
		(TEXTNS,u'index-entry-page-number'),
		(TEXTNS,u'index-entry-span'),
		(TEXTNS,u'index-entry-tab-stop'),
		(TEXTNS,u'index-entry-text'),
	),
# allowed_children
	(TEXTNS,u'user-index-mark') : (
	),
	(TEXTNS,u'user-index-mark-end') : (
	),
	(TEXTNS,u'user-index-mark-start') : (
	),
	(TEXTNS,u'user-index-source') : (
		(TEXTNS,u'index-source-styles'),
		(TEXTNS,u'index-title-template'),
		(TEXTNS,u'user-index-entry-template'),
	),
	(TEXTNS,u'variable-decl') : (
	),
	(TEXTNS,u'variable-decls') : (
		(TEXTNS,u'variable-decl'),
	),
	(TEXTNS,u'variable-get') : (
	),
	(TEXTNS,u'variable-input') : (
	),
	(TEXTNS,u'variable-set') : (
	),
	(TEXTNS,u'word-count') : (
	),
}

#
# List of elements that allows text nodes
#
allows_text = (
	(CONFIGNS,u'config-item'),
	(DCNS,u'creator'),
	(DCNS,u'date'),
	(DCNS,u'description'),
	(DCNS,u'language'),
	(DCNS,u'subject'),
	(DCNS,u'title'),
# Completes Dublin Core start
#	(DCNS,'contributor'),
#	(DCNS,'coverage'),
#	(DCNS,'format'),
#	(DCNS,'identifier'),
#	(DCNS,'publisher'),
#	(DCNS,'relation'),
#	(DCNS,'rights'),
#	(DCNS,'source'),
#	(DCNS,'type'),
# Completes Dublin Core end
	(FORMNS,u'item'),
	(FORMNS,u'option'),
	(MATHNS,u'math'),
	(METANS,u'creation-date'),
	(METANS,u'date-string'),
	(METANS,u'editing-cycles'),
	(METANS,u'editing-duration'),
# allows_text
	(METANS,u'generator'),
	(METANS,u'initial-creator'),
	(METANS,u'keyword'),
	(METANS,u'print-date'),
	(METANS,u'printed-by'),
	(METANS,u'user-defined'),
	(NUMBERNS,u'currency-symbol'),
	(NUMBERNS,u'embedded-text'),
	(NUMBERNS,u'text'),
	(OFFICENS,u'binary-data'),
	(OFFICENS,u'script'),
	(PRESENTATIONNS,u'date-time-decl'),
	(PRESENTATIONNS,u'footer-decl'),
	(PRESENTATIONNS,u'header-decl'),
	(SVGNS,u'desc'),
	(SVGNS,u'title'),
	(TABLENS,u'desc'),
	(TABLENS,u'title'),
	(TEXTNS,u'a'),
	(TEXTNS,u'author-initials'),
	(TEXTNS,u'author-name'),
	(TEXTNS,u'bibliography-mark'),
	(TEXTNS,u'bookmark-ref'),
	(TEXTNS,u'chapter'),
	(TEXTNS,u'character-count'),
	(TEXTNS,u'conditional-text'),
	(TEXTNS,u'creation-date'),
	(TEXTNS,u'creation-time'),
	(TEXTNS,u'creator'),
	(TEXTNS,u'database-display'),
	(TEXTNS,u'database-name'),
	(TEXTNS,u'database-row-number'),
	(TEXTNS,u'date'),
	(TEXTNS,u'dde-connection'),
	(TEXTNS,u'description'),
	(TEXTNS,u'editing-cycles'),
	(TEXTNS,u'editing-duration'),
	(TEXTNS,u'execute-macro'),
	(TEXTNS,u'expression'),
	(TEXTNS,u'file-name'),
	(TEXTNS,u'h'),
	(TEXTNS,u'hidden-paragraph'),
	(TEXTNS,u'hidden-text'),
	(TEXTNS,u'image-count'),
# allows_text
	(TEXTNS,u'index-entry-span'),
	(TEXTNS,u'index-title-template'),
	(TEXTNS,u'initial-creator'),
	(TEXTNS,u'keywords'),
	(TEXTNS,u'linenumbering-separator'),
	(TEXTNS,u'measure'),
	(TEXTNS,u'meta'),
	(TEXTNS,u'meta-field'),
	(TEXTNS,u'modification-date'),
	(TEXTNS,u'modification-time'),
	(TEXTNS,u'note-citation'),
	(TEXTNS,u'note-continuation-notice-backward'),
	(TEXTNS,u'note-continuation-notice-forward'),
	(TEXTNS,u'note-ref'),
	(TEXTNS,u'number'),
	(TEXTNS,u'object-count'),
	(TEXTNS,u'p'),
	(TEXTNS,u'page-continuation'),
	(TEXTNS,u'page-count'),
	(TEXTNS,u'page-number'),
	(TEXTNS,u'page-variable-get'),
	(TEXTNS,u'page-variable-set'),
	(TEXTNS,u'paragraph-count'),
	(TEXTNS,u'placeholder'),
	(TEXTNS,u'print-date'),
	(TEXTNS,u'print-time'),
	(TEXTNS,u'printed-by'),
	(TEXTNS,u'reference-ref'),
	(TEXTNS,u'ruby-base'),
	(TEXTNS,u'ruby-text'),
# allows_text
	(TEXTNS,u'script'),
	(TEXTNS,u'sender-city'),
	(TEXTNS,u'sender-company'),
	(TEXTNS,u'sender-country'),
	(TEXTNS,u'sender-email'),
	(TEXTNS,u'sender-fax'),
	(TEXTNS,u'sender-firstname'),
	(TEXTNS,u'sender-initials'),
	(TEXTNS,u'sender-lastname'),
	(TEXTNS,u'sender-phone-private'),
	(TEXTNS,u'sender-phone-work'),
	(TEXTNS,u'sender-position'),
	(TEXTNS,u'sender-postal-code'),
	(TEXTNS,u'sender-state-or-province'),
	(TEXTNS,u'sender-street'),
	(TEXTNS,u'sender-title'),
	(TEXTNS,u'sequence'),
	(TEXTNS,u'sequence-ref'),
	(TEXTNS,u'sheet-name'),
# allows_text
	(TEXTNS,u'span'),
	(TEXTNS,u'subject'),
	(TEXTNS,u'table-count'),
	(TEXTNS,u'table-formula'),
	(TEXTNS,u'template-name'),
	(TEXTNS,u'text-input'),
	(TEXTNS,u'time'),
	(TEXTNS,u'title'),
	(TEXTNS,u'user-defined'),
	(TEXTNS,u'user-field-get'),
	(TEXTNS,u'user-field-input'),
	(TEXTNS,u'variable-get'),
	(TEXTNS,u'variable-input'),
	(TEXTNS,u'variable-set'),
	(TEXTNS,u'word-count'),
)

# Only the elements with at least one required attribute is listed

required_attributes = {
	(ANIMNS,u'animate'): (
		(SMILNS,u'attributeName'),
	),
	(ANIMNS,u'animateColor'): (
		(SMILNS,u'attributeName'),
	),
	(ANIMNS,u'animateMotion'): (
		(SMILNS,u'attributeName'),
	),
	(ANIMNS,u'animateTransform'): (
		(SVGNS,u'type'),
		(SMILNS,u'attributeName'),
	),
	(ANIMNS,u'command'): (
		(ANIMNS,u'command'),
	),
	(ANIMNS,u'param'): (
		(ANIMNS,u'name'),
		(ANIMNS,u'value'),
	),
	(ANIMNS,u'set'): (
		(SMILNS,u'attributeName'),
	),
# required_attributes
	(ANIMNS,u'transitionFilter'): (
		(SMILNS,u'type'),
	),
	(CHARTNS,u'axis'): (
		(CHARTNS,u'dimension'),
	),
	(CHARTNS,u'chart'): (
		(CHARTNS,u'class'),
	),
# required_attributes
	(CHARTNS,u'error-indicator'): (
		(CHARTNS,u'dimension'),
	),
	(CHARTNS,u'symbol-image'): (
		(XLINKNS,u'href'),
	),
	(CONFIGNS,u'config-item'): (
		(CONFIGNS,u'type'),
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-map-indexed'): (
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-map-named'): (
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-set'): (
		(CONFIGNS,u'name'),
	),
# required_attributes
	(NUMBERNS,u'boolean-style'): (
		(STYLENS,u'name'),
	),
	(NUMBERNS,u'currency-style'): (
		(STYLENS,u'name'),
	),
	(NUMBERNS,u'date-style'): (
		(STYLENS,u'name'),
	),
	(NUMBERNS,u'embedded-text'): (
		(NUMBERNS,u'position'),
	),
	(NUMBERNS,u'number-style'): (
		(STYLENS,u'name'),
	),
	(NUMBERNS,u'percentage-style'): (
		(STYLENS,u'name'),
	),
	(NUMBERNS,u'text-style'): (
		(STYLENS,u'name'),
	),
	(NUMBERNS,u'time-style'): (
		(STYLENS,u'name'),
	),
	(DR3DNS,u'extrude'): (
		(SVGNS,u'd'),
		(SVGNS,u'viewBox'),
	),
	(DR3DNS,u'light'): (
		(DR3DNS,u'direction'),
	),
	(DR3DNS,u'rotate'): (
		(SVGNS,u'viewBox'),
		(SVGNS,u'd'),
	),
# required_attributes
	(DRAWNS,u'a'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(DRAWNS,u'area-circle'): (
		(SVGNS,u'cy'),
		(SVGNS,u'cx'),
		(SVGNS,u'r'),
	),
	(DRAWNS,u'area-polygon'): (
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'points'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(SVGNS,u'viewBox'),
	),
	(DRAWNS,u'area-rectangle'): (
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
	),
# required_attributes
	(DRAWNS,u'connector'): (
		(SVGNS,u'viewBox'),
	),
	(DRAWNS,u'contour-path'): (
		(DRAWNS,u'recreate-on-edit'),
		(SVGNS,u'viewBox'),
		(SVGNS,u'd'),
	),
	(DRAWNS,u'contour-polygon'): (
		(DRAWNS,u'points'),
		(DRAWNS,u'recreate-on-edit'),
		(SVGNS,u'viewBox'),
	),
	(DRAWNS,u'control'): (
		(DRAWNS,u'control'),
	),
	(DRAWNS,u'fill-image'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(DRAWNS,u'name'),
	),
	(DRAWNS,u'floating-frame'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(DRAWNS,u'glue-point'): (
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(DRAWNS,u'id'),
		(DRAWNS,u'escape-direction'),
	),
# required_attributes
	(DRAWNS,u'gradient'): (
		(DRAWNS,u'style'),
	),
	(DRAWNS,u'handle'): (
		(DRAWNS,u'handle-position'),
	),
	(DRAWNS,u'hatch'): (
		(DRAWNS,u'style'),
		(DRAWNS,u'name'),
	),
	(DRAWNS,u'layer'): (
		(DRAWNS,u'name'),
	),
	(DRAWNS,u'line'): (
		(SVGNS,u'y1'),
		(SVGNS,u'x2'),
		(SVGNS,u'x1'),
		(SVGNS,u'y2'),
	),
	(DRAWNS,u'marker'): (
		(SVGNS,u'd'),
		(DRAWNS,u'name'),
		(SVGNS,u'viewBox'),
	),
	(DRAWNS,u'measure'): (
		(SVGNS,u'y1'),
		(SVGNS,u'x2'),
		(SVGNS,u'x1'),
		(SVGNS,u'y2'),
	),
	(DRAWNS,u'opacity'): (
		(DRAWNS,u'style'),
	),
	(DRAWNS,u'page'): (
		(DRAWNS,u'master-page-name'),
	),
	(DRAWNS,u'path'): (
		(SVGNS,u'd'),
		(SVGNS,u'viewBox'),
	),
	(DRAWNS,u'plugin'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(DRAWNS,u'polygon'): (
		(DRAWNS,u'points'),
		(SVGNS,u'viewBox'),
	),
# required_attributes
	(DRAWNS,u'polyline'): (
		(DRAWNS,u'points'),
		(SVGNS,u'viewBox'),
	),
	(DRAWNS,u'regular-polygon'): (
		(DRAWNS,u'corners'),
	),
	(DRAWNS,u'stroke-dash'): (
		(DRAWNS,u'name'),
	),
	(FORMNS,u'button'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'checkbox'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'combobox'): (
		(XMLNS,u'id'),
	),
# required_attributes
	(FORMNS,u'connection-resource'): (
		(XLINKNS,u'href'),
	),
	(FORMNS,u'date'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'file'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'fixed-text'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'formatted-text'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'frame'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'generic-control'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'grid'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'hidden'): (
		(XMLNS,u'id'),
	),
# required_attributes
	(FORMNS,u'image'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'image-frame'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'list-property'): (
		(FORMNS,u'property-name'),
	),
	(FORMNS,u'list-value'): (
		(OFFICENS,u'string-value'),
	),
	(FORMNS,u'listbox'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'number'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'password'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'property'): (
		(FORMNS,u'property-name'),
	),
	(FORMNS,u'radio'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'text'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'textarea'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'time'): (
		(XMLNS,u'id'),
	),
	(FORMNS,u'value-range'): (
		(XMLNS,u'id'),
	),
	(MANIFESTNS,u'algorithm') : (
		(MANIFESTNS,u'algorithm-name'),
		(MANIFESTNS,u'initialisation-vector'),
	),
	(MANIFESTNS,u'encryption-data') : (
		(MANIFESTNS,u'checksum-type'),
		(MANIFESTNS,u'checksum'),
	),
	(MANIFESTNS,u'file-entry') : (
		(MANIFESTNS,u'full-path'),
		(MANIFESTNS,u'media-type'),
	),
	(MANIFESTNS,u'key-derivation') : (
		(MANIFESTNS,u'key-derivation-name'),
		(MANIFESTNS,u'salt'),
		(MANIFESTNS,u'iteration-count'),
	),
# required_attributes
	(METANS,u'template'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(METANS,u'user-defined'): (
		(METANS,u'name'),
	),
# required_attributes
	(OFFICENS,u'annotation-end'): (
		(OFFICENS,u'name'),
	),
	(OFFICENS,u'dde-source'): (
		(OFFICENS,u'dde-topic'),
		(OFFICENS,u'dde-application'),
		(OFFICENS,u'dde-item'),
	),
	(OFFICENS,u'document'): (
		(OFFICENS,u'mimetype'),
		(OFFICENS,u'version'),
	),
# required_attributes
	(OFFICENS,u'document-content'): (
		(OFFICENS,u'version'),
	),
# required_attributes
	(OFFICENS,u'document-meta'): (
		(OFFICENS,u'version'),
	),
# required_attributes
	(OFFICENS,u'document-settings'): (
		(OFFICENS,u'version'),
	),
# required_attributes
	(OFFICENS,u'document-styles'): (
		(OFFICENS,u'version'),
	),
	(OFFICENS,u'script'): (
		(SCRIPTNS,u'language'),
	),
	(PRESENTATIONNS,u'date-time-decl'): (
		(PRESENTATIONNS,u'source'),
		(PRESENTATIONNS,u'name'),
	),
	(PRESENTATIONNS,u'dim'): (
		(DRAWNS,u'color'),
		(DRAWNS,u'shape-id'),
	),
# required_attributes
	(PRESENTATIONNS,u'event-listener'): (
		(PRESENTATIONNS,u'action'),
		(SCRIPTNS,u'event-name'),
	),
	(PRESENTATIONNS,u'footer-decl'): (
		(PRESENTATIONNS,u'name'),
	),
	(PRESENTATIONNS,u'header-decl'): (
		(PRESENTATIONNS,u'name'),
	),
	(PRESENTATIONNS,u'hide-shape'): (
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'hide-text'): (
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'placeholder'): (
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(SVGNS,u'height'),
		(PRESENTATIONNS,u'object'),
		(SVGNS,u'width'),
	),
	(PRESENTATIONNS,u'play'): (
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'show'): (
		(PRESENTATIONNS,u'name'),
		(PRESENTATIONNS,u'pages'),
	),
	(PRESENTATIONNS,u'show-shape'): (
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'show-text'): (
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'sound'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(SCRIPTNS,u'event-listener'): (
		(SCRIPTNS,u'language'),
		(SCRIPTNS,u'event-name'),
	),
	(STYLENS,u'column'): (
		(STYLENS,u'rel-width'),
	),
# required_attributes
	(STYLENS,u'column-sep'): (
		(STYLENS,u'width'),
	),
	(STYLENS,u'columns'): (
		(FONS,u'column-count'),
	),
	(STYLENS,u'font-face'): (
		(STYLENS,u'name'),
	),
	(STYLENS,u'handout-master'): (
		(STYLENS,u'page-layout-name'),
	),
	(STYLENS,u'map'): (
		(STYLENS,u'apply-style-name'),
		(STYLENS,u'condition'),
	),
# required_attributes
	(STYLENS,u'list-level-label-alignment'): (
		(TEXTNS,u'label-followed-by'),
	),
	(STYLENS,u'master-page'): (
		(STYLENS,u'page-layout-name'),
		(STYLENS,u'name'),
	),
	(STYLENS,u'page-layout'): (
		(STYLENS,u'name'),
	),
	(STYLENS,u'presentation-page-layout'): (
		(STYLENS,u'name'),
	),
	(STYLENS,u'style'): (
		(STYLENS,u'name'),
	),
	(STYLENS,u'tab-stop'): (
		(STYLENS,u'position'),
	),
	(SVGNS,u'definition-src'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(SVGNS,u'font-face-uri'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(SVGNS,u'linearGradient'): (
		(DRAWNS,u'name'),
	),
	(SVGNS,u'radialGradient'): (
		(DRAWNS,u'name'),
	),
	(SVGNS,u'stop'): (
		(SVGNS,u'offset'),
	),
# required_attributes
	(TABLENS,u'background'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'body'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'cell-address'): (
		(TABLENS,u'column'),
		(TABLENS,u'table'),
		(TABLENS,u'row'),
	),
	(TABLENS,u'cell-content-change'): (
		(TABLENS,u'id'),
	),
	(TABLENS,u'cell-range-source'): (
		(TABLENS,u'last-row-spanned'),
		(TABLENS,u'last-column-spanned'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'consolidation'): (
		(TABLENS,u'function'),
		(TABLENS,u'source-cell-range-addresses'),
		(TABLENS,u'target-cell-address'),
	),
	(TABLENS,u'content-validation'): (
		(TABLENS,u'name'),
	),
	(TABLENS,u'data-pilot-display-info'): (
		(TABLENS,u'member-count'),
		(TABLENS,u'data-field'),
		(TABLENS,u'enabled'),
		(TABLENS,u'display-member-mode'),
	),
# required_attributes
	(TABLENS,u'data-pilot-field'): (
		(TABLENS,u'source-field-name'),
	),
	(TABLENS,u'data-pilot-field-reference'): (
		(TABLENS,u'field-name'),
		(TABLENS,u'type'),
	),
	(TABLENS,u'data-pilot-group'): (
		(TABLENS,u'name'),
	),
	(TABLENS,u'data-pilot-group-member'): (
		(TABLENS,u'name'),
	),
	(TABLENS,u'data-pilot-groups'): (
		(TABLENS,u'source-field-name'),
		(TABLENS,u'step'),
		(TABLENS,u'grouped-by'),
	),
	(TABLENS,u'data-pilot-layout-info'): (
		(TABLENS,u'add-empty-lines'),
		(TABLENS,u'layout-mode'),
	),
	(TABLENS,u'data-pilot-member'): (
		(TABLENS,u'name'),
	),
	(TABLENS,u'data-pilot-sort-info'): (
		(TABLENS,u'order'),
	),
	(TABLENS,u'data-pilot-subtotal'): (
		(TABLENS,u'function'),
	),
	(TABLENS,u'data-pilot-table'): (
		(TABLENS,u'target-range-address'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'database-range'): (
		(TABLENS,u'target-range-address'),
	),
# required_attributes
	(TABLENS,u'database-source-query'): (
		(TABLENS,u'query-name'),
		(TABLENS,u'database-name'),
	),
	(TABLENS,u'database-source-sql'): (
		(TABLENS,u'database-name'),
		(TABLENS,u'sql-statement'),
	),
	(TABLENS,u'database-source-table'): (
		(TABLENS,u'database-table-name'),
		(TABLENS,u'database-name'),
	),
	(TABLENS,u'deletion'): (
		(TABLENS,u'position'),
		(TABLENS,u'type'),
		(TABLENS,u'id'),
	),
	(TABLENS,u'dependency'): (
		(TABLENS,u'id'),
	),
# required_attributes
	(TABLENS,u'even-columns'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'even-rows'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'filter-condition'): (
		(TABLENS,u'operator'),
		(TABLENS,u'field-number'),
		(TABLENS,u'value'),
	),
# required_attributes
	(TABLENS,u'filter-set-item'): (
		(TABLENS,u'value'),
	),
	(TABLENS,u'first-column'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'first-row'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'insertion'): (
		(TABLENS,u'position'),
		(TABLENS,u'type'),
		(TABLENS,u'id'),
	),
	(TABLENS,u'insertion-cut-off'): (
		(TABLENS,u'position'),
		(TABLENS,u'id'),
	),
# required_attributes
	(TABLENS,u'label-range'): (
		(TABLENS,u'label-cell-range-address'),
		(TABLENS,u'data-cell-range-address'),
		(TABLENS,u'orientation'),
	),
	(TABLENS,u'last-column'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'last-row'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'movement'): (
		(TABLENS,u'id'),
	),
	(TABLENS,u'named-expression'): (
		(TABLENS,u'expression'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'named-range'): (
		(TABLENS,u'name'),
		(TABLENS,u'cell-range-address'),
	),
	(TABLENS,u'odd-columns'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'odd-rows'): (
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'operation'): (
		(TABLENS,u'index'),
		(TABLENS,u'name'),
	),
# required_attributes
	(TABLENS,u'scenario'): (
		(TABLENS,u'is-active'),
		(TABLENS,u'scenario-ranges'),
	),
	(TABLENS,u'sort-by'): (
		(TABLENS,u'field-number'),
	),
	(TABLENS,u'source-cell-range'): (
		(TABLENS,u'cell-range-address'),
	),
	(TABLENS,u'source-service'): (
		(TABLENS,u'source-name'),
		(TABLENS,u'object-name'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'subtotal-field'): (
		(TABLENS,u'function'),
		(TABLENS,u'field-number'),
	),
	(TABLENS,u'subtotal-rule'): (
		(TABLENS,u'group-by-field-number'),
	),
	(TABLENS,u'table-source'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(TABLENS,u'table-template'): (
#		(TABLENS,u'last-row-end-column'), # Deprecated
#		(TABLENS,u'first-row-end-column'), # Deprecated
		(TABLENS,u'name'),
#		(TABLENS,u'last-row-start-column'), # Deprecated
#		(TABLENS,u'first-row-start-column'), # Deprecated
	),
	(TEXTNS,u'a'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
# required_attributes
	(TEXTNS,u'alphabetical-index'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'alphabetical-index-auto-mark-file'): (
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(TEXTNS,u'alphabetical-index-entry-template'): (
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'alphabetical-index-mark'): (
		(TEXTNS,u'string-value'),
	),
	(TEXTNS,u'alphabetical-index-mark-end'): (
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'alphabetical-index-mark-start'): (
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'bibliography'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'bibliography-entry-template'): (
		(TEXTNS,u'style-name'),
		(TEXTNS,u'bibliography-type'),
	),
	(TEXTNS,u'bibliography-mark'): (
		(TEXTNS,u'bibliography-type'),
	),
	(TEXTNS,u'bookmark'): (
		(TEXTNS,u'name'),
	),
# required_attributes
	(TEXTNS,u'bookmark-end'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'bookmark-start'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'change'): (
		(TEXTNS,u'change-id'),
	),
	(TEXTNS,u'change-end'): (
		(TEXTNS,u'change-id'),
	),
	(TEXTNS,u'change-start'): (
		(TEXTNS,u'change-id'),
	),
	(TEXTNS,u'changed-region'): (
		(XMLNS,u'id'),
	),
	(TEXTNS,u'chapter'): (
		(TEXTNS,u'display'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'conditional-text'): (
		(TEXTNS,u'string-value-if-true'),
		(TEXTNS,u'string-value-if-false'),
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'database-display'): (
		(TEXTNS,u'column-name'),
		(TEXTNS,u'table-name'),
	),
	(TEXTNS,u'database-name'): (
		(TEXTNS,u'table-name'),
	),
	(TEXTNS,u'database-next'): (
		(TEXTNS,u'table-name'),
	),
	(TEXTNS,u'database-row-number'): (
		(TEXTNS,u'table-name'),
	),
	(TEXTNS,u'database-row-select'): (
		(TEXTNS,u'table-name'),
	),
	(TEXTNS,u'dde-connection'): (
		(TEXTNS,u'connection-name'),
	),
# required_attributes
	(TEXTNS,u'dde-connection-decl'): (
		(OFFICENS,u'dde-topic'),
		(OFFICENS,u'dde-application'),
		(OFFICENS,u'name'),
		(OFFICENS,u'dde-item'),
	),
	(TEXTNS,u'h'): (
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'hidden-paragraph'): (
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'hidden-text'): (
		(TEXTNS,u'string-value'),
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'illustration-index'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'illustration-index-entry-template'): (
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-entry-bibliography'): (
		(TEXTNS,u'bibliography-data-field'),
	),
	(TEXTNS,u'index-source-style'): (
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-source-styles'): (
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'index-title'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'list-level-style-bullet'): (
		(TEXTNS,u'bullet-char'),
		(TEXTNS,u'level'),
	),
	(TEXTNS,u'list-level-style-image'): (
		(TEXTNS,u'level'),
	),
	(TEXTNS,u'list-level-style-number'): (
		(TEXTNS,u'level'),
	),
	(TEXTNS,u'list-style'): (
		(STYLENS,u'name'),
	),
	(TEXTNS,u'meta-field'): (
		(XMLNS,u'id'),
	),
# required_attributes
	(TEXTNS,u'measure'): (
		(TEXTNS,u'kind'),
	),
	(TEXTNS,u'note'): (
		(TEXTNS,u'note-class'),
	),
	(TEXTNS,u'note-ref'): (
		(TEXTNS,u'note-class'),
	),
	(TEXTNS,u'notes-configuration'): (
		(TEXTNS,u'note-class'),
	),
# required_attributes
	(TEXTNS,u'numbered-paragraph'): (
		(TEXTNS,u'list-id'),
	),
	(TEXTNS,u'object-index'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'object-index-entry-template'): (
		(TEXTNS,u'style-name'),
	),
# required_attributes
	(TEXTNS,u'outline-style'): (
		(STYLENS,u'name'),
	),
	(TEXTNS,u'outline-level-style'): (
		(TEXTNS,u'level'),
	),
	(TEXTNS,u'page'): (
		(TEXTNS,u'master-page-name'),
	),
	(TEXTNS,u'page-continuation'): (
		(TEXTNS,u'select-page'),
	),
	(TEXTNS,u'placeholder'): (
		(TEXTNS,u'placeholder-type'),
	),
	(TEXTNS,u'reference-mark'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'reference-mark-end'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'reference-mark-start'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'section'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'sequence'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'sequence-decl'): (
		(TEXTNS,u'display-outline-level'),
		(TEXTNS,u'name'),
	),
# required_attributes
	(TEXTNS,u'sort-key'): (
		(TEXTNS,u'key'),
	),
	(TEXTNS,u'table-index'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'table-index-entry-template'): (
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'table-of-content'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'table-of-content-entry-template'): (
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'toc-mark'): (
		(TEXTNS,u'string-value'),
	),
	(TEXTNS,u'toc-mark-end'): (
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'toc-mark-start'): (
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'user-defined'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'user-field-decl'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'user-field-get'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'user-field-input'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'user-index'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'user-index-entry-template'): (
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
	),
# required_attributes
	(TEXTNS,u'user-index-mark'): (
		(TEXTNS,u'index-name'),
		(TEXTNS,u'string-value'),
	),
	(TEXTNS,u'user-index-mark-end'): (
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'user-index-mark-start'): (
		(TEXTNS,u'index-name'),
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'user-index-source'): (
		(TEXTNS,u'index-name'),
	),
	(TEXTNS,u'variable-decl'): (
		(TEXTNS,u'name'),
		(OFFICENS,u'value-type'),
	),
	(TEXTNS,u'variable-get'): (
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'variable-input'): (
		(TEXTNS,u'name'),
		(OFFICENS,u'value-type'),
	),
	(TEXTNS,u'variable-set'): (
		(TEXTNS,u'name'),
	),
}

# Empty list means the element has no allowed attributes
# None means anything goes

allowed_attributes = {
	(DCNS,u'creator'):(
	),
	(DCNS,u'date'):(
	),
	(DCNS,u'description'):(
	),
	(DCNS,u'language'):(
	),
	(DCNS,u'subject'):(
	),
	(DCNS,u'title'):(
	),
# Completes Dublin Core start
#	(DCNS,'contributor') : (
#	),
#	(DCNS,'coverage') : (
#	),
#	(DCNS,'format') : (
#	),
#	(DCNS,'identifier') : (
#	),
#	(DCNS,'publisher') : (
#	),
#	(DCNS,'relation') : (
#	),
#	(DCNS,'rights') : (
#	),
#	(DCNS,'source') : (
#	),
#	(DCNS,'type') : (
#	),
# Completes Dublin Core end
	(MATHNS,u'math'): None,
	(XFORMSNS,u'model'): None,
# allowed_attributes
	(ANIMNS,u'animate'):(
		(ANIMNS,u'formula'),
		(ANIMNS,u'sub-item'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'accumulate'),
		(SMILNS,u'additive'),
		(SMILNS,u'attributeName'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'by'),
		(SMILNS,u'calcMode'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'from'),
		(SMILNS,u'keySplines'),
		(SMILNS,u'keyTimes'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'targetElement'),
		(SMILNS,u'to'),
		(SMILNS,u'values'),
	),
# allowed_attributes
	(ANIMNS,u'animateColor'):(
		(ANIMNS,u'color-interpolation'),
		(ANIMNS,u'color-interpolation-direction'),
		(ANIMNS,u'formula'),
		(ANIMNS,u'sub-item'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'accumulate'),
		(SMILNS,u'additive'),
		(SMILNS,u'attributeName'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'by'),
		(SMILNS,u'calcMode'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'from'),
		(SMILNS,u'keySplines'),
		(SMILNS,u'keyTimes'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'targetElement'),
		(SMILNS,u'to'),
		(SMILNS,u'values'),
	),
# allowed_attributes
	(ANIMNS,u'animateMotion'):(
		(ANIMNS,u'formula'),
		(ANIMNS,u'sub-item'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'accumulate'),
		(SMILNS,u'additive'),
		(SMILNS,u'attributeName'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'by'),
		(SMILNS,u'calcMode'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'from'),
		(SMILNS,u'keySplines'),
		(SMILNS,u'keyTimes'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'targetElement'),
		(SMILNS,u'to'),
		(SMILNS,u'values'),
		(SVGNS,u'origin'),
		(SVGNS,u'path'),
	),
# allowed_attributes
	(ANIMNS,u'animateTransform'):(
		(ANIMNS,u'formula'),
		(ANIMNS,u'sub-item'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'accumulate'),
		(SMILNS,u'additive'),
		(SMILNS,u'attributeName'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'by'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'from'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'targetElement'),
		(SMILNS,u'to'),
		(SMILNS,u'values'),
		(SVGNS,u'type'),
	),
# allowed_attributes
	(ANIMNS,u'audio'):(
		(ANIMNS,u'audio-level'),
		(ANIMNS,u'id'),
		(PRESENTATIONNS,u'group-id'),
		(PRESENTATIONNS,u'master-element'),
		(PRESENTATIONNS,u'node-type'),
		(PRESENTATIONNS,u'preset-class'),
		(PRESENTATIONNS,u'preset-id'),
		(PRESENTATIONNS,u'preset-sub-type'),
		(SMILNS,u'begin'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(XLINKNS,u'href'),
		(XMLNS,u'id'),
	),
	(ANIMNS,u'command'):(
		(PRESENTATIONNS,u'node-type'),
		(SMILNS,u'begin'),
		(SMILNS,u'end'),
		(PRESENTATIONNS,u'group-id'),
		(PRESENTATIONNS,u'preset-class'),
		(PRESENTATIONNS,u'preset-id'),
		(ANIMNS,u'sub-item'),
		(ANIMNS,u'command'),
		(PRESENTATIONNS,u'preset-sub-type'),
		(SMILNS,u'targetElement'),
		(ANIMNS,u'id'),
		(PRESENTATIONNS,u'master-element'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(ANIMNS,u'iterate'):(
		(ANIMNS,u'id'),
		(ANIMNS,u'iterate-interval'),
		(ANIMNS,u'iterate-type'),
		(ANIMNS,u'sub-item'),
		(PRESENTATIONNS,u'group-id'),
		(PRESENTATIONNS,u'master-element'),
		(PRESENTATIONNS,u'node-type'),
		(PRESENTATIONNS,u'preset-class'),
		(PRESENTATIONNS,u'preset-id'),
		(PRESENTATIONNS,u'preset-sub-type'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'endsync'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'targetElement'),
		(XMLNS,u'id'),
	),
	(ANIMNS,u'par'):(
		(PRESENTATIONNS,u'node-type'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'begin'),
		(SMILNS,u'end'),
		(PRESENTATIONNS,u'group-id'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'autoReverse'),
		(PRESENTATIONNS,u'preset-class'),
		(SMILNS,u'fillDefault'),
		(PRESENTATIONNS,u'preset-id'),
		(PRESENTATIONNS,u'preset-sub-type'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'endsync'),
		(SMILNS,u'dur'),
		(SMILNS,u'fill'),
		(ANIMNS,u'id'),
		(SMILNS,u'restart'),
		(PRESENTATIONNS,u'master-element'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(ANIMNS,u'param'):(
		(ANIMNS,u'name'),
		(ANIMNS,u'value'),
	),
	(ANIMNS,u'seq'):(
		(ANIMNS,u'id'),
		(PRESENTATIONNS,u'group-id'),
		(PRESENTATIONNS,u'master-element'),
		(PRESENTATIONNS,u'node-type'),
		(PRESENTATIONNS,u'preset-class'),
		(PRESENTATIONNS,u'preset-id'),
		(PRESENTATIONNS,u'preset-sub-type'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'endsync'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(XMLNS,u'id'),
	),
	(ANIMNS,u'set'):(
		(ANIMNS,u'sub-item'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'accumulate'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'additive'),
		(SMILNS,u'attributeName'),
		(SMILNS,u'begin'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'targetElement'),
		(SMILNS,u'to'),

	),
# allowed_attributes
	(ANIMNS,u'transitionFilter'):(
		(ANIMNS,u'formula'),
		(ANIMNS,u'sub-item'),
		(SMILNS,u'accelerate'),
		(SMILNS,u'accumulate'),
		(SMILNS,u'additive'),
		(SMILNS,u'autoReverse'),
		(SMILNS,u'begin'),
		(SMILNS,u'by'),
		(SMILNS,u'calcMode'),
		(SMILNS,u'decelerate'),
		(SMILNS,u'direction'),
		(SMILNS,u'dur'),
		(SMILNS,u'end'),
		(SMILNS,u'fadeColor'),
		(SMILNS,u'fill'),
		(SMILNS,u'fillDefault'),
		(SMILNS,u'from'),
		(SMILNS,u'mode'),
		(SMILNS,u'repeatCount'),
		(SMILNS,u'repeatDur'),
		(SMILNS,u'restart'),
		(SMILNS,u'restartDefault'),
		(SMILNS,u'subtype'),
		(SMILNS,u'targetElement'),
		(SMILNS,u'to'),
		(SMILNS,u'type'),
		(SMILNS,u'values'),

	),
# allowed_attributes
	(CHARTNS,u'axis'):(
		(CHARTNS,u'style-name'),
		(CHARTNS,u'dimension'),
		(CHARTNS,u'name'),
	),
	(CHARTNS,u'categories'):(
		(TABLENS,u'cell-range-address'),
	),
	(CHARTNS,u'chart'):(
		(CHARTNS,u'class'),
		(CHARTNS,u'column-mapping'),
		(CHARTNS,u'row-mapping'),
		(CHARTNS,u'style-name'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(XMLNS,u'id'),
	),
	(CHARTNS,u'data-label'):(
		(CHARTNS,u'style-name'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
	),
	(CHARTNS,u'data-point'):(
		(CHARTNS,u'repeated'),
		(CHARTNS,u'style-name'),
		(XMLNS,u'id'),
	),
	(CHARTNS,u'domain'):(
		(TABLENS,u'cell-range-address'),
	),
# allowed_attributes
	(CHARTNS,u'equation'):(
		(CHARTNS,u'automatic-content'),
		(CHARTNS,u'display-equation'),
		(CHARTNS,u'display-r-square'),
		(CHARTNS,u'style-name'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
	),
	(CHARTNS,u'error-indicator'):(
		(CHARTNS,u'dimension'),
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'floor'):(
		(SVGNS,u'width'),
		(CHARTNS,u'style-name'),
	),
# allowed_attributes
	(CHARTNS,u'footer'):(
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'cell-range'),
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'grid'):(
		(CHARTNS,u'style-name'),
		(CHARTNS,u'class'),
	),
# allowed_attributes
	(CHARTNS,u'label-separator'):(
	),
	(CHARTNS,u'legend'):(
		(CHARTNS,u'legend-align'),
		(STYLENS,u'legend-expansion-aspect-ratio'),
		(STYLENS,u'legend-expansion'),
		(CHARTNS,u'legend-position'),
		(CHARTNS,u'style-name'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
	),
	(CHARTNS,u'mean-value'):(
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'plot-area'):(
		(CHARTNS,u'data-source-has-labels'),
		(CHARTNS,u'style-name'),
		(DR3DNS,u'ambient-color'),
		(DR3DNS,u'distance'),
		(DR3DNS,u'focal-length'),
		(DR3DNS,u'lighting-mode'),
		(DR3DNS,u'projection'),
		(DR3DNS,u'shade-mode'),
		(DR3DNS,u'shadow-slant'),
		(DR3DNS,u'transform'),
		(DR3DNS,u'vpn'),
		(DR3DNS,u'vrp'),
		(DR3DNS,u'vup'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'cell-range-address'),
		(XMLNS,u'id'),
	),
	(CHARTNS,u'regression-curve'):(
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'series'):(
		(CHARTNS,u'style-name'),
		(CHARTNS,u'attached-axis'),
		(CHARTNS,u'values-cell-range-address'),
		(CHARTNS,u'label-cell-address'),
		(CHARTNS,u'class'),
		(XMLNS,u'id'),
	),
	(CHARTNS,u'stock-gain-marker'):(
		(CHARTNS,u'style-name'),
	),
# allowed_attributes
	(CHARTNS,u'stock-loss-marker'):(
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'stock-range-line'):(
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'subtitle'):(
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'cell-range'),
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'symbol-image'):(
		(XLINKNS,u'href'),
	),
	(CHARTNS,u'title'):(
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'cell-range'),
		(CHARTNS,u'style-name'),
	),
	(CHARTNS,u'wall'):(
		(SVGNS,u'width'),
		(CHARTNS,u'style-name'),
	),
	(CONFIGNS,u'config-item'):(
		(CONFIGNS,u'type'),
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-map-entry'):(
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-map-indexed'):(
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-map-named'):(
		(CONFIGNS,u'name'),
	),
	(CONFIGNS,u'config-item-set'):(
		(CONFIGNS,u'name'),
	),
# allowed_attributes
	(NUMBERNS,u'am-pm'):(
	),
	(NUMBERNS,u'boolean'):(
	),
	(NUMBERNS,u'boolean-style'):(
		(NUMBERNS,u'country'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'transliteration-country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-language'),
		(NUMBERNS,u'transliteration-style'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
		(STYLENS,u'volatile'),
	),
	(NUMBERNS,u'currency-style'):(
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'automatic-order'),
		(NUMBERNS,u'country'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'transliteration-country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-language'),
		(NUMBERNS,u'transliteration-style'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
		(STYLENS,u'volatile'),
	),
	(NUMBERNS,u'currency-symbol'):(
		(NUMBERNS,u'country'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
	),
# allowed_attributes
	(NUMBERNS,u'date-style'):(
		(NUMBERNS,u'automatic-order'),
		(NUMBERNS,u'country'),
		(NUMBERNS,u'format-source'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'transliteration-country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-language'),
		(NUMBERNS,u'transliteration-style'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
		(STYLENS,u'volatile'),
	),
	(NUMBERNS,u'day'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'calendar'),
	),
	(NUMBERNS,u'day-of-week'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'calendar'),
	),
	(NUMBERNS,u'embedded-text'):(
		(NUMBERNS,u'position'),
	),
	(NUMBERNS,u'era'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'calendar'),
	),
	(NUMBERNS,u'fraction'):(
		(NUMBERNS,u'grouping'),
		(NUMBERNS,u'min-denominator-digits'),
		(NUMBERNS,u'min-numerator-digits'),
		(NUMBERNS,u'min-integer-digits'),
		(NUMBERNS,u'denominator-value'),
	),
	(NUMBERNS,u'hours'):(
		(NUMBERNS,u'style'),
	),
# allowed_attributes
	(NUMBERNS,u'minutes'):(
		(NUMBERNS,u'style'),
	),
	(NUMBERNS,u'month'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'calendar'),
		(NUMBERNS,u'possessive-form'),
		(NUMBERNS,u'textual'),
	),
	(NUMBERNS,u'number'):(
		(NUMBERNS,u'display-factor'),
		(NUMBERNS,u'decimal-places'),
		(NUMBERNS,u'decimal-replacement'),
		(NUMBERNS,u'min-integer-digits'),
		(NUMBERNS,u'grouping'),
	),
	(NUMBERNS,u'number-style'):(
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'transliteration-language'),
		(STYLENS,u'name'),
		(STYLENS,u'display-name'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-style'),
		(STYLENS,u'volatile'),
		(NUMBERNS,u'transliteration-country'),
	),
# allowed_attributes
	(NUMBERNS,u'percentage-style'):(
		(NUMBERNS,u'country'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'transliteration-country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-language'),
		(NUMBERNS,u'transliteration-style'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
		(STYLENS,u'volatile'),
	),
	(NUMBERNS,u'quarter'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'calendar'),
	),
	(NUMBERNS,u'scientific-number'):(
		(NUMBERNS,u'min-exponent-digits'),
		(NUMBERNS,u'decimal-places'),
		(NUMBERNS,u'min-integer-digits'),
		(NUMBERNS,u'grouping'),
	),
	(NUMBERNS,u'seconds'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'decimal-places'),
	),
	(NUMBERNS,u'text'):(
	),
	(NUMBERNS,u'text-content'):(
	),
	(NUMBERNS,u'text-style'):(
		(NUMBERNS,u'country'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'transliteration-country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-language'),
		(NUMBERNS,u'transliteration-style'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
		(STYLENS,u'volatile'),
	),
	(NUMBERNS,u'time-style'):(
		(NUMBERNS,u'country'),
		(NUMBERNS,u'format-source'),
		(NUMBERNS,u'language'),
		(NUMBERNS,u'rfc-language-tag'),
		(NUMBERNS,u'script'),
		(NUMBERNS,u'title'),
		(NUMBERNS,u'transliteration-country'),
		(NUMBERNS,u'transliteration-format'),
		(NUMBERNS,u'transliteration-language'),
		(NUMBERNS,u'transliteration-style'),
		(NUMBERNS,u'truncate-on-overflow'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
		(STYLENS,u'volatile'),
	),
	(NUMBERNS,u'week-of-year'):(
		(NUMBERNS,u'calendar'),
	),
	(NUMBERNS,u'year'):(
		(NUMBERNS,u'style'),
		(NUMBERNS,u'calendar'),
	),
	(DR3DNS,u'cube'):(
		(DR3DNS,u'min-edge'),
		(DR3DNS,u'max-edge'),
		(DRAWNS,u'layer'),
		(DR3DNS,u'transform'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(DRAWNS,u'id'),
		(XMLNS,u'id'),
	),
	(DR3DNS,u'extrude'):(
		(DRAWNS,u'layer'),
		(SVGNS,u'd'),
		(DR3DNS,u'transform'),
		(SVGNS,u'viewBox'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(DRAWNS,u'id'),
		(XMLNS,u'id'),
	),
	(DR3DNS,u'light'):(
		(DR3DNS,u'diffuse-color'),
		(DR3DNS,u'direction'),
		(DR3DNS,u'specular'),
		(DR3DNS,u'enabled'),
	),
	(DR3DNS,u'rotate'):(
		(DRAWNS,u'layer'),
		(SVGNS,u'd'),
		(DR3DNS,u'transform'),
		(SVGNS,u'viewBox'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(DRAWNS,u'id'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DR3DNS,u'scene'):(
		(DR3DNS,u'ambient-color'),
		(DR3DNS,u'distance'),
		(DR3DNS,u'focal-length'),
		(DR3DNS,u'lighting-mode'),
		(DR3DNS,u'projection'),
		(DR3DNS,u'shade-mode'),
		(DR3DNS,u'shadow-slant'),
		(DR3DNS,u'transform'),
		(DR3DNS,u'vpn'),
		(DR3DNS,u'vrp'),
		(DR3DNS,u'vup'),
		(DRAWNS,u'id'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DR3DNS,u'sphere'):(
		(DRAWNS,u'layer'),
		(DR3DNS,u'center'),
		(DR3DNS,u'transform'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(DRAWNS,u'id'),
		(DR3DNS,u'size'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'a'):(
		(OFFICENS,u'name'),
		(OFFICENS,u'server-map'),
		(OFFICENS,u'target-frame-name'),
		(OFFICENS,u'title'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'show'),
		(XLINKNS,u'type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'applet'):(
		(DRAWNS,u'code'),
		(XLINKNS,u'show'),
		(DRAWNS,u'object'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(DRAWNS,u'archive'),
		(DRAWNS,u'may-script'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'area-circle'):(
		(OFFICENS,u'name'),
		(XLINKNS,u'show'),
		(SVGNS,u'cx'),
		(XLINKNS,u'type'),
		(DRAWNS,u'nohref'),
		(SVGNS,u'cy'),
		(XLINKNS,u'href'),
		(SVGNS,u'r'),
		(OFFICENS,u'target-frame-name'),
	),
	(DRAWNS,u'area-polygon'):(
		(OFFICENS,u'name'),
		(XLINKNS,u'show'),
		(XLINKNS,u'type'),
		(SVGNS,u'height'),
		(DRAWNS,u'nohref'),
		(SVGNS,u'width'),
		(XLINKNS,u'href'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(OFFICENS,u'target-frame-name'),
		(SVGNS,u'viewBox'),
		(DRAWNS,u'points'),
	),
	(DRAWNS,u'area-rectangle'):(
		(OFFICENS,u'name'),
		(XLINKNS,u'show'),
		(XLINKNS,u'type'),
		(SVGNS,u'height'),
		(DRAWNS,u'nohref'),
		(SVGNS,u'width'),
		(XLINKNS,u'href'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(OFFICENS,u'target-frame-name'),
	),
	(DRAWNS,u'caption'):(
		(TABLENS,u'table-background'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'caption-id'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'caption-point-y'),
		(DRAWNS,u'caption-point-x'),
		(DRAWNS,u'transform'),
		(TABLENS,u'end-y'),
		(DRAWNS,u'corner-radius'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(SVGNS,u'height'),
		(DRAWNS,u'id'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'circle'):(
		(DRAWNS,u'end-angle'),
		(DRAWNS,u'id'),
		(DRAWNS,u'kind'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'name'),
		(DRAWNS,u'start-angle'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'z-index'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'cx'),
		(SVGNS,u'cy'),
		(SVGNS,u'height'),
		(SVGNS,u'r'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'connector'):(
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'end-glue-point'),
		(DRAWNS,u'end-shape'),
		(DRAWNS,u'id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'line-skew'),
		(DRAWNS,u'name'),
		(DRAWNS,u'start-glue-point'),
		(DRAWNS,u'start-shape'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'type'),
		(DRAWNS,u'z-index'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'd'),
		(SVGNS,u'viewBox'),
		(SVGNS,u'x1'),
		(SVGNS,u'x2'),
		(SVGNS,u'y1'),
		(SVGNS,u'y2'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'contour-path'):(
		(SVGNS,u'd'),
		(SVGNS,u'width'),
		(DRAWNS,u'recreate-on-edit'),
		(SVGNS,u'viewBox'),
		(SVGNS,u'height'),
	),
	(DRAWNS,u'contour-polygon'):(
		(SVGNS,u'width'),
		(DRAWNS,u'points'),
		(DRAWNS,u'recreate-on-edit'),
		(SVGNS,u'viewBox'),
		(SVGNS,u'height'),
	),
	(DRAWNS,u'control'):(
		(DRAWNS,u'control'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'caption-id'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(TABLENS,u'table-background'),
		(DRAWNS,u'transform'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'end-y'),
		(DRAWNS,u'id'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'custom-shape'):(
		(DRAWNS,u'engine'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'layer'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(TABLENS,u'table-background'),
		(DRAWNS,u'transform'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'end-y'),
		(DRAWNS,u'data'),
		(DRAWNS,u'id'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'ellipse'):(
		(DRAWNS,u'layer'),
		(DRAWNS,u'start-angle'),
		(SVGNS,u'cy'),
		(SVGNS,u'cx'),
		(TABLENS,u'table-background'),
		(TABLENS,u'end-cell-address'),
		(SVGNS,u'rx'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'id'),
		(SVGNS,u'width'),
		(TABLENS,u'end-y'),
		(TABLENS,u'end-x'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(DRAWNS,u'end-angle'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'caption-id'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'height'),
		(TEXTNS,u'anchor-type'),
		(SVGNS,u'ry'),
		(DRAWNS,u'kind'),
		(DRAWNS,u'name'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(DRAWNS,u'text-style-name'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'enhanced-geometry'):(
		(DR3DNS,u'projection'),
		(DR3DNS,u'shade-mode'),
		(DRAWNS,u'concentric-gradient-fill-allowed'),
		(DRAWNS,u'enhanced-path'),
		(DRAWNS,u'extrusion'),
		(DRAWNS,u'extrusion-allowed'),
		(DRAWNS,u'extrusion-brightness'),
		(DRAWNS,u'extrusion-color'),
		(DRAWNS,u'extrusion-depth'),
		(DRAWNS,u'extrusion-diffusion'),
		(DRAWNS,u'extrusion-first-light-direction'),
		(DRAWNS,u'extrusion-first-light-harsh'),
		(DRAWNS,u'extrusion-first-light-level'),
		(DRAWNS,u'extrusion-light-face'),
		(DRAWNS,u'extrusion-metal'),
		(DRAWNS,u'extrusion-number-of-line-segments'),
		(DRAWNS,u'extrusion-origin'),
		(DRAWNS,u'extrusion-rotation-angle'),
		(DRAWNS,u'extrusion-rotation-center'),
		(DRAWNS,u'extrusion-second-light-direction'),
		(DRAWNS,u'extrusion-second-light-harsh'),
		(DRAWNS,u'extrusion-second-light-level'),
		(DRAWNS,u'extrusion-shininess'),
		(DRAWNS,u'extrusion-skew'),
		(DRAWNS,u'extrusion-specularity'),
		(DRAWNS,u'extrusion-viewpoint'),
		(DRAWNS,u'glue-point-leaving-directions'),
		(DRAWNS,u'glue-points'),
		(DRAWNS,u'glue-point-type'),
		(DRAWNS,u'mirror-horizontal'),
		(DRAWNS,u'mirror-vertical'),
		(DRAWNS,u'modifiers'),
		(DRAWNS,u'path-stretchpoint-x'),
		(DRAWNS,u'path-stretchpoint-y'),
		(DRAWNS,u'text-areas'),
		(DRAWNS,u'text-path'),
		(DRAWNS,u'text-path-allowed'),
		(DRAWNS,u'text-path-mode'),
		(DRAWNS,u'text-path-same-letter-heights'),
		(DRAWNS,u'text-path-scale'),
		(DRAWNS,u'text-rotate-angle'),
		(DRAWNS,u'type'),
		(SVGNS,u'viewBox'),
	),
# allowed_attributes
	(DRAWNS,u'equation'):(
		(DRAWNS,u'formula'),
		(DRAWNS,u'name'),
	),
	(DRAWNS,u'fill-image'):(
		(DRAWNS,u'name'),
		(XLINKNS,u'show'),
		(XLINKNS,u'actuate'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(XLINKNS,u'href'),
		(DRAWNS,u'display-name'),
		(XLINKNS,u'type'),
	),
	(DRAWNS,u'floating-frame'):(
		(XLINKNS,u'href'),
		(XLINKNS,u'actuate'),
		(DRAWNS,u'frame-name'),
		(XLINKNS,u'type'),
		(XLINKNS,u'show'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'frame'):(
		(DRAWNS,u'copy-of'),
		(DRAWNS,u'id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'name'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'z-index'),
		(PRESENTATIONNS,u'class'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'placeholder'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'user-transformed'),
		(STYLENS,u'rel-height'),
		(STYLENS,u'rel-width'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'g'):(
		(DRAWNS,u'id'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'name'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'z-index'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'glue-point'):(
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(DRAWNS,u'align'),
		(DRAWNS,u'id'),
		(DRAWNS,u'escape-direction'),
	),
	(DRAWNS,u'gradient'):(
		(DRAWNS,u'style'),
		(DRAWNS,u'angle'),
		(DRAWNS,u'name'),
		(DRAWNS,u'end-color'),
		(DRAWNS,u'start-color'),
		(DRAWNS,u'cy'),
		(DRAWNS,u'cx'),
		(DRAWNS,u'display-name'),
		(DRAWNS,u'border'),
		(DRAWNS,u'end-intensity'),
		(DRAWNS,u'start-intensity'),
	),
	(DRAWNS,u'handle'):(
		(DRAWNS,u'handle-radius-range-minimum'),
		(DRAWNS,u'handle-switched'),
		(DRAWNS,u'handle-range-y-maximum'),
		(DRAWNS,u'handle-mirror-horizontal'),
		(DRAWNS,u'handle-range-x-maximum'),
		(DRAWNS,u'handle-mirror-vertical'),
		(DRAWNS,u'handle-range-y-minimum'),
		(DRAWNS,u'handle-radius-range-maximum'),
		(DRAWNS,u'handle-range-x-minimum'),
		(DRAWNS,u'handle-position'),
		(DRAWNS,u'handle-polar'),
	),
	(DRAWNS,u'hatch'):(
		(DRAWNS,u'distance'),
		(DRAWNS,u'style'),
		(DRAWNS,u'name'),
		(DRAWNS,u'color'),
		(DRAWNS,u'display-name'),
		(DRAWNS,u'rotation'),
	),
	(DRAWNS,u'image'):(
		(DRAWNS,u'filter-name'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'show'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'image-map'):(
	),
	(DRAWNS,u'layer'):(
		(DRAWNS,u'protected'),
		(DRAWNS,u'name'),
		(DRAWNS,u'display'),
	),
# allowed_attributes
	(DRAWNS,u'layer-set'):(
	),
	(DRAWNS,u'line'):(
		(DRAWNS,u'class-names'),
		(DRAWNS,u'id'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'name'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'z-index'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'x1'),
		(SVGNS,u'x2'),
		(SVGNS,u'y1'),
		(SVGNS,u'y2'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'marker'):(
		(SVGNS,u'd'),
		(DRAWNS,u'display-name'),
		(DRAWNS,u'name'),
		(SVGNS,u'viewBox'),
	),
# allowed_attributes
	(DRAWNS,u'measure'):(
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'layer'),
		(SVGNS,u'y2'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(TABLENS,u'table-background'),
		(SVGNS,u'x2'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y1'),
		(DRAWNS,u'caption-id'),
		(TABLENS,u'end-y'),
		(SVGNS,u'x1'),
		(DRAWNS,u'id'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'object'):(
		(XLINKNS,u'type'),
		(XLINKNS,u'href'),
		(DRAWNS,u'notify-on-update-of-ranges'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'show'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'object-ole'):(
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(DRAWNS,u'class-id'),
		(XLINKNS,u'show'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'opacity'):(
		(DRAWNS,u'style'),
		(DRAWNS,u'angle'),
		(DRAWNS,u'name'),
		(DRAWNS,u'start'),
		(DRAWNS,u'cy'),
		(DRAWNS,u'cx'),
		(DRAWNS,u'end'),
		(DRAWNS,u'display-name'),
		(DRAWNS,u'border'),
	),
	(DRAWNS,u'page'):(
		(PRESENTATIONNS,u'presentation-page-layout-name'),
		(DRAWNS,u'name'),
		(DRAWNS,u'nav-order'),
		(PRESENTATIONNS,u'use-footer-name'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'use-header-name'),
		(DRAWNS,u'master-page-name'),
		(DRAWNS,u'id'),
		(PRESENTATIONNS,u'use-date-time-name'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'page-thumbnail'):(
		(TABLENS,u'table-background'),
		(DRAWNS,u'caption-id'),
		(PRESENTATIONNS,u'user-transformed'),
		(DRAWNS,u'layer'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'name'),
		(DRAWNS,u'id'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'page-number'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'end-y'),
		(PRESENTATIONNS,u'placeholder'),
		(PRESENTATIONNS,u'class'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'param'):(
		(DRAWNS,u'name'),
		(DRAWNS,u'value'),
	),
# allowed_attributes
	(DRAWNS,u'path'):(
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'name'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'z-index'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'd'),
		(SVGNS,u'height'),
		(SVGNS,u'viewBox'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'plugin'):(
		(DRAWNS,u'mime-type'),
		(XLINKNS,u'type'),
		(XLINKNS,u'href'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'show'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'polygon'):(
		(DRAWNS,u'caption-id'),
		(TABLENS,u'table-background'),
		(DRAWNS,u'layer'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'id'),
		(DRAWNS,u'transform'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'points'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'end-y'),
		(SVGNS,u'viewBox'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'polyline'):(
		(TABLENS,u'table-background'),
		(DRAWNS,u'layer'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'id'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'transform'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'points'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(TEXTNS,u'anchor-page-number'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'end-y'),
		(SVGNS,u'viewBox'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'rect'):(
		(DRAWNS,u'corner-radius'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(SVGNS,u'rx'),
		(SVGNS,u'ry'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(DRAWNS,u'regular-polygon'):(
		(TABLENS,u'table-background'),
		(DRAWNS,u'layer'),
		(TABLENS,u'end-cell-address'),
		(DRAWNS,u'caption-id'),
		(DRAWNS,u'name'),
		(DRAWNS,u'text-style-name'),
		(TEXTNS,u'anchor-page-number'),
		(DRAWNS,u'concave'),
		(DRAWNS,u'sharpness'),
		(DRAWNS,u'transform'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(DRAWNS,u'z-index'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'style-name'),
		(PRESENTATIONNS,u'class-names'),
		(TABLENS,u'end-x'),
		(DRAWNS,u'corners'),
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(TABLENS,u'end-y'),
		(DRAWNS,u'id'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(DRAWNS,u'stroke-dash'):(
		(DRAWNS,u'distance'),
		(DRAWNS,u'dots1-length'),
		(DRAWNS,u'name'),
		(DRAWNS,u'dots2-length'),
		(DRAWNS,u'style'),
		(DRAWNS,u'dots1'),
		(DRAWNS,u'display-name'),
		(DRAWNS,u'dots2'),
	),
	(DRAWNS,u'text-box'):(
		(FONS,u'min-width'),
		(DRAWNS,u'corner-radius'),
		(FONS,u'max-height'),
		(FONS,u'min-height'),
		(DRAWNS,u'chain-next-name'),
		(FONS,u'max-width'),
		(TEXTNS,u'id'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(FORMNS,u'button'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'focus-on-click'),
		(FORMNS,u'image-align'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'button-type'),
		(FORMNS,u'title'),
		(FORMNS,u'default-button'),
		(FORMNS,u'value'),
		(FORMNS,u'label'),
		(FORMNS,u'delay-for-repeat'),
		(FORMNS,u'repeat'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'image-data'),
		(XLINKNS,u'href'),
		(FORMNS,u'toggle'),
		(FORMNS,u'xforms-submission'),
		(OFFICENS,u'target-frame'),
		(FORMNS,u'image-position'),
	),
	(FORMNS,u'checkbox'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'image-align'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'data-field'),
		(FORMNS,u'title'),
		(FORMNS,u'is-tristate'),
		(FORMNS,u'current-state'),
		(FORMNS,u'value'),
		(FORMNS,u'label'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'state'),
		(FORMNS,u'visual-effect'),
		(FORMNS,u'image-position'),
	),
	(FORMNS,u'column'):(
		(FORMNS,u'control-implementation'),
		(FORMNS,u'text-style-name'),
		(FORMNS,u'name'),
		(FORMNS,u'label'),
	),
# allowed_attributes
	(FORMNS,u'combobox'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'auto-complete'),
		(FORMNS,u'control-implementation'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'current-value'),
		(FORMNS,u'data-field'),
		(FORMNS,u'disabled'),
		(FORMNS,u'dropdown'),
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'list-source'),
		(FORMNS,u'list-source-type'),
		(FORMNS,u'max-length'),
		(FORMNS,u'name'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'size'),
		(FORMNS,u'source-cell-range'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(XFORMSNS,u'bind'),
	),
# allowed_attributes
	(FORMNS,u'connection-resource'):(
		(XLINKNS,u'href'),
	),
	(FORMNS,u'date'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'control-implementation'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'current-value'),
		(FORMNS,u'data-field'),
		(FORMNS,u'delay-for-repeat'),
		(FORMNS,u'disabled'),
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'max-length'),
		(FORMNS,u'max-value'),
		(FORMNS,u'min-value'),
		(FORMNS,u'name'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'repeat'),
		(FORMNS,u'spin-button'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(XFORMSNS,u'bind'),
	),
	(FORMNS,u'file'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'max-length'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'current-value'),
	),
	(FORMNS,u'fixed-text'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'name'),
		(FORMNS,u'for'),
		(FORMNS,u'title'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'multi-line'),
		(FORMNS,u'label'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
	),
# allowed_attributes
	(FORMNS,u'form'):(
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(FORMNS,u'allow-deletes'),
		(FORMNS,u'command-type'),
		(FORMNS,u'apply-filter'),
		(XLINKNS,u'type'),
		(FORMNS,u'method'),
		(OFFICENS,u'target-frame'),
		(FORMNS,u'navigation-mode'),
		(FORMNS,u'detail-fields'),
		(FORMNS,u'master-fields'),
		(FORMNS,u'allow-updates'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-cycle'),
		(FORMNS,u'control-implementation'),
		(FORMNS,u'escape-processing'),
		(FORMNS,u'filter'),
		(FORMNS,u'command'),
		(FORMNS,u'datasource'),
		(FORMNS,u'enctype'),
		(FORMNS,u'allow-inserts'),
		(FORMNS,u'ignore-result'),
		(FORMNS,u'order'),
	),
# allowed_attributes
	(FORMNS,u'formatted-text'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'control-implementation'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'current-value'),
		(FORMNS,u'data-field'),
		(FORMNS,u'delay-for-repeat'),
		(FORMNS,u'disabled'),
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'max-length'),
		(FORMNS,u'max-value'),
		(FORMNS,u'min-value'),
		(FORMNS,u'name'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'repeat'),
		(FORMNS,u'spin-button'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'title'),
		(FORMNS,u'validation'),
		(FORMNS,u'value'),
		(XFORMSNS,u'bind'),
	),
	(FORMNS,u'frame'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'name'),
		(FORMNS,u'for'),
		(FORMNS,u'title'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'label'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
	),
# allowed_attributes
	(FORMNS,u'generic-control'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'name'),
	),
	(FORMNS,u'grid'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'title'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
	),
	(FORMNS,u'hidden'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'name'),
		(FORMNS,u'value'),
	),
	(FORMNS,u'image'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'button-type'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(OFFICENS,u'target-frame'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'image-data'),
		(XLINKNS,u'href'),
	),
	(FORMNS,u'image-frame'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'name'),
		(FORMNS,u'title'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'data-field'),
		(FORMNS,u'readonly'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'image-data'),
	),
	(FORMNS,u'item'):(
		(FORMNS,u'label'),
	),
	(FORMNS,u'list-property'):(
		(FORMNS,u'property-name'),
		(OFFICENS,u'value-type'),
	),
	(FORMNS,u'list-value'):(
		(OFFICENS,u'string-value'),
	),
# allowed_attributes
	(FORMNS,u'listbox'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'bound-column'),
		(FORMNS,u'control-implementation'),
		(FORMNS,u'data-field'),
		(FORMNS,u'disabled'),
		(FORMNS,u'dropdown'),
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'list-linkage-type'),
		(FORMNS,u'list-source'),
		(FORMNS,u'list-source-type'),
		(FORMNS,u'multiple'),
		(FORMNS,u'name'),
		(FORMNS,u'printable'),
		(FORMNS,u'size'),
		(FORMNS,u'source-cell-range'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'title'),
		(FORMNS,u'xforms-list-source'),
		(XFORMSNS,u'bind'),
	),
	(FORMNS,u'number'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'control-implementation'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'current-value'),
		(FORMNS,u'data-field'),
		(FORMNS,u'delay-for-repeat'),
		(FORMNS,u'disabled'),
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'max-length'),
		(FORMNS,u'max-value'),
		(FORMNS,u'min-value'),
		(FORMNS,u'name'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'repeat'),
		(FORMNS,u'spin-button'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(XFORMSNS,u'bind'),
	),
	(FORMNS,u'option'):(
		(FORMNS,u'current-selected'),
		(FORMNS,u'selected'),
		(FORMNS,u'value'),
		(FORMNS,u'label'),
	),
	(FORMNS,u'password'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'max-length'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'echo-char'),
	),
	(FORMNS,u'properties'):(
	),
	(FORMNS,u'property'):(
		(OFFICENS,u'string-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'boolean-value'),
		(FORMNS,u'property-name'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(OFFICENS,u'value-type'),
		(OFFICENS,u'time-value'),
	),
	(FORMNS,u'radio'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'selected'),
		(FORMNS,u'image-align'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'data-field'),
		(FORMNS,u'current-selected'),
		(FORMNS,u'value'),
		(FORMNS,u'label'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'title'),
		(FORMNS,u'visual-effect'),
		(FORMNS,u'image-position'),
	),
	(FORMNS,u'text'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'max-length'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'data-field'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'current-value'),
	),
	(FORMNS,u'textarea'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'max-length'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'data-field'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'current-value'),
	),
	(FORMNS,u'time'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'control-implementation'),
		(FORMNS,u'convert-empty-to-null'),
		(FORMNS,u'current-value'),
		(FORMNS,u'data-field'),
		(FORMNS,u'delay-for-repeat'),
		(FORMNS,u'disabled'),
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'max-length'),
		(FORMNS,u'max-value'),
		(FORMNS,u'min-value'),
		(FORMNS,u'name'),
		(FORMNS,u'printable'),
		(FORMNS,u'readonly'),
		(FORMNS,u'repeat'),
		(FORMNS,u'spin-button'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(XFORMSNS,u'bind'),
	),
	(FORMNS,u'value-range'):(
		(XMLNS,u'id'), # First choice
		(FORMNS,u'id'),
		(FORMNS,u'linked-cell'),
		(FORMNS,u'tab-stop'),
		(FORMNS,u'max-value'),
		(FORMNS,u'name'),
		(FORMNS,u'tab-index'),
		(FORMNS,u'control-implementation'),
		(XFORMSNS,u'bind'),
		(FORMNS,u'title'),
		(FORMNS,u'value'),
		(FORMNS,u'disabled'),
		(FORMNS,u'printable'),
		(FORMNS,u'orientation'),
		(FORMNS,u'page-step-size'),
		(FORMNS,u'delay-for-repeat'),
		(FORMNS,u'repeat'),
		(FORMNS,u'min-value'),
		(FORMNS,u'step-size'),
	),
	(MANIFESTNS,'algorithm') : (
		(MANIFESTNS,'algorithm-name'),
		(MANIFESTNS,'initialisation-vector'),
	),
	(MANIFESTNS,'encryption-data') : (
		(MANIFESTNS,'checksum-type'),
		(MANIFESTNS,'checksum'),
	),
	(MANIFESTNS,'file-entry') : (
		(MANIFESTNS,'full-path'),
		(MANIFESTNS,'media-type'),
		(MANIFESTNS,'preferred-view-mode'),
		(MANIFESTNS,'size'),
		(MANIFESTNS,'version'),
	),
	(MANIFESTNS,'key-derivation') : (
		(MANIFESTNS,'key-derivation-name'),
		(MANIFESTNS,'salt'),
		(MANIFESTNS,'iteration-count'),
	),
	(MANIFESTNS,u'manifest'):(
	),
# allowed_attributes
	(METANS,u'auto-reload'):(
		(METANS,u'delay'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(XLINKNS,u'show'),
	),
	(METANS,u'creation-date'):(
	),
	(METANS,u'date-string'):(
	),
	(METANS,u'document-statistic'):(
		(METANS,u'non-whitespace-character-count'),
		(METANS,u'ole-object-count'),
		(METANS,u'table-count'),
		(METANS,u'row-count'),
		(METANS,u'character-count'),
		(METANS,u'sentence-count'),
		(METANS,u'draw-count'),
		(METANS,u'paragraph-count'),
		(METANS,u'word-count'),
		(METANS,u'object-count'),
		(METANS,u'syllable-count'),
		(METANS,u'image-count'),
		(METANS,u'page-count'),
		(METANS,u'frame-count'),
		(METANS,u'cell-count'),
	),
	(METANS,u'editing-cycles'):(
	),
	(METANS,u'editing-duration'):(
	),
	(METANS,u'generator'):(
	),
# allowed_attributes
	(METANS,u'hyperlink-behaviour'):(
		(OFFICENS,u'target-frame-name'),
		(XLINKNS,u'show'),
	),
	(METANS,u'initial-creator'):(
	),
	(METANS,u'keyword'):(
	),
	(METANS,u'print-date'):(
	),
	(METANS,u'printed-by'):(
	),
	(METANS,u'template'):(
		(METANS,u'date'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(XLINKNS,u'title'),
	),
	(METANS,u'user-defined'):(
		(METANS,u'name'),
		(METANS,u'value-type'),
	),
# allowed_attributes
	(OFFICENS,u'annotation'):(
		(DRAWNS,u'caption-point-x'),
		(DRAWNS,u'caption-point-y'),
		(DRAWNS,u'class-names'),
		(DRAWNS,u'corner-radius'),
		(DRAWNS,u'id'),
		(DRAWNS,u'layer'),
		(DRAWNS,u'name'),
		(DRAWNS,u'style-name'),
		(DRAWNS,u'text-style-name'),
		(DRAWNS,u'transform'),
		(DRAWNS,u'z-index'),
		(OFFICENS,u'display'),
		(OFFICENS,u'name'),
		(PRESENTATIONNS,u'class-names'),
		(PRESENTATIONNS,u'style-name'),
		(SVGNS,u'height'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TABLENS,u'end-cell-address'),
		(TABLENS,u'end-x'),
		(TABLENS,u'end-y'),
		(TABLENS,u'table-background'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(XMLNS,u'id'),
	),
	(OFFICENS,u'annotation-end'): (
		(OFFICENS,u'name'),
	),
	(OFFICENS,u'automatic-styles'):(
	),
	(OFFICENS,u'binary-data'):(
	),
	(OFFICENS,u'body'):(
	),
	(OFFICENS,u'change-info'):(
	),
	(OFFICENS,u'chart'):(
	),
# allowed_attributes
	(OFFICENS,u'dde-source'):(
		(OFFICENS,u'dde-application'),
		(OFFICENS,u'automatic-update'),
		(OFFICENS,u'conversion-mode'),
		(OFFICENS,u'dde-item'),
		(OFFICENS,u'dde-topic'),
		(OFFICENS,u'name'),
	),
	(OFFICENS,u'document'):(
		(OFFICENS,u'mimetype'),
		(OFFICENS,u'version'),
		(GRDDLNS,u'transformation'),
	),
	(OFFICENS,u'document-content'):(
		(OFFICENS,u'version'),
		(GRDDLNS,u'transformation'),
	),
	(OFFICENS,u'document-meta'):(
		(OFFICENS,u'version'),
		(GRDDLNS,u'transformation'),
	),
	(OFFICENS,u'document-settings'):(
		(OFFICENS,u'version'),
		(GRDDLNS,u'transformation'),
	),
	(OFFICENS,u'document-styles'):(
		(OFFICENS,u'version'),
		(GRDDLNS,u'transformation'),
	),
	(OFFICENS,u'drawing'):(
	),
	(OFFICENS,u'event-listeners'):(
	),
	(OFFICENS,u'font-face-decls'):(
	),
	(OFFICENS,u'forms'):(
		(FORMNS,u'automatic-focus'),
		(FORMNS,u'apply-design-mode'),
	),
	(OFFICENS,u'image'):(
	),
# allowed_attributes
	(OFFICENS,u'master-styles'):(
	),
	(OFFICENS,u'meta'):(
	),
	(OFFICENS,u'presentation'):(
	),
	(OFFICENS,u'script'):(
		(SCRIPTNS,u'language'),
	),
	(OFFICENS,u'scripts'):(
	),
	(OFFICENS,u'settings'):(
	),
	(OFFICENS,u'spreadsheet'):(
		(TABLENS,u'structure-protected'),
		(TABLENS,u'protection-key'),
		(TABLENS,u'protection-key-digest-algorithm'),
	),
	(OFFICENS,u'styles'):(
	),
	(OFFICENS,u'text'):(
		(TEXTNS,u'global'),
		(TEXTNS,u'use-soft-page-breaks'),
	),
	(PRESENTATIONNS,u'animation-group'):(
	),
	(PRESENTATIONNS,u'animations'):(
	),
	(PRESENTATIONNS,u'date-time'):(
	),
	(PRESENTATIONNS,u'date-time-decl'):(
		(PRESENTATIONNS,u'source'),
		(STYLENS,u'data-style-name'),
		(PRESENTATIONNS,u'name'),
	),
	(PRESENTATIONNS,u'dim'):(
		(DRAWNS,u'color'),
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'event-listener'):(
		(PRESENTATIONNS,u'direction'),
		(XLINKNS,u'show'),
		(XLINKNS,u'type'),
		(XLINKNS,u'actuate'),
		(PRESENTATIONNS,u'effect'),
		(SCRIPTNS,u'event-name'),
		(PRESENTATIONNS,u'start-scale'),
		(XLINKNS,u'href'),
		(PRESENTATIONNS,u'verb'),
		(PRESENTATIONNS,u'action'),
		(PRESENTATIONNS,u'speed'),
	),
	(PRESENTATIONNS,u'footer'):(
	),
	(PRESENTATIONNS,u'footer-decl'):(
		(PRESENTATIONNS,u'name'),
	),
	(PRESENTATIONNS,u'header'):(
	),
	(PRESENTATIONNS,u'header-decl'):(
		(PRESENTATIONNS,u'name'),
	),
	(PRESENTATIONNS,u'hide-shape'):(
		(PRESENTATIONNS,u'direction'),
		(PRESENTATIONNS,u'effect'),
		(PRESENTATIONNS,u'delay'),
		(PRESENTATIONNS,u'start-scale'),
		(PRESENTATIONNS,u'path-id'),
		(PRESENTATIONNS,u'speed'),
		(DRAWNS,u'shape-id'),
	),
# allowed_attributes
	(PRESENTATIONNS,u'hide-text'):(
		(PRESENTATIONNS,u'direction'),
		(PRESENTATIONNS,u'effect'),
		(PRESENTATIONNS,u'delay'),
		(PRESENTATIONNS,u'start-scale'),
		(PRESENTATIONNS,u'path-id'),
		(PRESENTATIONNS,u'speed'),
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'notes'):(
		(STYLENS,u'page-layout-name'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'use-header-name'),
		(PRESENTATIONNS,u'use-date-time-name'),
		(PRESENTATIONNS,u'use-footer-name'),
	),
	(PRESENTATIONNS,u'placeholder'):(
		(SVGNS,u'y'),
		(SVGNS,u'x'),
		(SVGNS,u'height'),
		(PRESENTATIONNS,u'object'),
		(SVGNS,u'width'),
	),
	(PRESENTATIONNS,u'play'):(
		(PRESENTATIONNS,u'speed'),
		(DRAWNS,u'shape-id'),
	),
# allowed_attributes
	(PRESENTATIONNS,u'settings'):(
		(PRESENTATIONNS,u'animations'),
		(PRESENTATIONNS,u'endless'),
		(PRESENTATIONNS,u'force-manual'),
		(PRESENTATIONNS,u'full-screen'),
		(PRESENTATIONNS,u'mouse-as-pen'),
		(PRESENTATIONNS,u'mouse-visible'),
		(PRESENTATIONNS,u'pause'),
		(PRESENTATIONNS,u'show'),
		(PRESENTATIONNS,u'show-end-of-presentation-slide'),
		(PRESENTATIONNS,u'show-logo'),
		(PRESENTATIONNS,u'start-page'),
		(PRESENTATIONNS,u'start-with-navigator'),
		(PRESENTATIONNS,u'stay-on-top'),
		(PRESENTATIONNS,u'transition-on-click'),
	),
	(PRESENTATIONNS,u'show'):(
		(PRESENTATIONNS,u'name'),
		(PRESENTATIONNS,u'pages'),
	),
	(PRESENTATIONNS,u'show-shape'):(
		(PRESENTATIONNS,u'direction'),
		(PRESENTATIONNS,u'effect'),
		(PRESENTATIONNS,u'delay'),
		(PRESENTATIONNS,u'start-scale'),
		(PRESENTATIONNS,u'path-id'),
		(PRESENTATIONNS,u'speed'),
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'show-text'):(
		(PRESENTATIONNS,u'direction'),
		(PRESENTATIONNS,u'effect'),
		(PRESENTATIONNS,u'delay'),
		(PRESENTATIONNS,u'start-scale'),
		(PRESENTATIONNS,u'path-id'),
		(PRESENTATIONNS,u'speed'),
		(DRAWNS,u'shape-id'),
	),
	(PRESENTATIONNS,u'sound'):(
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(PRESENTATIONNS,u'play-full'),
		(XLINKNS,u'show'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(SCRIPTNS,u'event-listener'):(
		(SCRIPTNS,u'language'),
		(SCRIPTNS,u'macro-name'),
		(XLINKNS,u'actuate'),
		(SCRIPTNS,u'event-name'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(STYLENS,u'background-image'):(
		(DRAWNS,u'opacity'),
		(STYLENS,u'repeat'),
		(XLINKNS,u'show'),
		(XLINKNS,u'actuate'),
		(STYLENS,u'filter-name'),
		(XLINKNS,u'href'),
		(STYLENS,u'position'),
		(XLINKNS,u'type'),
	),
# allowed_attributes
	(STYLENS,u'chart-properties'): (
		(CHARTNS,u'angle-offset'),
		(CHARTNS,u'auto-position'),
		(CHARTNS,u'auto-size'),
		(CHARTNS,u'axis-label-position'),
		(CHARTNS,u'axis-position'),
		(CHARTNS,u'connect-bars'),
		(CHARTNS,u'data-label-number'),
		(CHARTNS,u'data-label-symbol'),
		(CHARTNS,u'data-label-text'),
		(CHARTNS,u'deep'),
		(CHARTNS,u'display-label'),
		(CHARTNS,u'error-category'),
		(CHARTNS,u'error-lower-indicator'),
		(CHARTNS,u'error-lower-limit'),
		(CHARTNS,u'error-lower-range'),
		(CHARTNS,u'error-margin'),
		(CHARTNS,u'error-percentage'),
		(CHARTNS,u'error-upper-indicator'),
		(CHARTNS,u'error-upper-limit'),
		(CHARTNS,u'error-upper-range'),
		(CHARTNS,u'gap-width'),
		(CHARTNS,u'group-bars-per-axis'),
		(CHARTNS,u'hole-size'),
		(CHARTNS,u'include-hidden-cells'),
		(CHARTNS,u'interpolation'),
		(CHARTNS,u'interval-major'),
		(CHARTNS,u'interval-minor-divisor'),
		(CHARTNS,u'japanese-candle-stick'),
		(CHARTNS,u'label-arrangement'),
		(CHARTNS,u'label-position'),
		(CHARTNS,u'label-position-negative'),
		(CHARTNS,u'lines'),
		(CHARTNS,u'link-data-style-to-source'),
		(CHARTNS,u'logarithmic'),
		(CHARTNS,u'maximum'),
		(CHARTNS,u'mean-value'),
		(CHARTNS,u'minimum'),
		(CHARTNS,u'origin'),
		(CHARTNS,u'overlap'),
		(CHARTNS,u'percentage'),
		(CHARTNS,u'pie-offset'),
		(CHARTNS,u'regression-type'),
		(CHARTNS,u'reverse-direction'),
		(CHARTNS,u'right-angled-axes'),
		(CHARTNS,u'scale-text'),
		(CHARTNS,u'series-source'),
		(CHARTNS,u'solid-type'),
		(CHARTNS,u'sort-by-x-values'),
		(CHARTNS,u'spline-order'),
		(CHARTNS,u'spline-resolution'),
		(CHARTNS,u'stacked'),
		(CHARTNS,u'symbol-height'),
		(CHARTNS,u'symbol-name'),
		(CHARTNS,u'symbol-type'),
		(CHARTNS,u'symbol-width'),
		(CHARTNS,u'text-overlap'),
		(CHARTNS,u'three-dimensional'),
		(CHARTNS,u'tick-mark-position'),
		(CHARTNS,u'tick-marks-major-inner'),
		(CHARTNS,u'tick-marks-major-outer'),
		(CHARTNS,u'tick-marks-minor-inner'),
		(CHARTNS,u'tick-marks-minor-outer'),
		(CHARTNS,u'treat-empty-cells'),
		(CHARTNS,u'vertical'),
		(CHARTNS,u'visible'),
		(STYLENS,u'direction'),
		(STYLENS,u'rotation-angle'),
		(TEXTNS,u'line-break'),
	),
	(STYLENS,u'column'):(
		(FONS,u'end-indent'),
		(FONS,u'space-before'),
		(FONS,u'start-indent'),
		(FONS,u'space-after'),
		(STYLENS,u'rel-width'),
	),
	(STYLENS,u'column-sep'):(
		(STYLENS,u'color'),
		(STYLENS,u'width'),
		(STYLENS,u'style'),
		(STYLENS,u'vertical-align'),
		(STYLENS,u'height'),
	),
	(STYLENS,u'columns'):(
		(FONS,u'column-count'),
		(FONS,u'column-gap'),
	),
# allowed_attributes
	(STYLENS,u'default-page-layout'):(
	),
	(STYLENS,u'default-style'):(
		(STYLENS,u'family'),
	),
# allowed_attributes
	(STYLENS,u'drawing-page-properties'): (
		(DRAWNS,u'fill'),
		(DRAWNS,u'fill-color'),
		(DRAWNS,u'secondary-fill-color'),
		(DRAWNS,u'fill-gradient-name'),
		(DRAWNS,u'gradient-step-count'),
		(DRAWNS,u'fill-hatch-name'),
		(DRAWNS,u'fill-hatch-solid'),
		(DRAWNS,u'fill-image-name'),
		(STYLENS,u'repeat'),
		(DRAWNS,u'fill-image-width'),
		(DRAWNS,u'fill-image-height'),
		(DRAWNS,u'fill-image-ref-point-x'),
		(DRAWNS,u'fill-image-ref-point-y'),
		(DRAWNS,u'fill-image-ref-point'),
		(DRAWNS,u'tile-repeat-offset'),
		(DRAWNS,u'opacity'),
		(DRAWNS,u'opacity-name'),
		(SVGNS,u'fill-rule'),
		(PRESENTATIONNS,u'transition-type'),
		(PRESENTATIONNS,u'transition-style'),
		(PRESENTATIONNS,u'transition-speed'),
		(SMILNS,u'type'),
		(SMILNS,u'subtype'),
		(SMILNS,u'direction'),
		(SMILNS,u'fadeColor'),
		(PRESENTATIONNS,u'duration'),
		(PRESENTATIONNS,u'visibility'),
		(DRAWNS,u'background-size'),
		(PRESENTATIONNS,u'background-objects-visible'),
		(PRESENTATIONNS,u'background-visible'),
		(PRESENTATIONNS,u'display-header'),
		(PRESENTATIONNS,u'display-footer'),
		(PRESENTATIONNS,u'display-page-number'),
		(PRESENTATIONNS,u'display-date-time'),
	),
	(STYLENS,u'drop-cap'):(
		(STYLENS,u'distance'),
		(STYLENS,u'length'),
		(STYLENS,u'style-name'),
		(STYLENS,u'lines'),
	),
# allowed_attributes
	(STYLENS,u'font-face'):(
		(STYLENS,u'font-adornments'),
		(STYLENS,u'font-charset'),
		(STYLENS,u'font-family-generic'),
		(STYLENS,u'font-pitch'),
		(STYLENS,u'name'),
		(SVGNS,u'accent-height'),
		(SVGNS,u'alphabetic'),
		(SVGNS,u'ascent'),
		(SVGNS,u'bbox'),
		(SVGNS,u'cap-height'),
		(SVGNS,u'descent'),
		(SVGNS,u'font-family'),
		(SVGNS,u'font-size'),
		(SVGNS,u'font-stretch'),
		(SVGNS,u'font-style'),
		(SVGNS,u'font-variant'),
		(SVGNS,u'font-weight'),
		(SVGNS,u'hanging'),
		(SVGNS,u'ideographic'),
		(SVGNS,u'mathematical'),
		(SVGNS,u'overline-position'),
		(SVGNS,u'overline-thickness'),
		(SVGNS,u'panose-1'),
		(SVGNS,u'slope'),
		(SVGNS,u'stemh'),
		(SVGNS,u'stemv'),
		(SVGNS,u'strikethrough-position'),
		(SVGNS,u'strikethrough-thickness'),
		(SVGNS,u'underline-position'),
		(SVGNS,u'underline-thickness'),
		(SVGNS,u'unicode-range'),
		(SVGNS,u'units-per-em'),
		(SVGNS,u'v-alphabetic'),
		(SVGNS,u'v-hanging'),
		(SVGNS,u'v-ideographic'),
		(SVGNS,u'v-mathematical'),
		(SVGNS,u'widths'),
		(SVGNS,u'x-height'),
	),
	(STYLENS,u'footer'):(
		(STYLENS,u'display'),
	),
	(STYLENS,u'footer-left'):(
		(STYLENS,u'display'),
	),
	(STYLENS,u'footer-style'):(
	),
	(STYLENS,u'footnote-sep'):(
		(STYLENS,u'distance-after-sep'),
		(STYLENS,u'color'),
		(STYLENS,u'rel-width'),
		(STYLENS,u'width'),
		(STYLENS,u'distance-before-sep'),
		(STYLENS,u'line-style'),
		(STYLENS,u'adjustment'),
	),
# allowed_attributes
	(STYLENS,u'graphic-properties'): (
		(DR3DNS,u'ambient-color'),
		(DR3DNS,u'back-scale'),
		(DR3DNS,u'backface-culling'),
		(DR3DNS,u'close-back'),
		(DR3DNS,u'close-front'),
		(DR3DNS,u'depth'),
		(DR3DNS,u'diffuse-color'),
		(DR3DNS,u'edge-rounding'),
		(DR3DNS,u'edge-rounding-mode'),
		(DR3DNS,u'emissive-color'),
		(DR3DNS,u'end-angle'),
		(DR3DNS,u'horizontal-segments'),
		(DR3DNS,u'lighting-mode'),
		(DR3DNS,u'normals-direction'),
		(DR3DNS,u'normals-kind'),
		(DR3DNS,u'shadow'),
		(DR3DNS,u'shininess'),
		(DR3DNS,u'specular-color'),
		(DR3DNS,u'texture-filter'),
		(DR3DNS,u'texture-generation-mode-x'),
		(DR3DNS,u'texture-generation-mode-y'),
		(DR3DNS,u'texture-kind'),
		(DR3DNS,u'texture-mode'),
		(DR3DNS,u'vertical-segments'),
		(DRAWNS,u'auto-grow-height'),
		(DRAWNS,u'auto-grow-width'),
		(DRAWNS,u'blue'),
		(DRAWNS,u'caption-angle'),
		(DRAWNS,u'caption-angle-type'),
		(DRAWNS,u'caption-escape'),
		(DRAWNS,u'caption-escape-direction'),
		(DRAWNS,u'caption-fit-line-length'),
		(DRAWNS,u'caption-gap'),
		(DRAWNS,u'caption-line-length'),
		(DRAWNS,u'caption-type'),
		(DRAWNS,u'color-inversion'),
		(DRAWNS,u'color-mode'),
		(DRAWNS,u'contrast'),
		(DRAWNS,u'decimal-places'),
		(DRAWNS,u'draw-aspect'),
		(DRAWNS,u'end-guide'),
		(DRAWNS,u'end-line-spacing-horizontal'),
		(DRAWNS,u'end-line-spacing-vertical'),
		(DRAWNS,u'fill'),
		(DRAWNS,u'fill-color'),
		(DRAWNS,u'fill-gradient-name'),
		(DRAWNS,u'fill-hatch-name'),
		(DRAWNS,u'fill-hatch-solid'),
		(DRAWNS,u'fill-image-height'),
		(DRAWNS,u'fill-image-name'),
		(DRAWNS,u'fill-image-ref-point'),
		(DRAWNS,u'fill-image-ref-point-x'),
		(DRAWNS,u'fill-image-ref-point-y'),
		(DRAWNS,u'fill-image-width'),
# allowed_attributes
		(DRAWNS,u'fit-to-contour'),
		(DRAWNS,u'fit-to-size'),
		(DRAWNS,u'frame-display-border'),
		(DRAWNS,u'frame-display-scrollbar'),
		(DRAWNS,u'frame-margin-horizontal'),
		(DRAWNS,u'frame-margin-vertical'),
		(DRAWNS,u'gamma'),
		(DRAWNS,u'gradient-step-count'),
		(DRAWNS,u'green'),
		(DRAWNS,u'guide-distance'),
		(DRAWNS,u'guide-overhang'),
		(DRAWNS,u'image-opacity'),
		(DRAWNS,u'line-distance'),
		(DRAWNS,u'luminance'),
		(DRAWNS,u'marker-end'),
		(DRAWNS,u'marker-end-center'),
		(DRAWNS,u'marker-end-width'),
		(DRAWNS,u'marker-start'),
		(DRAWNS,u'marker-start-center'),
		(DRAWNS,u'marker-start-width'),
		(DRAWNS,u'measure-align'),
		(DRAWNS,u'measure-vertical-align'),
		(DRAWNS,u'ole-draw-aspect'),
		(DRAWNS,u'opacity'),
		(DRAWNS,u'opacity-name'),
		(DRAWNS,u'parallel'),
		(DRAWNS,u'placing'),
		(DRAWNS,u'red'),
		(DRAWNS,u'secondary-fill-color'),
		(DRAWNS,u'shadow'),
		(DRAWNS,u'shadow-color'),
		(DRAWNS,u'shadow-offset-x'),
		(DRAWNS,u'shadow-offset-y'),
		(DRAWNS,u'shadow-opacity'),
		(DRAWNS,u'show-unit'),
		(DRAWNS,u'start-guide'),
		(DRAWNS,u'start-line-spacing-horizontal'),
		(DRAWNS,u'start-line-spacing-vertical'),
		(DRAWNS,u'stroke'),
		(DRAWNS,u'stroke-dash'),
		(DRAWNS,u'stroke-dash-names'),
		(DRAWNS,u'stroke-linejoin'),
		(DRAWNS,u'symbol-color'),
		(DRAWNS,u'textarea-horizontal-align'),
		(DRAWNS,u'textarea-vertical-align'),
		(DRAWNS,u'tile-repeat-offset'),
		(DRAWNS,u'unit'),
		(DRAWNS,u'visible-area-height'),
		(DRAWNS,u'visible-area-left'),
		(DRAWNS,u'visible-area-top'),
		(DRAWNS,u'visible-area-width'),
		(DRAWNS,u'wrap-influence-on-position'),
# allowed_attributes
		(FONS,u'background-color'),
		(FONS,u'border'),
		(FONS,u'border-bottom'),
		(FONS,u'border-left'),
		(FONS,u'border-right'),
		(FONS,u'border-top'),
		(FONS,u'clip'),
		(FONS,u'margin'),
		(FONS,u'margin-bottom'),
		(FONS,u'margin-left'),
		(FONS,u'margin-right'),
		(FONS,u'margin-top'),
		(FONS,u'max-height'),
		(FONS,u'max-width'),
		(FONS,u'min-height'),
		(FONS,u'min-width'),
		(FONS,u'padding'),
		(FONS,u'padding-bottom'),
		(FONS,u'padding-left'),
		(FONS,u'padding-right'),
		(FONS,u'padding-top'),
		(FONS,u'wrap-option'),
		(STYLENS,u'background-transparency'),
		(STYLENS,u'border-line-width'),
		(STYLENS,u'border-line-width-bottom'),
		(STYLENS,u'border-line-width-left'),
		(STYLENS,u'border-line-width-right'),
		(STYLENS,u'border-line-width-top'),
		(STYLENS,u'editable'),
		(STYLENS,u'flow-with-text'),
		(STYLENS,u'horizontal-pos'),
		(STYLENS,u'horizontal-rel'),
		(STYLENS,u'mirror'),
		(STYLENS,u'number-wrapped-paragraphs'),
		(STYLENS,u'overflow-behavior'),
		(STYLENS,u'print-content'),
		(STYLENS,u'protect'),
		(STYLENS,u'rel-height'),
		(STYLENS,u'rel-width'),
		(STYLENS,u'repeat'),
		(STYLENS,u'run-through'),
		(STYLENS,u'shadow'),
		(STYLENS,u'shrink-to-fit'),
		(STYLENS,u'vertical-pos'),
		(STYLENS,u'vertical-rel'),
		(STYLENS,u'wrap'),
		(STYLENS,u'wrap-contour'),
		(STYLENS,u'wrap-contour-mode'),
		(STYLENS,u'wrap-dynamic-threshold'),
		(STYLENS,u'writing-mode'),
		(SVGNS,u'fill-rule'),
		(SVGNS,u'height'),
		(SVGNS,u'stroke-color'),
		(SVGNS,u'stroke-linecap'),
		(SVGNS,u'stroke-opacity'),
		(SVGNS,u'stroke-width'),
		(SVGNS,u'width'),
		(SVGNS,u'x'),
		(SVGNS,u'y'),
		(TEXTNS,u'anchor-page-number'),
		(TEXTNS,u'anchor-type'),
		(TEXTNS,u'animation'),
		(TEXTNS,u'animation-delay'),
		(TEXTNS,u'animation-direction'),
		(TEXTNS,u'animation-repeat'),
		(TEXTNS,u'animation-start-inside'),
		(TEXTNS,u'animation-steps'),
		(TEXTNS,u'animation-stop-inside'),
	),
	(STYLENS,u'handout-master'):(
		(PRESENTATIONNS,u'presentation-page-layout-name'),
		(STYLENS,u'page-layout-name'),
		(PRESENTATIONNS,u'use-footer-name'),
		(DRAWNS,u'style-name'),
		(PRESENTATIONNS,u'use-header-name'),
		(PRESENTATIONNS,u'use-date-time-name'),
	),
# allowed_attributes
	(STYLENS,u'header'):(
		(STYLENS,u'display'),
	),
	(STYLENS,u'header-footer-properties'): (
		(FONS,u'background-color'),
		(FONS,u'border'),
		(FONS,u'border-bottom'),
		(FONS,u'border-left'),
		(FONS,u'border-right'),
		(FONS,u'border-top'),
		(FONS,u'margin'),
		(FONS,u'margin-bottom'),
		(FONS,u'margin-left'),
		(FONS,u'margin-right'),
		(FONS,u'margin-top'),
		(FONS,u'min-height'),
		(FONS,u'padding'),
		(FONS,u'padding-bottom'),
		(FONS,u'padding-left'),
		(FONS,u'padding-right'),
		(FONS,u'padding-top'),
		(STYLENS,u'border-line-width'),
		(STYLENS,u'border-line-width-bottom'),
		(STYLENS,u'border-line-width-left'),
		(STYLENS,u'border-line-width-right'),
		(STYLENS,u'border-line-width-top'),
		(STYLENS,u'dynamic-spacing'),
		(STYLENS,u'shadow'),
		(SVGNS,u'height'),
	),
	(STYLENS,u'header-left'):(
		(STYLENS,u'display'),
	),
	(STYLENS,u'header-style'):(
	),
	(STYLENS,u'list-level-label-alignment'):(
		(FONS,u'text-indent'),
		(TEXTNS,u'label-followed-by'),
		(TEXTNS,u'list-tab-stop-position'),
		(FONS,u'margin-left'),
	),
# allowed_attributes
	(STYLENS,u'list-level-properties'): (
		(FONS,u'height'),
		(FONS,u'text-align'),
		(FONS,u'width'),
		(STYLENS,u'font-name'),
		(STYLENS,u'vertical-pos'),
		(STYLENS,u'vertical-rel'),
		(SVGNS,u'y'),
		(TEXTNS,u'list-level-position-and-space-mode'),
		(TEXTNS,u'min-label-distance'),
		(TEXTNS,u'min-label-width'),
		(TEXTNS,u'space-before'),
	),
	(STYLENS,u'map'):(
		(STYLENS,u'apply-style-name'),
		(STYLENS,u'base-cell-address'),
		(STYLENS,u'condition'),
	),
	(STYLENS,u'master-page'):(
		(STYLENS,u'page-layout-name'),
		(STYLENS,u'display-name'),
		(DRAWNS,u'style-name'),
		(STYLENS,u'name'),
		(STYLENS,u'next-style-name'),
	),
	(STYLENS,u'page-layout'):(
		(STYLENS,u'name'),
		(STYLENS,u'page-usage'),
	),
# allowed_attributes
	(STYLENS,u'page-layout-properties'): (
		(FONS,u'background-color'),
		(FONS,u'border'),
		(FONS,u'border-bottom'),
		(FONS,u'border-left'),
		(FONS,u'border-right'),
		(FONS,u'border-top'),
		(FONS,u'margin'),
		(FONS,u'margin-bottom'),
		(FONS,u'margin-left'),
		(FONS,u'margin-right'),
		(FONS,u'margin-top'),
		(FONS,u'padding'),
		(FONS,u'padding-bottom'),
		(FONS,u'padding-left'),
		(FONS,u'padding-right'),
		(FONS,u'padding-top'),
		(FONS,u'page-height'),
		(FONS,u'page-width'),
		(STYLENS,u'border-line-width'),
		(STYLENS,u'border-line-width-bottom'),
		(STYLENS,u'border-line-width-left'),
		(STYLENS,u'border-line-width-right'),
		(STYLENS,u'border-line-width-top'),
		(STYLENS,u'first-page-number'),
		(STYLENS,u'footnote-max-height'),
		(STYLENS,u'layout-grid-base-height'),
		(STYLENS,u'layout-grid-base-width'),
		(STYLENS,u'layout-grid-color'),
		(STYLENS,u'layout-grid-display'),
		(STYLENS,u'layout-grid-lines'),
		(STYLENS,u'layout-grid-mode'),
		(STYLENS,u'layout-grid-print'),
		(STYLENS,u'layout-grid-ruby-below'),
		(STYLENS,u'layout-grid-ruby-height'),
		(STYLENS,u'layout-grid-snap-to'),
		(STYLENS,u'layout-grid-standard-mode'),
		(STYLENS,u'num-format'),
		(STYLENS,u'num-letter-sync'),
		(STYLENS,u'num-prefix'),
		(STYLENS,u'num-suffix'),
		(STYLENS,u'paper-tray-name'),
		(STYLENS,u'print'),
		(STYLENS,u'print-orientation'),
		(STYLENS,u'print-page-order'),
		(STYLENS,u'register-truth-ref-style-name'),
		(STYLENS,u'scale-to'),
		(STYLENS,u'scale-to-pages'),
		(STYLENS,u'shadow'),
		(STYLENS,u'table-centering'),
		(STYLENS,u'writing-mode'),
                (LOEXTNS,u'scale-to-X'),
                (LOEXTNS,u'scale-to-Y')
	),
# allowed_attributes
	(STYLENS,u'paragraph-properties'): (
		(FONS,u'background-color'),
		(FONS,u'border'),
		(FONS,u'border-bottom'),
		(FONS,u'border-left'),
		(FONS,u'border-right'),
		(FONS,u'border-top'),
		(FONS,u'break-after'),
		(FONS,u'break-before'),
		(FONS,u'hyphenation-keep'),
		(FONS,u'hyphenation-ladder-count'),
		(FONS,u'keep-together'),
		(FONS,u'keep-with-next'),
		(FONS,u'line-height'),
		(FONS,u'margin'),
		(FONS,u'margin-bottom'),
		(FONS,u'margin-left'),
		(FONS,u'margin-right'),
		(FONS,u'margin-top'),
		(FONS,u'orphans'),
		(FONS,u'padding'),
		(FONS,u'padding-bottom'),
		(FONS,u'padding-left'),
		(FONS,u'padding-right'),
		(FONS,u'padding-top'),
		(FONS,u'text-align'),
		(FONS,u'text-align-last'),
		(FONS,u'text-indent'),
		(FONS,u'widows'),
		(LOEXTNS,u'contextual-spacing'),
		(STYLENS,u'auto-text-indent'),
		(STYLENS,u'background-transparency'),
		(STYLENS,u'border-line-width'),
		(STYLENS,u'border-line-width-bottom'),
		(STYLENS,u'border-line-width-left'),
		(STYLENS,u'border-line-width-right'),
		(STYLENS,u'border-line-width-top'),
		(STYLENS,u'font-independent-line-spacing'),
		(STYLENS,u'join-border'),
		(STYLENS,u'justify-single-word'),
		(STYLENS,u'line-break'),
		(STYLENS,u'line-height-at-least'),
		(STYLENS,u'line-spacing'),
		(STYLENS,u'page-number'),
		(STYLENS,u'punctuation-wrap'),
		(STYLENS,u'register-true'),
		(STYLENS,u'shadow'),
		(STYLENS,u'snap-to-layout-grid'),
		(STYLENS,u'tab-stop-distance'),
		(STYLENS,u'text-autospace'),
		(STYLENS,u'vertical-align'),
		(STYLENS,u'writing-mode'),
		(STYLENS,u'writing-mode-automatic'),
		(TEXTNS,u'line-number'),
		(TEXTNS,u'number-lines'),
	),
	(STYLENS,u'presentation-page-layout'):(
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
	),
# allowed_attributes
	(STYLENS,u'region-center'):(
	),
	(STYLENS,u'region-left'):(
	),
	(STYLENS,u'region-right'):(
	),
	(STYLENS,u'ruby-properties'): (
		(STYLENS,u'ruby-position'),
		(STYLENS,u'ruby-align'),
	),
	(STYLENS,u'section-properties'): (
		(FONS,u'background-color'),
		(FONS,u'margin-left'),
		(FONS,u'margin-right'),
		(STYLENS,u'editable'),
		(STYLENS,u'protect'),
		(STYLENS,u'writing-mode'),
		(TEXTNS,u'dont-balance-text-columns'),
	),
	(STYLENS,u'style'):(
		(STYLENS,u'auto-update'),
		(STYLENS,u'class'),
		(STYLENS,u'data-style-name'),
		(STYLENS,u'default-outline-level'),
		(STYLENS,u'display-name'),
		(STYLENS,u'family'),
		(STYLENS,u'list-level'),
		(STYLENS,u'list-style-name'),
		(STYLENS,u'master-page-name'),
		(STYLENS,u'name'),
		(STYLENS,u'next-style-name'),
		(STYLENS,u'parent-style-name'),
		(STYLENS,u'percentage-data-style-name'),
	),
# allowed_attributes
	(STYLENS,u'tab-stop'):(
		(STYLENS,u'leader-text-style'),
		(STYLENS,u'leader-width'),
		(STYLENS,u'leader-style'),
		(STYLENS,u'char'),
		(STYLENS,u'leader-color'),
		(STYLENS,u'position'),
		(STYLENS,u'leader-text'),
		(STYLENS,u'type'),
		(STYLENS,u'leader-type'),
	),
	(STYLENS,u'tab-stops'):(
	),
	(STYLENS,u'table-cell-properties'): (
		(FONS,u'background-color'),
		(FONS,u'border'),
		(FONS,u'border-bottom'),
		(FONS,u'border-left'),
		(FONS,u'border-right'),
		(FONS,u'border-top'),
		(FONS,u'padding'),
		(FONS,u'padding-bottom'),
		(FONS,u'padding-left'),
		(FONS,u'padding-right'),
		(FONS,u'padding-top'),
		(FONS,u'wrap-option'),
		(STYLENS,u'border-line-width'),
		(STYLENS,u'border-line-width-bottom'),
		(STYLENS,u'border-line-width-left'),
		(STYLENS,u'border-line-width-right'),
		(STYLENS,u'border-line-width-top'),
		(STYLENS,u'cell-protect'),
		(STYLENS,u'decimal-places'),
		(STYLENS,u'diagonal-bl-tr'),
		(STYLENS,u'diagonal-bl-tr-widths'),
		(STYLENS,u'diagonal-tl-br'),
		(STYLENS,u'diagonal-tl-br-widths'),
		(STYLENS,u'direction'),
		(STYLENS,u'glyph-orientation-vertical'),
		(STYLENS,u'print-content'),
		(STYLENS,u'repeat-content'),
		(STYLENS,u'rotation-align'),
		(STYLENS,u'rotation-angle'),
		(STYLENS,u'shadow'),
		(STYLENS,u'shrink-to-fit'),
		(STYLENS,u'text-align-source'),
		(STYLENS,u'vertical-align'),
		(STYLENS,u'writing-mode'),
	),
# allowed_attributes
	(STYLENS,u'table-column-properties'): (
		(FONS,u'break-after'),
		(FONS,u'break-before'),
		(STYLENS,u'column-width'),
		(STYLENS,u'rel-column-width'),
		(STYLENS,u'use-optimal-column-width'),
	),
	(STYLENS,u'table-properties'): (
		(FONS,u'background-color'),
		(FONS,u'break-after'),
		(FONS,u'break-before'),
		(FONS,u'keep-with-next'),
		(FONS,u'margin'),
		(FONS,u'margin-bottom'),
		(FONS,u'margin-left'),
		(FONS,u'margin-right'),
		(FONS,u'margin-top'),
		(STYLENS,u'may-break-between-rows'),
		(STYLENS,u'page-number'),
		(STYLENS,u'rel-width'),
		(STYLENS,u'shadow'),
		(STYLENS,u'width'),
		(STYLENS,u'writing-mode'),
		(TABLENS,u'align'),
		(TABLENS,u'border-model'),
		(TABLENS,u'display'),
	),
	(STYLENS,u'table-row-properties'): (
		(FONS,u'background-color'),
		(FONS,u'break-after'),
		(FONS,u'break-before'),
		(FONS,u'keep-together'),
		(STYLENS,u'min-row-height'),
		(STYLENS,u'row-height'),
		(STYLENS,u'use-optimal-row-height'),
	),
# allowed_attributes
	(STYLENS,u'text-properties'): (
		(FONS,u'background-color'),
		(FONS,u'color'),
		(FONS,u'country'),
		(FONS,u'font-family'),
		(FONS,u'font-size'),
		(FONS,u'font-style'),
		(FONS,u'font-variant'),
		(FONS,u'font-weight'),
		(FONS,u'hyphenate'),
		(FONS,u'hyphenation-push-char-count'),
		(FONS,u'hyphenation-remain-char-count'),
		(FONS,u'language'),
		(FONS,u'letter-spacing'),
		(FONS,u'script'),
		(FONS,u'text-shadow'),
		(FONS,u'text-transform'),
		(STYLENS,u'country-asian'),
		(STYLENS,u'country-complex'),
		(STYLENS,u'font-charset'),
		(STYLENS,u'font-charset-asian'),
		(STYLENS,u'font-charset-complex'),
		(STYLENS,u'font-family-asian'),
		(STYLENS,u'font-family-complex'),
		(STYLENS,u'font-family-generic'),
		(STYLENS,u'font-family-generic-asian'),
		(STYLENS,u'font-family-generic-complex'),
		(STYLENS,u'font-name'),
		(STYLENS,u'font-name-asian'),
		(STYLENS,u'font-name-complex'),
		(STYLENS,u'font-pitch'),
		(STYLENS,u'font-pitch-asian'),
		(STYLENS,u'font-pitch-complex'),
		(STYLENS,u'font-relief'),
		(STYLENS,u'font-size-asian'),
		(STYLENS,u'font-size-complex'),
		(STYLENS,u'font-size-rel'),
		(STYLENS,u'font-size-rel-asian'),
		(STYLENS,u'font-size-rel-complex'),
		(STYLENS,u'font-style-asian'),
		(STYLENS,u'font-style-complex'),
		(STYLENS,u'font-style-name'),
		(STYLENS,u'font-style-name-asian'),
		(STYLENS,u'font-style-name-complex'),
		(STYLENS,u'font-weight-asian'),
		(STYLENS,u'font-weight-complex'),
		(STYLENS,u'language-asian'),
		(STYLENS,u'language-complex'),
		(STYLENS,u'letter-kerning'),
		(STYLENS,u'rfc-language-tag'),
		(STYLENS,u'rfc-language-tag-asian'),
		(STYLENS,u'rfc-language-tag-complex'),
		(STYLENS,u'script-asian'),
		(STYLENS,u'script-complex'),
		(STYLENS,u'script-type'),
		(STYLENS,u'text-blinking'),
		(STYLENS,u'text-combine'),
		(STYLENS,u'text-combine-end-char'),
		(STYLENS,u'text-combine-start-char'),
		(STYLENS,u'text-emphasize'),
		(STYLENS,u'text-line-through-color'),
		(STYLENS,u'text-line-through-mode'),
		(STYLENS,u'text-line-through-style'),
		(STYLENS,u'text-line-through-text'),
		(STYLENS,u'text-line-through-text-style'),
		(STYLENS,u'text-line-through-type'),
		(STYLENS,u'text-line-through-width'),
		(STYLENS,u'text-outline'),
		(STYLENS,u'text-overline-color'),
		(STYLENS,u'text-overline-mode'),
		(STYLENS,u'text-overline-style'),
		(STYLENS,u'text-overline-type'),
		(STYLENS,u'text-overline-width'),
		(STYLENS,u'text-position'),
		(STYLENS,u'text-rotation-angle'),
		(STYLENS,u'text-rotation-scale'),
		(STYLENS,u'text-scale'),
		(STYLENS,u'text-underline-color'),
		(STYLENS,u'text-underline-mode'),
		(STYLENS,u'text-underline-style'),
		(STYLENS,u'text-underline-type'),
		(STYLENS,u'text-underline-width'),
		(STYLENS,u'use-window-font-color'),
		(TEXTNS,u'condition'),
		(TEXTNS,u'display'),
	),
	(SVGNS,u'definition-src'):(
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(SVGNS,u'desc'):(
	),
	(SVGNS,u'font-face-format'):(
		(SVGNS,u'string'),
	),
# allowed_attributes
	(SVGNS,u'font-face-name'):(
		(SVGNS,u'name'),
	),
	(SVGNS,u'font-face-src'):(
	),
	(SVGNS,u'font-face-uri'):(
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(SVGNS,u'linearGradient'):(
		(SVGNS,u'y2'),
		(DRAWNS,u'name'),
		(SVGNS,u'spreadMethod'),
		(SVGNS,u'gradientUnits'),
		(SVGNS,u'x2'),
		(SVGNS,u'gradientTransform'),
		(SVGNS,u'y1'),
		(DRAWNS,u'display-name'),
		(SVGNS,u'x1'),
	),
	(SVGNS,u'radialGradient'):(
		(DRAWNS,u'name'),
		(SVGNS,u'fx'),
		(SVGNS,u'fy'),
		(SVGNS,u'spreadMethod'),
		(SVGNS,u'gradientUnits'),
		(SVGNS,u'cy'),
		(SVGNS,u'cx'),
		(SVGNS,u'gradientTransform'),
		(DRAWNS,u'display-name'),
		(SVGNS,u'r'),
	),
	(SVGNS,u'stop'):(
		(SVGNS,u'stop-color'),
		(SVGNS,u'stop-opacity'),
		(SVGNS,u'offset'),
	),
	(SVGNS,u'title'):(
	),
# allowed_attributes
	(TABLENS,u'background'):(
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'body'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'calculation-settings'):(
		(TABLENS,u'automatic-find-labels'),
		(TABLENS,u'case-sensitive'),
		(TABLENS,u'search-criteria-must-apply-to-whole-cell'),
		(TABLENS,u'precision-as-shown'),
		(TABLENS,u'use-regular-expressions'),
		(TABLENS,u'use-wildcards'),
		(TABLENS,u'null-year'),
	),
	(TABLENS,u'cell-address'):(
		(TABLENS,u'column'),
		(TABLENS,u'table'),
		(TABLENS,u'row'),
	),
	(TABLENS,u'cell-content-change'):(
		(TABLENS,u'id'),
		(TABLENS,u'rejecting-change-id'),
		(TABLENS,u'acceptance-state'),
	),
# allowed_attributes
	(TABLENS,u'cell-content-deletion'):(
		(TABLENS,u'id'),
	),
	(TABLENS,u'cell-range-source'):(
		(TABLENS,u'last-row-spanned'),
		(TABLENS,u'last-column-spanned'),
		(TABLENS,u'name'),
		(TABLENS,u'filter-options'),
		(XLINKNS,u'actuate'),
		(TABLENS,u'filter-name'),
		(XLINKNS,u'href'),
		(TABLENS,u'refresh-delay'),
		(XLINKNS,u'type'),
	),
	(TABLENS,u'change-deletion'):(
		(TABLENS,u'id'),
	),
	(TABLENS,u'change-track-table-cell'):(
		(OFFICENS,u'string-value'),
		(TABLENS,u'cell-address'),
		(TABLENS,u'number-matrix-columns-spanned'),
		(TABLENS,u'number-matrix-rows-spanned'),
		(TABLENS,u'matrix-covered'),
		(OFFICENS,u'value-type'),
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(OFFICENS,u'value'),
		(TABLENS,u'formula'),
		(OFFICENS,u'time-value'),
	),
	(TABLENS,u'consolidation'):(
		(TABLENS,u'function'),
		(TABLENS,u'source-cell-range-addresses'),
		(TABLENS,u'target-cell-address'),
		(TABLENS,u'link-to-source-data'),
		(TABLENS,u'use-labels'),
	),
	(TABLENS,u'content-validation'):(
		(TABLENS,u'base-cell-address'),
		(TABLENS,u'display-list'),
		(TABLENS,u'allow-empty-cell'),
		(TABLENS,u'name'),
		(TABLENS,u'condition'),
	),
	(TABLENS,u'content-validations'):(
	),
# allowed_attributes
	(TABLENS,u'covered-table-cell'):(
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(OFFICENS,u'string-value'),
		(OFFICENS,u'time-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'value-type'),
		(TABLENS,u'content-validation-name'),
		(TABLENS,u'formula'),
		(TABLENS,u'number-columns-repeated'),
		(TABLENS,u'protect'),
		(TABLENS,u'protected'),
		(TABLENS,u'style-name'),
		(XHTMLNS,u'about'),
		(XHTMLNS,u'content'),
		(XHTMLNS,u'datatype'),
		(XHTMLNS,u'property'),
		(XMLNS,u'id'),
	),
	(TABLENS,u'cut-offs'):(
	),
	(TABLENS,u'data-pilot-display-info'):(
		(TABLENS,u'member-count'),
		(TABLENS,u'data-field'),
		(TABLENS,u'enabled'),
		(TABLENS,u'display-member-mode'),
	),
	(TABLENS,u'data-pilot-field'):(
		(TABLENS,u'selected-page'),
		(TABLENS,u'function'),
		(TABLENS,u'orientation'),
		(TABLENS,u'used-hierarchy'),
		(TABLENS,u'is-data-layout-field'),
		(TABLENS,u'source-field-name'),
	),
	(TABLENS,u'data-pilot-field-reference'):(
		(TABLENS,u'member-name'),
		(TABLENS,u'field-name'),
		(TABLENS,u'member-type'),
		(TABLENS,u'type'),
	),
# allowed_attributes
	(TABLENS,u'data-pilot-group'):(
		(TABLENS,u'name'),
	),
	(TABLENS,u'data-pilot-group-member'):(
		(TABLENS,u'name'),
	),
	(TABLENS,u'data-pilot-groups'):(
		(TABLENS,u'date-end'),
		(TABLENS,u'end'),
		(TABLENS,u'start'),
		(TABLENS,u'source-field-name'),
		(TABLENS,u'step'),
		(TABLENS,u'date-start'),
		(TABLENS,u'grouped-by'),
	),
	(TABLENS,u'data-pilot-layout-info'):(
		(TABLENS,u'add-empty-lines'),
		(TABLENS,u'layout-mode'),
	),
	(TABLENS,u'data-pilot-level'):(
		(TABLENS,u'show-empty'),
	),
# allowed_attributes
	(TABLENS,u'data-pilot-member'):(
		(TABLENS,u'show-details'),
		(TABLENS,u'name'),
		(TABLENS,u'display'),
	),
	(TABLENS,u'data-pilot-members'):(
	),
	(TABLENS,u'data-pilot-sort-info'):(
		(TABLENS,u'data-field'),
		(TABLENS,u'sort-mode'),
		(TABLENS,u'order'),
	),
	(TABLENS,u'data-pilot-subtotal'):(
		(TABLENS,u'function'),
	),
	(TABLENS,u'data-pilot-subtotals'):(
	),
	(TABLENS,u'data-pilot-table'):(
		(TABLENS,u'buttons'),
		(TABLENS,u'application-data'),
		(TABLENS,u'name'),
		(TABLENS,u'drill-down-on-double-click'),
		(TABLENS,u'target-range-address'),
		(TABLENS,u'ignore-empty-rows'),
		(TABLENS,u'identify-categories'),
		(TABLENS,u'show-filter-button'),
		(TABLENS,u'grand-total'),
	),
# allowed_attributes
	(TABLENS,u'data-pilot-tables'):(
	),
	(TABLENS,u'database-range'):(
		(TABLENS,u'orientation'),
		(TABLENS,u'target-range-address'),
		(TABLENS,u'contains-header'),
		(TABLENS,u'on-update-keep-size'),
		(TABLENS,u'name'),
		(TABLENS,u'is-selection'),
		(TABLENS,u'refresh-delay'),
		(TABLENS,u'display-filter-buttons'),
		(TABLENS,u'has-persistent-data'),
		(TABLENS,u'on-update-keep-styles'),
	),
	(TABLENS,u'database-ranges'):(
	),
	(TABLENS,u'database-source-query'):(
		(TABLENS,u'query-name'),
		(TABLENS,u'database-name'),
	),
# allowed_attributes
	(TABLENS,u'database-source-sql'):(
		(TABLENS,u'parse-sql-statement'),
		(TABLENS,u'database-name'),
		(TABLENS,u'sql-statement'),
	),
	(TABLENS,u'database-source-table'):(
		(TABLENS,u'database-table-name'),
		(TABLENS,u'database-name'),
	),
	(TABLENS,u'dde-link'):(
	),
	(TABLENS,u'dde-links'):(
	),
	(TABLENS,u'deletion'):(
		(TABLENS,u'rejecting-change-id'),
		(TABLENS,u'multi-deletion-spanned'),
		(TABLENS,u'acceptance-state'),
		(TABLENS,u'table'),
		(TABLENS,u'position'),
		(TABLENS,u'type'),
		(TABLENS,u'id'),
	),
# allowed_attributes
	(TABLENS,u'deletions'):(
	),
	(TABLENS,u'dependencies'):(
	),
	(TABLENS,u'dependency'):(
		(TABLENS,u'id'),
	),
	(TABLENS,u'desc'):(
	),
	(TABLENS,u'detective'):(
	),
	(TABLENS,u'error-macro'):(
		(TABLENS,u'execute'),
	),
	(TABLENS,u'error-message'):(
		(TABLENS,u'display'),
		(TABLENS,u'message-type'),
		(TABLENS,u'title'),
	),
	(TABLENS,u'even-columns'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'even-rows'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
# allowed_attributes
	(TABLENS,u'filter'):(
		(TABLENS,u'target-range-address'),
		(TABLENS,u'display-duplicates'),
		(TABLENS,u'condition-source-range-address'),
		(TABLENS,u'condition-source'),
	),
	(TABLENS,u'filter-and'):(
	),
	(TABLENS,u'filter-condition'):(
		(TABLENS,u'operator'),
		(TABLENS,u'field-number'),
		(TABLENS,u'data-type'),
		(TABLENS,u'case-sensitive'),
		(TABLENS,u'value'),
	),
	(TABLENS,u'filter-or'):(
	),
# allowed_attributes
	(TABLENS,u'filter-set-item'):(
		(TABLENS,u'value'),
	),
	(TABLENS,u'first-column'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'first-row'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
# allowed_attributes
	(TABLENS,u'help-message'):(
		(TABLENS,u'display'),
		(TABLENS,u'title'),
	),
	(TABLENS,u'highlighted-range'):(
		(TABLENS,u'contains-error'),
		(TABLENS,u'direction'),
		(TABLENS,u'marked-invalid'),
		(TABLENS,u'cell-range-address'),
	),
	(TABLENS,u'insertion'):(
		(TABLENS,u'count'),
		(TABLENS,u'rejecting-change-id'),
		(TABLENS,u'acceptance-state'),
		(TABLENS,u'table'),
		(TABLENS,u'position'),
		(TABLENS,u'type'),
		(TABLENS,u'id'),
	),
	(TABLENS,u'insertion-cut-off'):(
		(TABLENS,u'position'),
		(TABLENS,u'id'),
	),
	(TABLENS,u'iteration'):(
		(TABLENS,u'status'),
		(TABLENS,u'maximum-difference'),
		(TABLENS,u'steps'),
	),
# allowed_attributes
	(TABLENS,u'label-range'):(
		(TABLENS,u'label-cell-range-address'),
		(TABLENS,u'data-cell-range-address'),
		(TABLENS,u'orientation'),
	),
	(TABLENS,u'label-ranges'):(
	),
	(TABLENS,u'last-column'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'last-row'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'movement'):(
		(TABLENS,u'id'),
		(TABLENS,u'rejecting-change-id'),
		(TABLENS,u'acceptance-state'),
	),
	(TABLENS,u'movement-cut-off'):(
		(TABLENS,u'position'),
		(TABLENS,u'end-position'),
		(TABLENS,u'start-position'),
	),
	(TABLENS,u'named-expression'):(
		(TABLENS,u'base-cell-address'),
		(TABLENS,u'expression'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'named-expressions'):(
	),
	(TABLENS,u'named-range'):(
		(TABLENS,u'range-usable-as'),
		(TABLENS,u'base-cell-address'),
		(TABLENS,u'name'),
		(TABLENS,u'cell-range-address'),
	),
	(TABLENS,u'null-date'):(
		(TABLENS,u'date-value'),
		(TABLENS,u'value-type'),
	),
	(TABLENS,u'odd-columns'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'odd-rows'):(
		(TABLENS,u'paragraph-style-name'),
		(TABLENS,u'style-name'),
	),
	(TABLENS,u'operation'):(
		(TABLENS,u'index'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'previous'):(
		(TABLENS,u'id'),
	),
	(TABLENS,u'scenario'):(
		(TABLENS,u'comment'),
		(TABLENS,u'border-color'),
		(TABLENS,u'copy-back'),
		(TABLENS,u'is-active'),
		(TABLENS,u'protected'),
		(TABLENS,u'copy-formulas'),
		(TABLENS,u'copy-styles'),
		(TABLENS,u'scenario-ranges'),
		(TABLENS,u'display-border'),
	),
	(TABLENS,u'shapes'):(
	),
# allowed_attributes
	(TABLENS,u'sort'):(
		(TABLENS,u'case-sensitive'),
		(TABLENS,u'embedded-number-behavior'),
		(TABLENS,u'algorithm'),
		(TABLENS,u'target-range-address'),
		(TABLENS,u'country'),
		(TABLENS,u'language'),
		(TABLENS,u'bind-styles-to-content'),
		(TABLENS,u'rfc-language-tag'),
		(TABLENS,u'script'),
	),
	(TABLENS,u'sort-by'):(
		(TABLENS,u'field-number'),
		(TABLENS,u'data-type'),
		(TABLENS,u'order'),
	),
	(TABLENS,u'sort-groups'):(
		(TABLENS,u'data-type'),
		(TABLENS,u'order'),
	),
	(TABLENS,u'source-cell-range'):(
		(TABLENS,u'cell-range-address'),
	),
	(TABLENS,u'source-range-address'):(
		(TABLENS,u'column'),
		(TABLENS,u'end-column'),
		(TABLENS,u'start-table'),
		(TABLENS,u'end-row'),
		(TABLENS,u'table'),
		(TABLENS,u'start-row'),
		(TABLENS,u'row'),
		(TABLENS,u'end-table'),
		(TABLENS,u'start-column'),
	),
# allowed_attributes
	(TABLENS,u'source-service'):(
		(TABLENS,u'user-name'),
		(TABLENS,u'source-name'),
		(TABLENS,u'password'),
		(TABLENS,u'object-name'),
		(TABLENS,u'name'),
	),
	(TABLENS,u'subtotal-field'):(
		(TABLENS,u'function'),
		(TABLENS,u'field-number'),
	),
	(TABLENS,u'subtotal-rule'):(
		(TABLENS,u'group-by-field-number'),
	),
	(TABLENS,u'subtotal-rules'):(
		(TABLENS,u'bind-styles-to-content'),
		(TABLENS,u'page-breaks-on-group-change'),
		(TABLENS,u'case-sensitive'),
	),
	(TABLENS,u'table'):(
		(TABLENS,u'is-sub-table'),
		(TABLENS,u'name'),
		(TABLENS,u'print'),
		(TABLENS,u'print-ranges'),
		(TABLENS,u'protected'),
		(TABLENS,u'protection-key'),
		(TABLENS,u'protection-key-digest-algorithm'),
		(TABLENS,u'style-name'),
		(TABLENS,u'template-name'),
		(TABLENS,u'use-banding-columns-styles'),
		(TABLENS,u'use-banding-rows-styles'),
		(TABLENS,u'use-first-column-styles'),
		(TABLENS,u'use-first-row-styles'),
		(TABLENS,u'use-last-column-styles'),
		(TABLENS,u'use-last-row-styles'),
		(XMLNS,u'id'),
	),
	(TABLENS,u'table-cell'):(
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(OFFICENS,u'string-value'),
		(OFFICENS,u'time-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'value-type'),
		(TABLENS,u'content-validation-name'),
		(TABLENS,u'formula'),
		(TABLENS,u'number-columns-repeated'),
		(TABLENS,u'number-columns-spanned'),
		(TABLENS,u'number-matrix-columns-spanned'),
		(TABLENS,u'number-matrix-rows-spanned'),
		(TABLENS,u'number-rows-spanned'),
		(TABLENS,u'protect'),
		(TABLENS,u'protected'),
		(TABLENS,u'style-name'),
		(XHTMLNS,u'about'),
		(XHTMLNS,u'content'),
		(XHTMLNS,u'datatype'),
		(XHTMLNS,u'property'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(TABLENS,u'table-column'):(
		(TABLENS,u'style-name'),
		(TABLENS,u'default-cell-style-name'),
		(TABLENS,u'visibility'),
		(TABLENS,u'number-columns-repeated'),
		(XMLNS,u'id'),
	),
	(TABLENS,u'table-column-group'):(
		(TABLENS,u'display'),
	),
	(TABLENS,u'table-columns'):(
	),
	(TABLENS,u'table-header-columns'):(
	),
	(TABLENS,u'table-header-rows'):(
	),
	(TABLENS,u'table-row'):(
		(TABLENS,u'number-rows-repeated'),
		(TABLENS,u'style-name'),
		(TABLENS,u'visibility'),
		(TABLENS,u'default-cell-style-name'),
		(XMLNS,u'id'),
	),
	(TABLENS,u'table-row-group'):(
		(TABLENS,u'display'),
	),
	(TABLENS,u'table-rows'):(
	),
	(TABLENS,u'table-source'):(
		(TABLENS,u'filter-options'),
		(XLINKNS,u'actuate'),
		(TABLENS,u'filter-name'),
		(XLINKNS,u'href'),
		(TABLENS,u'mode'),
		(TABLENS,u'table-name'),
		(XLINKNS,u'type'),
		(TABLENS,u'refresh-delay'),
	),
# allowed_attributes
	(TABLENS,u'table-template'):(
		(TABLENS,u'last-row-end-column'),
		(TABLENS,u'first-row-end-column'),
		(TABLENS,u'name'),
		(TABLENS,u'last-row-start-column'),
		(TABLENS,u'first-row-start-column'),
	),
	(TABLENS,u'target-range-address'):(
		(TABLENS,u'column'),
		(TABLENS,u'end-column'),
		(TABLENS,u'start-table'),
		(TABLENS,u'end-row'),
		(TABLENS,u'table'),
		(TABLENS,u'start-row'),
		(TABLENS,u'row'),
		(TABLENS,u'end-table'),
		(TABLENS,u'start-column'),
	),
	(TABLENS,u'title'):(
	),
	(TABLENS,u'tracked-changes'):(
		(TABLENS,u'track-changes'),
	),
# allowed_attributes
	(TEXTNS,u'a'):(
		(OFFICENS,u'name'),
		(OFFICENS,u'target-frame-name'),
		(OFFICENS,u'title'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'visited-style-name'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'show'),
		(XLINKNS,u'type'),
	),
	(TEXTNS,u'alphabetical-index'):(
		(TEXTNS,u'name'),
		(TEXTNS,u'protected'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(TEXTNS,u'style-name'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'alphabetical-index-auto-mark-file'):(
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
	),
	(TEXTNS,u'alphabetical-index-entry-template'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'alphabetical-index-mark'):(
		(TEXTNS,u'main-entry'),
		(TEXTNS,u'key1-phonetic'),
		(TEXTNS,u'key2'),
		(TEXTNS,u'key1'),
		(TEXTNS,u'string-value'),
		(TEXTNS,u'key2-phonetic'),
		(TEXTNS,u'string-value-phonetic'),
	),
# allowed_attributes
	(TEXTNS,u'alphabetical-index-mark-end'):(
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'alphabetical-index-mark-start'):(
		(TEXTNS,u'main-entry'),
		(TEXTNS,u'key1-phonetic'),
		(TEXTNS,u'key2'),
		(TEXTNS,u'key1'),
		(TEXTNS,u'string-value-phonetic'),
		(TEXTNS,u'key2-phonetic'),
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'alphabetical-index-source'):(
		(FONS,u'country'),
		(FONS,u'language'),
		(FONS,u'script'),
		(STYLENS,u'rfc-language-tag'),
		(TEXTNS,u'alphabetical-separators'),
		(TEXTNS,u'capitalize-entries'),
		(TEXTNS,u'combine-entries'),
		(TEXTNS,u'combine-entries-with-dash'),
		(TEXTNS,u'combine-entries-with-pp'),
		(TEXTNS,u'comma-separated'),
		(TEXTNS,u'ignore-case'),
		(TEXTNS,u'index-scope'),
		(TEXTNS,u'main-entry-style-name'),
		(TEXTNS,u'relative-tab-stop-position'),
		(TEXTNS,u'sort-algorithm'),
		(TEXTNS,u'use-keys-as-entries'),
	),
# allowed_attributes
	(TEXTNS,u'author-initials'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'author-name'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'bibliography'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'bibliography-configuration'):(
		(FONS,u'country'),
		(FONS,u'language'),
		(FONS,u'script'),
		(STYLENS,u'rfc-language-tag'),
		(TEXTNS,u'numbered-entries'),
		(TEXTNS,u'prefix'),
		(TEXTNS,u'sort-algorithm'),
		(TEXTNS,u'sort-by-position'),
		(TEXTNS,u'suffix'),
	),
	(TEXTNS,u'bibliography-entry-template'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'bibliography-type'),
	),
# allowed_attributes
	(TEXTNS,u'bibliography-mark'):(
		(TEXTNS,u'address'),
		(TEXTNS,u'annote'),
		(TEXTNS,u'author'),
		(TEXTNS,u'bibliography-type'),
		(TEXTNS,u'booktitle'),
		(TEXTNS,u'chapter'),
		(TEXTNS,u'custom1'),
		(TEXTNS,u'custom2'),
		(TEXTNS,u'custom3'),
		(TEXTNS,u'custom4'),
		(TEXTNS,u'custom5'),
		(TEXTNS,u'edition'),
		(TEXTNS,u'editor'),
		(TEXTNS,u'howpublished'),
		(TEXTNS,u'identifier'),
		(TEXTNS,u'institution'),
		(TEXTNS,u'isbn'),
		(TEXTNS,u'issn'),
		(TEXTNS,u'journal'),
		(TEXTNS,u'month'),
		(TEXTNS,u'note'),
		(TEXTNS,u'number'),
		(TEXTNS,u'organizations'),
		(TEXTNS,u'pages'),
		(TEXTNS,u'publisher'),
		(TEXTNS,u'report-type'),
		(TEXTNS,u'school'),
		(TEXTNS,u'series'),
		(TEXTNS,u'title'),
		(TEXTNS,u'url'),
		(TEXTNS,u'volume'),
		(TEXTNS,u'year'),
	),
	(TEXTNS,u'bibliography-source'):(
	),
	(TEXTNS,u'bookmark'):(
		(TEXTNS,u'name'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'bookmark-end'):(
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'bookmark-ref'):(
		(TEXTNS,u'ref-name'),
		(TEXTNS,u'reference-format'),
	),
	(TEXTNS,u'bookmark-start'):(
		(XHTMLNS,u'about'),
		(XHTMLNS,u'content'),
		(XHTMLNS,u'datatype'),
		(XHTMLNS,u'property'),
		(TEXTNS,u'name'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(TEXTNS,u'change'):(
		(TEXTNS,u'change-id'),
	),
	(TEXTNS,u'change-end'):(
		(TEXTNS,u'change-id'),
	),
	(TEXTNS,u'change-start'):(
		(TEXTNS,u'change-id'),
	),
	(TEXTNS,u'changed-region'):(
		(TEXTNS,u'id'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'chapter'):(
		(TEXTNS,u'display'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'conditional-text'):(
		(TEXTNS,u'string-value-if-true'),
		(TEXTNS,u'current-value'),
		(TEXTNS,u'string-value-if-false'),
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'creation-date'):(
		(TEXTNS,u'date-value'),
		(TEXTNS,u'fixed'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'creation-time'):(
		(TEXTNS,u'fixed'),
		(TEXTNS,u'time-value'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'creator'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'database-display'):(
		(TEXTNS,u'column-name'),
		(TEXTNS,u'table-name'),
		(TEXTNS,u'table-type'),
		(TEXTNS,u'database-name'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'database-name'):(
		(TEXTNS,u'table-name'),
		(TEXTNS,u'table-type'),
		(TEXTNS,u'database-name'),
	),
	(TEXTNS,u'database-next'):(
		(TEXTNS,u'table-name'),
		(TEXTNS,u'table-type'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'database-row-number'):(
		(STYLENS,u'num-format'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'value'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'table-name'),
		(TEXTNS,u'table-type'),
	),
	(TEXTNS,u'database-row-select'):(
		(TEXTNS,u'row-number'),
		(TEXTNS,u'table-name'),
		(TEXTNS,u'table-type'),
		(TEXTNS,u'database-name'),
		(TEXTNS,u'condition'),
	),
# allowed_attributes
	(TEXTNS,u'date'):(
		(TEXTNS,u'date-value'),
		(TEXTNS,u'fixed'),
		(TEXTNS,u'date-adjust'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'dde-connection'):(
		(TEXTNS,u'connection-name'),
	),
	(TEXTNS,u'dde-connection-decl'):(
		(OFFICENS,u'automatic-update'),
		(OFFICENS,u'dde-topic'),
		(OFFICENS,u'dde-application'),
		(OFFICENS,u'name'),
		(OFFICENS,u'dde-item'),
	),
	(TEXTNS,u'dde-connection-decls'):(
	),
	(TEXTNS,u'deletion'):(
	),
	(TEXTNS,u'description'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'editing-cycles'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'editing-duration'):(
		(TEXTNS,u'duration'),
		(TEXTNS,u'fixed'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'execute-macro'):(
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'expression'):(
		(TEXTNS,u'display'),
		(OFFICENS,u'string-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(STYLENS,u'data-style-name'),
		(OFFICENS,u'value-type'),
		(TEXTNS,u'formula'),
		(OFFICENS,u'time-value'),
	),
	(TEXTNS,u'file-name'):(
		(TEXTNS,u'fixed'),
		(TEXTNS,u'display'),
	),
# allowed_attributes
	(TEXTNS,u'format-change'):(
	),
	(TEXTNS,u'h'):(
		(TEXTNS,u'restart-numbering'),
		(TEXTNS,u'cond-style-name'),
		(TEXTNS,u'is-list-header'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'class-names'),
		(TEXTNS,u'start-value'),
		(TEXTNS,u'id'),
		(TEXTNS,u'outline-level'),
		(XHTMLNS,u'about'),
		(XHTMLNS,u'content'),
		(XHTMLNS,u'datatype'),
		(XHTMLNS,u'property'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'hidden-paragraph'):(
		(TEXTNS,u'is-hidden'),
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'hidden-text'):(
		(TEXTNS,u'string-value'),
		(TEXTNS,u'is-hidden'),
		(TEXTNS,u'condition'),
	),
	(TEXTNS,u'illustration-index'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'illustration-index-entry-template'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'illustration-index-source'):(
		(TEXTNS,u'index-scope'),
		(TEXTNS,u'caption-sequence-name'),
		(TEXTNS,u'use-caption'),
		(TEXTNS,u'caption-sequence-format'),
		(TEXTNS,u'relative-tab-stop-position'),
	),
	(TEXTNS,u'index-body'):(
	),
	(TEXTNS,u'index-entry-bibliography'):(
		(TEXTNS,u'bibliography-data-field'),
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-entry-chapter'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
		(TEXTNS,u'display'),
	),
# allowed_attributes
	(TEXTNS,u'index-entry-link-end'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-entry-link-start'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-entry-page-number'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-entry-span'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-entry-tab-stop'):(
		(STYLENS,u'position'),
		(TEXTNS,u'style-name'),
		(STYLENS,u'type'),
		(STYLENS,u'leader-char'),
	),
	(TEXTNS,u'index-entry-text'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-source-style'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'index-source-styles'):(
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'index-title'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'index-title-template'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'initial-creator'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'insertion'):(
	),
# allowed_attributes
	(TEXTNS,u'keywords'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'line-break'):(
	),
	(TEXTNS,u'linenumbering-configuration'):(
		(TEXTNS,u'number-position'),
		(TEXTNS,u'number-lines'),
		(STYLENS,u'num-format'),
		(TEXTNS,u'count-empty-lines'),
		(TEXTNS,u'count-in-text-boxes'),
		(TEXTNS,u'style-name'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'increment'),
		(TEXTNS,u'offset'),
		(TEXTNS,u'restart-on-page'),
	),
	(TEXTNS,u'linenumbering-separator'):(
		(TEXTNS,u'increment'),
	),
	(TEXTNS,u'list'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'continue-numbering'),
		(TEXTNS,u'continue-list'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'list-header'):(
		(XMLNS,u'id'),
	),
# allowed_attributes
	(TEXTNS,u'list-item'):(
		(TEXTNS,u'start-value'),
		(TEXTNS,u'style-override'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'list-level-style-bullet'):(
		(TEXTNS,u'level'),
		(STYLENS,u'num-prefix'),
		(STYLENS,u'num-suffix'),
		(TEXTNS,u'bullet-relative-size'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'bullet-char'),
	),
	(TEXTNS,u'list-level-style-image'):(
		(XLINKNS,u'show'),
		(XLINKNS,u'actuate'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(TEXTNS,u'level'),
	),
	(TEXTNS,u'list-level-style-number'):(
		(TEXTNS,u'level'),
		(TEXTNS,u'display-levels'),
		(STYLENS,u'num-format'),
		(STYLENS,u'num-suffix'),
		(TEXTNS,u'style-name'),
		(STYLENS,u'num-prefix'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'start-value'),
	),
# allowed_attributes
	(TEXTNS,u'list-style'):(
		(TEXTNS,u'consecutive-numbering'),
		(STYLENS,u'display-name'),
		(STYLENS,u'name'),
	),
	(TEXTNS,u'measure'):(
		(TEXTNS,u'kind'),
	),
# allowed_attributes
	(TEXTNS,u'meta'):(
		(XHTMLNS,u'about'),
		(XHTMLNS,u'content'),
		(XMLNS,u'id'),
		(XHTMLNS,u'property'),
		(XHTMLNS,u'datatype'),
	),
# allowed_attributes
	(TEXTNS,u'meta-field'):(
		(STYLENS,u'data-style-name'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'modification-date'):(
		(TEXTNS,u'date-value'),
		(TEXTNS,u'fixed'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'modification-time'):(
		(TEXTNS,u'fixed'),
		(TEXTNS,u'time-value'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'note'):(
		(TEXTNS,u'note-class'),
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'note-body'):(
	),
	(TEXTNS,u'note-citation'):(
		(TEXTNS,u'label'),
	),
	(TEXTNS,u'note-continuation-notice-backward'):(
	),
	(TEXTNS,u'note-continuation-notice-forward'):(
	),
	(TEXTNS,u'note-ref'):(
		(TEXTNS,u'ref-name'),
		(TEXTNS,u'note-class'),
		(TEXTNS,u'reference-format'),
	),
	(TEXTNS,u'notes-configuration'):(
		(TEXTNS,u'citation-body-style-name'),
		(STYLENS,u'num-format'),
		(TEXTNS,u'default-style-name'),
		(STYLENS,u'num-suffix'),
		(TEXTNS,u'start-numbering-at'),
		(STYLENS,u'num-prefix'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'citation-style-name'),
		(TEXTNS,u'footnotes-position'),
		(TEXTNS,u'master-page-name'),
		(TEXTNS,u'start-value'),
		(TEXTNS,u'note-class'),
	),
	(TEXTNS,u'number'):(
	),
	(TEXTNS,u'numbered-paragraph'):(
		(TEXTNS,u'continue-numbering'),
		(TEXTNS,u'level'),
		(TEXTNS,u'list-id'),
		(TEXTNS,u'start-value'),
		(TEXTNS,u'style-name'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'object-count'):(
		(STYLENS,u'num-format'),
		(STYLENS,u'num-letter-sync'),
	),
	(TEXTNS,u'object-index'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
# allowed_attributes
	(TEXTNS,u'object-index-entry-template'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'object-index-source'):(
		(TEXTNS,u'use-draw-objects'),
		(TEXTNS,u'use-math-objects'),
		(TEXTNS,u'relative-tab-stop-position'),
		(TEXTNS,u'use-chart-objects'),
		(TEXTNS,u'index-scope'),
		(TEXTNS,u'use-spreadsheet-objects'),
		(TEXTNS,u'use-other-objects'),
	),
	(TEXTNS,u'outline-level-style'):(
		(TEXTNS,u'level'),
		(TEXTNS,u'display-levels'),
		(STYLENS,u'num-format'),
		(STYLENS,u'num-suffix'),
		(TEXTNS,u'style-name'),
		(STYLENS,u'num-prefix'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'start-value'),
	),
	(TEXTNS,u'outline-style'):(
		(STYLENS,u'name'),
	),
# allowed_attributes
	(TEXTNS,u'p'):(
		(TEXTNS,u'class-names'),
		(TEXTNS,u'cond-style-name'),
		(TEXTNS,u'id'),
		(TEXTNS,u'style-name'),
		(XHTMLNS,u'about'),
		(XHTMLNS,u'content'),
		(XHTMLNS,u'datatype'),
		(XHTMLNS,u'property'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'page'):(
		(TEXTNS,u'master-page-name'),
	),
	(TEXTNS,u'page-continuation'):(
		(TEXTNS,u'string-value'),
		(TEXTNS,u'select-page'),
	),
	(TEXTNS,u'page-number'):(
		(TEXTNS,u'page-adjust'),
		(STYLENS,u'num-format'),
		(TEXTNS,u'fixed'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'select-page'),
	),
	(TEXTNS,u'page-sequence'):(
	),
	(TEXTNS,u'page-variable-get'):(
		(STYLENS,u'num-format'),
		(STYLENS,u'num-letter-sync'),
	),
	(TEXTNS,u'page-variable-set'):(
		(TEXTNS,u'active'),
		(TEXTNS,u'page-adjust'),
	),
	(TEXTNS,u'placeholder'):(
		(TEXTNS,u'placeholder-type'),
		(TEXTNS,u'description'),
	),
	(TEXTNS,u'print-date'):(
		(TEXTNS,u'date-value'),
		(TEXTNS,u'fixed'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'print-time'):(
		(TEXTNS,u'fixed'),
		(TEXTNS,u'time-value'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'printed-by'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'reference-mark'):(
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'reference-mark-end'):(
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'reference-mark-start'):(
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'ruby'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'ruby-base'):(
	),
	(TEXTNS,u'ruby-text'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u's'):(
		(TEXTNS,u'c'),
	),
	(TEXTNS,u'script'):(
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(SCRIPTNS,u'language'),
	),
# allowed_attributes
	(TEXTNS,u'section'):(
		(TEXTNS,u'condition'),
		(TEXTNS,u'display'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protected'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(TEXTNS,u'style-name'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'section-source'):(
		(TEXTNS,u'filter-name'),
		(XLINKNS,u'href'),
		(XLINKNS,u'type'),
		(TEXTNS,u'section-name'),
		(XLINKNS,u'show'),
	),
	(TEXTNS,u'sender-city'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-company'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-country'):(
		(TEXTNS,u'fixed'),
	),
# allowed_attributes
	(TEXTNS,u'sender-email'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-fax'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-firstname'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-initials'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-lastname'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-phone-private'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-phone-work'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-position'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-postal-code'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-state-or-province'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-street'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sender-title'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'sequence'):(
		(TEXTNS,u'formula'),
		(STYLENS,u'num-format'),
		(STYLENS,u'num-letter-sync'),
		(TEXTNS,u'name'),
		(TEXTNS,u'ref-name'),
	),
	(TEXTNS,u'sequence-decl'):(
		(TEXTNS,u'separation-character'),
		(TEXTNS,u'display-outline-level'),
		(TEXTNS,u'name'),
	),
	(TEXTNS,u'sequence-decls'):(
	),
	(TEXTNS,u'sequence-ref'):(
		(TEXTNS,u'ref-name'),
		(TEXTNS,u'reference-format'),
	),
	(TEXTNS,u'sheet-name'):(
	),
	(TEXTNS,u'soft-page-break'):(
	),
	(TEXTNS,u'sort-key'):(
		(TEXTNS,u'sort-ascending'),
		(TEXTNS,u'key'),
	),
# allowed_attributes
	(TEXTNS,u'span'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'class-names'),
	),
	(TEXTNS,u'subject'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'tab'):(
		(TEXTNS,u'tab-ref'),
	),
	(TEXTNS,u'table-formula'):(
		(TEXTNS,u'formula'),
		(STYLENS,u'data-style-name'),
		(TEXTNS,u'display'),
	),
	(TEXTNS,u'table-index'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'table-index-entry-template'):(
		(TEXTNS,u'style-name'),
	),
	(TEXTNS,u'table-index-source'):(
		(TEXTNS,u'index-scope'),
		(TEXTNS,u'caption-sequence-name'),
		(TEXTNS,u'use-caption'),
		(TEXTNS,u'caption-sequence-format'),
		(TEXTNS,u'relative-tab-stop-position'),
	),
# allowed_attributes
	(TEXTNS,u'table-of-content'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'table-of-content-entry-template'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'table-of-content-source'):(
		(TEXTNS,u'index-scope'),
		(TEXTNS,u'outline-level'),
		(TEXTNS,u'relative-tab-stop-position'),
		(TEXTNS,u'use-index-marks'),
		(TEXTNS,u'use-outline-level'),
		(TEXTNS,u'use-index-source-styles'),
	),
	(TEXTNS,u'template-name'):(
		(TEXTNS,u'display'),
	),
	(TEXTNS,u'text-input'):(
		(TEXTNS,u'description'),
	),
	(TEXTNS,u'time'):(
		(TEXTNS,u'time-adjust'),
		(TEXTNS,u'fixed'),
		(TEXTNS,u'time-value'),
		(STYLENS,u'data-style-name'),
	),
	(TEXTNS,u'title'):(
		(TEXTNS,u'fixed'),
	),
	(TEXTNS,u'toc-mark'):(
		(TEXTNS,u'string-value'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'toc-mark-end'):(
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'toc-mark-start'):(
		(TEXTNS,u'id'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'tracked-changes'):(
		(TEXTNS,u'track-changes'),
	),
	(TEXTNS,u'user-defined'):(
		(TEXTNS,u'name'),
		(OFFICENS,u'string-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'date-value'),
		(STYLENS,u'data-style-name'),
		(TEXTNS,u'fixed'),
		(OFFICENS,u'time-value'),
	),
	(TEXTNS,u'user-field-decl'):(
		(TEXTNS,u'name'),
		(OFFICENS,u'string-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(OFFICENS,u'value-type'),
		(TEXTNS,u'formula'),
		(OFFICENS,u'time-value'),
	),
	(TEXTNS,u'user-field-decls'):(
	),
	(TEXTNS,u'user-field-get'):(
		(STYLENS,u'data-style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'display'),
	),
# allowed_attributes
	(TEXTNS,u'user-field-input'):(
		(STYLENS,u'data-style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'description'),
	),
	(TEXTNS,u'user-index'):(
		(TEXTNS,u'protected'),
		(TEXTNS,u'style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'protection-key'),
		(TEXTNS,u'protection-key-digest-algorithm'),
		(XMLNS,u'id'),
	),
	(TEXTNS,u'user-index-entry-template'):(
		(TEXTNS,u'style-name'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'user-index-mark'):(
		(TEXTNS,u'index-name'),
		(TEXTNS,u'string-value'),
		(TEXTNS,u'outline-level'),
	),
	(TEXTNS,u'user-index-mark-end'):(
		(TEXTNS,u'id'),
	),
	(TEXTNS,u'user-index-mark-start'):(
		(TEXTNS,u'index-name'),
		(TEXTNS,u'id'),
		(TEXTNS,u'outline-level'),
	),
# allowed_attributes
	(TEXTNS,u'user-index-source'):(
		(TEXTNS,u'copy-outline-levels'),
		(TEXTNS,u'index-name'),
		(TEXTNS,u'index-scope'),
		(TEXTNS,u'relative-tab-stop-position'),
		(TEXTNS,u'use-floating-frames'),
		(TEXTNS,u'use-graphics'),
		(TEXTNS,u'use-index-marks'),
		(TEXTNS,u'use-index-source-styles'),
		(TEXTNS,u'use-objects'),
		(TEXTNS,u'use-tables'),
	),
	(TEXTNS,u'variable-decl'):(
		(TEXTNS,u'name'),
		(OFFICENS,u'value-type'),
	),
	(TEXTNS,u'variable-decls'):(
	),
	(TEXTNS,u'variable-get'):(
		(STYLENS,u'data-style-name'),
		(TEXTNS,u'name'),
		(TEXTNS,u'display'),
	),
	(TEXTNS,u'variable-input'):(
		(STYLENS,u'data-style-name'),
		(TEXTNS,u'display'),
		(TEXTNS,u'name'),
		(OFFICENS,u'value-type'),
		(TEXTNS,u'description'),
	),
	(TEXTNS,u'variable-set'):(
		(TEXTNS,u'name'),
		(TEXTNS,u'display'),
		(OFFICENS,u'string-value'),
		(OFFICENS,u'value'),
		(OFFICENS,u'boolean-value'),
		(OFFICENS,u'currency'),
		(OFFICENS,u'date-value'),
		(STYLENS,u'data-style-name'),
		(OFFICENS,u'value-type'),
		(TEXTNS,u'formula'),
		(OFFICENS,u'time-value'),
	),
# allowed_attributes
}

''',
    },
    'odf.load': {
        'is_package': False,
        'source': r'''
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright (C) 2007-2008 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

# This script is to be embedded in opendocument.py later
# The purpose is to read an ODT/ODP/ODS file and create the datastructure
# in memory. The user should then be able to make operations and then save
# the structure again.

from defusedxml.sax import make_parser
from xml.sax import handler
from xml.sax.xmlreader import InputSource
import xml.sax.saxutils
from odf.element import Element
from odf.namespaces import OFFICENS
try:
    from cStringIO import StringIO
except ImportError:
    from io import StringIO

#
# Parse the XML files
#
class LoadParser(handler.ContentHandler):
    """ Extract headings from content.xml of an ODT file """
    triggers = (
       (OFFICENS, 'automatic-styles'), (OFFICENS, 'body'),
       (OFFICENS, 'font-face-decls'), (OFFICENS, 'master-styles'),
       (OFFICENS, 'meta'), (OFFICENS, 'scripts'),
       (OFFICENS, 'settings'), (OFFICENS, 'styles') )

    def __init__(self, document):
        self.doc = document
        self.data = []
        self.level = 0
        self.parse = False

    def characters(self, data):
        if self.parse == False:
            return
        self.data.append(data)

    def startElementNS(self, tag, qname, attrs):
        if tag in self.triggers:
            self.parse = True
        if self.doc._parsing != "styles.xml" and tag == (OFFICENS, 'font-face-decls'):
            self.parse = False
        if self.parse == False:
            return

        self.level = self.level + 1
        # Add any accumulated text content
        content = ''.join(self.data)
        if content:
            self.parent.addText(content, check_grammar=False)
            self.data = []
        # Create the element
        attrdict = {}
        for (att,value) in attrs.items():
            attrdict[att] = value
        try:
            e = Element(qname = tag, qattributes=attrdict, check_grammar=False)
            self.curr = e
        except AttributeError as v:
            print ("Error: %s" % v)

        if tag == (OFFICENS, 'automatic-styles'):
            e = self.doc.automaticstyles
        elif tag == (OFFICENS, 'body'):
            e = self.doc.body
        elif tag == (OFFICENS, 'master-styles'):
            e = self.doc.masterstyles
        elif tag == (OFFICENS, 'meta'):
            e = self.doc.meta
        elif tag == (OFFICENS,'scripts'):
            e = self.doc.scripts
        elif tag == (OFFICENS,'settings'):
            e = self.doc.settings
        elif tag == (OFFICENS,'styles'):
            e = self.doc.styles
        elif self.doc._parsing == "styles.xml" and tag == (OFFICENS, 'font-face-decls'):
            e = self.doc.fontfacedecls
        elif hasattr(self,'parent'):
            self.parent.addElement(e, check_grammar=False)
        self.parent = e


    def endElementNS(self, tag, qname):
        if self.parse == False:
            return
        self.level = self.level - 1
        str = ''.join(self.data)
        if str:
            self.curr.addText(str, check_grammar=False)
        self.data = []
        self.curr = self.curr.parentNode
        self.parent = self.curr
        if tag in self.triggers:
            self.parse = False

''',
    },
    'odf.manifest': {
        'is_package': False,
        'source': r"""
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#
#

import sys, os.path
sys.path.append(os.path.dirname(__file__))

from odf.namespaces import MANIFESTNS
from odf.element import Element

# Autogenerated
def Manifest(**args):
    return Element(qname = (MANIFESTNS,'manifest'), **args)

def FileEntry(**args):
    return Element(qname = (MANIFESTNS,'file-entry'), **args)

def EncryptionData(**args):
    return Element(qname = (MANIFESTNS,'encryption-data'), **args)

def Algorithm(**args):
    return Element(qname = (MANIFESTNS,'algorithm'), **args)

def KeyDerivation(**args):
    return Element(qname = (MANIFESTNS,'key-derivation'), **args)

""",
    },
    'odf.math': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import MATHNS
from odf.element import Element

# ODF 1.0 section 12.5
# Mathematical content is represented by MathML 2.0

# Autogenerated
def Math(**args):
    return Element(qname = (MATHNS,'math'), **args)


""",
    },
    'odf.meta': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import METANS
from odf.element import Element

# Autogenerated
def AutoReload(**args):
    return Element(qname = (METANS,'auto-reload'), **args)

def CreationDate(**args):
    return Element(qname = (METANS,'creation-date'), **args)

def DateString(**args):
    return Element(qname = (METANS,'date-string'), **args)

def DocumentStatistic(**args):
    return Element(qname = (METANS,'document-statistic'), **args)

def EditingCycles(**args):
    return Element(qname = (METANS,'editing-cycles'), **args)

def EditingDuration(**args):
    return Element(qname = (METANS,'editing-duration'), **args)

def Generator(**args):
    return Element(qname = (METANS,'generator'), **args)

def HyperlinkBehaviour(**args):
    return Element(qname = (METANS,'hyperlink-behaviour'), **args)

def InitialCreator(**args):
    return Element(qname = (METANS,'initial-creator'), **args)

def Keyword(**args):
    return Element(qname = (METANS,'keyword'), **args)

def PrintDate(**args):
    return Element(qname = (METANS,'print-date'), **args)

def PrintedBy(**args):
    return Element(qname = (METANS,'printed-by'), **args)

def Template(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (METANS,'template'), **args)

def UserDefined(**args):
    return Element(qname = (METANS,'user-defined'), **args)


""",
    },
    'odf.namespaces': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#
__version__ = "1.4.1"

TOOLSVERSION = u"ODFPY/" + __version__

ANIMNS         = u"urn:oasis:names:tc:opendocument:xmlns:animation:1.0"
CHARTNS        = u"urn:oasis:names:tc:opendocument:xmlns:chart:1.0"
CHARTOOONS     = u"http://openoffice.org/2010/chart"
CONFIGNS       = u"urn:oasis:names:tc:opendocument:xmlns:config:1.0"
CSS3TNS        = u"http://www.w3.org/TR/css3-text/"
#DBNS           = u"http://openoffice.org/2004/database"
DBNS           = u"urn:oasis:names:tc:opendocument:xmlns:database:1.0"
DCNS           = u"http://purl.org/dc/elements/1.1/"
DOMNS          = u"http://www.w3.org/2001/xml-events"
DR3DNS         = u"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"
DRAWNS         = u"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
FIELDNS        = u"urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"
FONS           = u"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
FORMNS         = u"urn:oasis:names:tc:opendocument:xmlns:form:1.0"
FORMXNS        = u"urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"
GRDDLNS        = u"http://www.w3.org/2003/g/data-view#"
KOFFICENS      = u"http://www.koffice.org/2005/"
LOEXTNS        = u"urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"
MANIFESTNS     = u"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
MATHNS         = u"http://www.w3.org/1998/Math/MathML"
METANS         = u"urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
NUMBERNS       = u"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
OFFICENS       = u"urn:oasis:names:tc:opendocument:xmlns:office:1.0"
OFNS           = u"urn:oasis:names:tc:opendocument:xmlns:of:1.2"
OOOCNS         = u"http://openoffice.org/2004/calc"
OOONS          = u"http://openoffice.org/2004/office"
OOOWNS         = u"http://openoffice.org/2004/writer"
PRESENTATIONNS = u"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
RDFANS         = u"http://docs.oasis-open.org/opendocument/meta/rdfa#"
RPTNS          = u"http://openoffice.org/2005/report"
SCRIPTNS       = u"urn:oasis:names:tc:opendocument:xmlns:script:1.0"
SMILNS         = u"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"
STYLENS        = u"urn:oasis:names:tc:opendocument:xmlns:style:1.0"
SVGNS          = u"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
TABLENS        = u"urn:oasis:names:tc:opendocument:xmlns:table:1.0"
TABLEOOONS     = u"http://openoffice.org/2009/table"
TEXTNS         = u"urn:oasis:names:tc:opendocument:xmlns:text:1.0"
XFORMSNS       = u"http://www.w3.org/2002/xforms"
XHTMLNS        = u"http://www.w3.org/1999/xhtml"
XLINKNS        = u"http://www.w3.org/1999/xlink"
XMLNS          = u"http://www.w3.org/XML/1998/namespace"
XSDNS          = u"http://www.w3.org/2001/XMLSchema"
XSINS          = u"http://www.w3.org/2001/XMLSchema-instance"

nsdict = {
   ANIMNS: u'anim',
   CHARTNS: u'chart',
   CHARTOOONS: u'chartooo',
   CONFIGNS: u'config',
   CSS3TNS: u'css3t',
   DBNS: u'db',
   DCNS: u'dc',
   DOMNS: u'dom',
   DR3DNS: u'dr3d',
   DRAWNS: u'draw',
   FIELDNS: u'field',
   FONS: u'fo',
   FORMNS: u'form',
   FORMXNS: u'formx',
   GRDDLNS: u'grddl',
   KOFFICENS: u'koffice',
   LOEXTNS: u'loext',
   MANIFESTNS: u'manifest',
   MATHNS: u'math',
   METANS: u'meta',
   NUMBERNS: u'number',
   OFFICENS: u'office',
   OFNS: u'of',
   OOONS: u'ooo',
   OOOWNS: u'ooow',
   OOOCNS: u'oooc',
   PRESENTATIONNS: u'presentation',
   RDFANS: u'rdfa',
   RPTNS:  u'rpt',
   SCRIPTNS: u'script',
   SMILNS: u'smil',
   STYLENS: u'style',
   SVGNS: u'svg',
   TABLENS: u'table',
   TABLEOOONS: u'tableooo',
   TEXTNS: u'text',
   XFORMSNS: u'xforms',
   XLINKNS: u'xlink',
   XHTMLNS: u'xhtml',
   XMLNS: u'xml',
   XSDNS: u'xsd',
   XSINS: u'xsi',
}

""",
    },
    'odf.number': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import NUMBERNS
from odf.element import Element
from odf.style import StyleElement


# Autogenerated
def AmPm(**args):
    return Element(qname = (NUMBERNS,'am-pm'), **args)

def Boolean(**args):
    return Element(qname = (NUMBERNS,'boolean'), **args)

def BooleanStyle(**args):
    return StyleElement(qname = (NUMBERNS,'boolean-style'), **args)

def CurrencyStyle(**args):
    return StyleElement(qname = (NUMBERNS,'currency-style'), **args)

def CurrencySymbol(**args):
    return Element(qname = (NUMBERNS,'currency-symbol'), **args)

def DateStyle(**args):
    return StyleElement(qname = (NUMBERNS,'date-style'), **args)

def Day(**args):
    return Element(qname = (NUMBERNS,'day'), **args)

def DayOfWeek(**args):
    return Element(qname = (NUMBERNS,'day-of-week'), **args)

def EmbeddedText(**args):
    return Element(qname = (NUMBERNS,'embedded-text'), **args)

def Era(**args):
    return Element(qname = (NUMBERNS,'era'), **args)

def Fraction(**args):
    return Element(qname = (NUMBERNS,'fraction'), **args)

def Hours(**args):
    return Element(qname = (NUMBERNS,'hours'), **args)

def Minutes(**args):
    return Element(qname = (NUMBERNS,'minutes'), **args)

def Month(**args):
    return Element(qname = (NUMBERNS,'month'), **args)

def Number(**args):
    return Element(qname = (NUMBERNS,'number'), **args)

def NumberStyle(**args):
    return StyleElement(qname = (NUMBERNS,'number-style'), **args)

def PercentageStyle(**args):
    return StyleElement(qname = (NUMBERNS,'percentage-style'), **args)

def Quarter(**args):
    return Element(qname = (NUMBERNS,'quarter'), **args)

def ScientificNumber(**args):
    return Element(qname = (NUMBERNS,'scientific-number'), **args)

def Seconds(**args):
    return Element(qname = (NUMBERNS,'seconds'), **args)

def Text(**args):
    return Element(qname = (NUMBERNS,'text'), **args)

def TextContent(**args):
    return Element(qname = (NUMBERNS,'text-content'), **args)

def TextStyle(**args):
    return StyleElement(qname = (NUMBERNS,'text-style'), **args)

def TimeStyle(**args):
    return StyleElement(qname = (NUMBERNS,'time-style'), **args)

def WeekOfYear(**args):
    return Element(qname = (NUMBERNS,'week-of-year'), **args)

def Year(**args):
    return Element(qname = (NUMBERNS,'year'), **args)


""",
    },
    'odf.odf2moinmoin': {
        'is_package': False,
        'source': '# -*- coding: utf-8 -*-\n# Copyright (C) 2006-2008 Søren Roug, European Environment Agency\n#\n# This library is free software; you can redistribute it and/or\n# modify it under the terms of the GNU Lesser General Public\n# License as published by the Free Software Foundation; either\n# version 2.1 of the License, or (at your option) any later version.\n#\n# This library is distributed in the hope that it will be useful,\n# but WITHOUT ANY WARRANTY; without even the implied warranty of\n# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU\n# Lesser General Public License for more details.\n#\n# You should have received a copy of the GNU Lesser General Public\n# License along with this library; if not, write to the Free Software\n# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA\n#\n# See http://trac.edgewall.org/wiki/WikiFormatting\n#\n# Contributor(s):\n#\n\nimport sys, zipfile, xml.dom.minidom\nfrom odf.namespaces import nsdict\nfrom odf.elementtypes import *\n\nIGNORED_TAGS = [\n    \'draw:a\'\n    \'draw:g\',\n    \'draw:line\',\n    \'draw:object-ole\',\n    \'office:annotation\',\n    \'presentation:notes\',\n    \'svg:desc\',\n] + [ nsdict[item[0]]+":"+item[1] for item in empty_elements]\n\nINLINE_TAGS = [ nsdict[item[0]]+":"+item[1] for item in inline_elements]\n\n\nclass TextProps:\n    """ Holds properties for a text style. """\n\n    def __init__(self):\n\n        self.italic = False\n        self.bold = False\n        self.fixed = False\n        self.underlined = False\n        self.strikethrough = False\n        self.superscript = False\n        self.subscript = False\n\n    def setItalic(self, value):\n        if value == "italic":\n            self.italic = True\n        elif value == "normal":\n            self.italic = False\n\n    def setBold(self, value):\n        if value == "bold":\n            self.bold = True\n        elif value == "normal":\n            self.bold = False\n\n    def setFixed(self, value):\n        self.fixed = value\n\n    def setUnderlined(self, value):\n        if value and value != "none":\n            self.underlined = True\n\n    def setStrikethrough(self, value):\n        if value and value != "none":\n            self.strikethrough = True\n\n    def setPosition(self, value):\n        if value is None or value == \'\':\n            return\n        posisize = value.split(\' \')\n        textpos = posisize[0]\n        if textpos.find(\'%\') == -1:\n            if textpos == "sub":\n                self.superscript = False\n                self.subscript = True\n            elif textpos == "super":\n                self.superscript = True\n                self.subscript = False\n        else:\n            itextpos = int(textpos[:textpos.find(\'%\')])\n            if itextpos > 10:\n                self.superscript = False\n                self.subscript = True\n            elif itextpos < -10:\n                self.superscript = True\n                self.subscript = False\n\n    def __str__(self):\n\n        return "[italic=%s, bold=i%s, fixed=%s]" % (str(self.italic),\n                                          str(self.bold),\n                                          str(self.fixed))\n\nclass ParagraphProps:\n    """ Holds properties of a paragraph style. """\n\n    def __init__(self):\n\n        self.blockquote = False\n        self.headingLevel = 0\n        self.code = False\n        self.title = False\n        self.indented = 0\n\n    def setIndented(self, value):\n        self.indented = value\n\n    def setHeading(self, level):\n        self.headingLevel = level\n\n    def setTitle(self, value):\n        self.title = value\n\n    def setCode(self, value):\n        self.code = value\n\n\n    def __str__(self):\n\n        return "[bq=%s, h=%d, code=%s]" % (str(self.blockquote),\n                                           self.headingLevel,\n                                           str(self.code))\n\n\nclass ListProperties:\n    """ Holds properties for a list style. """\n\n    def __init__(self):\n        self.ordered = False\n\n    def setOrdered(self, value):\n        self.ordered = value\n\n\n\nclass ODF2MoinMoin(object):\n\n\n    def __init__(self, filepath):\n        self.footnotes = []\n        self.footnoteCounter = 0\n        self.textStyles = {"Standard": TextProps()}\n        self.paragraphStyles = {"Standard": ParagraphProps()}\n        self.listStyles = {}\n        self.fixedFonts = []\n        self.hasTitle = 0\n        self.lastsegment = None\n\n        # Tags\n        self.elements = {\n         \'draw:page\': self.textToString,\n         \'draw:frame\': self.textToString,\n         \'draw:image\': self.draw_image,\n         \'draw:text-box\': self.textToString,\n         \'text:a\': self.text_a,\n         \'text:note\': self.text_note,\n        }\n        for tag in IGNORED_TAGS:\n            self.elements[tag] = self.do_nothing\n\n        for tag in INLINE_TAGS:\n            self.elements[tag] = self.inline_markup\n        self.elements[\'text:line-break\'] = self.text_line_break\n        self.elements[\'text:s\'] = self.text_s\n        self.elements[\'text:tab\'] = self.text_tab\n\n        self.load(filepath)\n\n    def processFontDeclarations(self, fontDecl):\n        """ Extracts necessary font information from a font-declaration\n            element.\n            """\n        for fontFace in fontDecl.getElementsByTagName("style:font-face"):\n            if fontFace.getAttribute("style:font-pitch") == "fixed":\n                self.fixedFonts.append(fontFace.getAttribute("style:name"))\n\n\n\n    def extractTextProperties(self, style, parent=None):\n        """ Extracts text properties from a style element. """\n\n        textProps = TextProps()\n\n        if parent:\n            parentProp = self.textStyles.get(parent, None)\n            if parentProp:\n                textProp = parentProp\n\n        textPropEl = style.getElementsByTagName("style:text-properties")\n        if not textPropEl: return textProps\n\n        textPropEl = textPropEl[0]\n\n        textProps.setItalic(textPropEl.getAttribute("fo:font-style"))\n        textProps.setBold(textPropEl.getAttribute("fo:font-weight"))\n        textProps.setUnderlined(textPropEl.getAttribute("style:text-underline-style"))\n        textProps.setStrikethrough(textPropEl.getAttribute("style:text-line-through-style"))\n        textProps.setPosition(textPropEl.getAttribute("style:text-position"))\n\n        if textPropEl.getAttribute("style:font-name") in self.fixedFonts:\n            textProps.setFixed(True)\n\n        return textProps\n\n    def extractParagraphProperties(self, style, parent=None):\n        """ Extracts paragraph properties from a style element. """\n\n        paraProps = ParagraphProps()\n\n        name = style.getAttribute("style:name")\n\n        if name.startswith("Heading_20_"):\n            level = name[11:]\n            try:\n                level = int(level)\n                paraProps.setHeading(level)\n            except:\n                level = 0\n\n        if name == "Title":\n            paraProps.setTitle(True)\n\n        paraPropEl = style.getElementsByTagName("style:paragraph-properties")\n        if paraPropEl:\n            paraPropEl = paraPropEl[0]\n            leftMargin = paraPropEl.getAttribute("fo:margin-left")\n            if leftMargin:\n                try:\n                    leftMargin = float(leftMargin[:-2])\n                    if leftMargin > 0.01:\n                        paraProps.setIndented(True)\n                except:\n                    pass\n\n        textProps = self.extractTextProperties(style)\n        if textProps.fixed:\n            paraProps.setCode(True)\n\n        return paraProps\n\n\n    def processStyles(self, styleElements):\n        """ Runs through "style" elements extracting necessary information.\n        """\n\n        for style in styleElements:\n\n            name = style.getAttribute("style:name")\n\n            if name == "Standard": continue\n\n            family = style.getAttribute("style:family")\n            parent = style.getAttribute("style:parent-style-name")\n\n            if family == "text":\n                self.textStyles[name] = self.extractTextProperties(style, parent)\n\n            elif family == "paragraph":\n                self.paragraphStyles[name] = \\\n                                 self.extractParagraphProperties(style, parent)\n                self.textStyles[name] = self.extractTextProperties(style, parent)\n\n    def processListStyles(self, listStyleElements):\n\n        for style in listStyleElements:\n            name = style.getAttribute("style:name")\n\n            prop = ListProperties()\n            if style.hasChildNodes():\n                subitems = [el for el in style.childNodes\n                     if el.nodeType == xml.dom.Node.ELEMENT_NODE\n                     and el.tagName == "text:list-level-style-number"]\n                if len(subitems) > 0:\n                    prop.setOrdered(True)\n\n            self.listStyles[name] = prop\n\n\n    def load(self, filepath):\n        """ Loads an ODT file. """\n\n        zip = zipfile.ZipFile(filepath)\n\n        styles_doc = xml.dom.minidom.parseString(zip.read("styles.xml"))\n        fontfacedecls = styles_doc.getElementsByTagName("office:font-face-decls")\n        if fontfacedecls:\n            self.processFontDeclarations(fontfacedecls[0])\n        self.processStyles(styles_doc.getElementsByTagName("style:style"))\n        self.processListStyles(styles_doc.getElementsByTagName("text:list-style"))\n\n        self.content = xml.dom.minidom.parseString(zip.read("content.xml"))\n        fontfacedecls = self.content.getElementsByTagName("office:font-face-decls")\n        if fontfacedecls:\n            self.processFontDeclarations(fontfacedecls[0])\n\n        self.processStyles(self.content.getElementsByTagName("style:style"))\n        self.processListStyles(self.content.getElementsByTagName("text:list-style"))\n\n    def compressCodeBlocks(self, text):\n        """ Removes extra blank lines from code blocks. """\n\n        return text\n        lines = text.split("\\n")\n        buffer = []\n        numLines = len(lines)\n        for i in range(numLines):\n\n            if (lines[i].strip() or i == numLines-1  or i == 0 or\n                not ( lines[i-1].startswith("    ")\n                      and lines[i+1].startswith("    ") ) ):\n                buffer.append("\\n" + lines[i])\n\n        return \'\'.join(buffer)\n\n#-----------------------------------\n    def do_nothing(self, node):\n        return \'\'\n\n    def draw_image(self, node):\n        """\n        """\n\n        link = node.getAttribute("xlink:href")\n        if link and link[:2] == \'./\': # Indicates a sub-object, which isn\'t supported\n            return "%s\\n" % link\n        if link and link[:9] == \'Pictures/\':\n            link = link[9:]\n        return "[[Image(%s)]]\\n" % link\n\n    def text_a(self, node):\n        text = self.textToString(node)\n        link = node.getAttribute("xlink:href")\n        if link.strip() == text.strip():\n            return "[%s] " % link.strip()\n        else:\n            return "[%s %s] " % (link.strip(), text.strip())\n\n\n    def text_line_break(self, node):\n        return "[[BR]]"\n\n    def text_note(self, node):\n        cite = (node.getElementsByTagName("text:note-citation")[0]\n                    .childNodes[0].nodeValue)\n        body = (node.getElementsByTagName("text:note-body")[0]\n                    .childNodes[0])\n        self.footnotes.append((cite, self.textToString(body)))\n        return "^%s^" % cite\n\n    def text_s(self, node):\n        try:\n            num = int(node.getAttribute("text:c"))\n            return " "*num\n        except:\n            return " "\n\n    def text_tab(self, node):\n        return "    "\n\n    def inline_markup(self, node):\n        text = self.textToString(node)\n\n        if not text.strip():\n            return \'\'  # don\'t apply styles to white space\n\n        styleName = node.getAttribute("text:style-name")\n        style = self.textStyles.get(styleName, TextProps())\n\n        if style.fixed:\n            return "`" + text + "`"\n\n        mark = []\n        if style:\n            if style.italic:\n                mark.append("\'\'")\n            if style.bold:\n                mark.append("\'\'\'")\n            if style.underlined:\n                mark.append("__")\n            if style.strikethrough:\n                mark.append("~~")\n            if style.superscript:\n                mark.append("^")\n            if style.subscript:\n                mark.append(",,")\n        revmark = mark[:]\n        revmark.reverse()\n        return "%s%s%s" % (\'\'.join(mark), text, \'\'.join(revmark))\n\n#-----------------------------------\n    def listToString(self, listElement, indent = 0):\n\n        self.lastsegment = listElement.tagName\n        buffer = []\n\n        styleName = listElement.getAttribute("text:style-name")\n        props = self.listStyles.get(styleName, ListProperties())\n\n        i = 0\n        for item in listElement.childNodes:\n            buffer.append(" "*indent)\n            i += 1\n            if props.ordered:\n                number = str(i)\n                number = " " + number + ". "\n                buffer.append(" 1. ")\n            else:\n                buffer.append(" * ")\n            subitems = [el for el in item.childNodes\n                          if el.tagName in ["text:p", "text:h", "text:list"]]\n            for subitem in subitems:\n                if subitem.tagName == "text:list":\n                    buffer.append("\\n")\n                    buffer.append(self.listToString(subitem, indent+3))\n                else:\n                    buffer.append(self.paragraphToString(subitem, indent+3))\n                self.lastsegment = subitem.tagName\n            self.lastsegment = item.tagName\n            buffer.append("\\n")\n\n        return \'\'.join(buffer)\n\n    def tableToString(self, tableElement):\n        """ MoinMoin uses || to delimit table cells\n        """\n\n        self.lastsegment = tableElement.tagName\n        buffer = []\n\n        for item in tableElement.childNodes:\n            self.lastsegment = item.tagName\n            if item.tagName == "table:table-header-rows":\n                buffer.append(self.tableToString(item))\n            if item.tagName == "table:table-row":\n                buffer.append("\\n||")\n                for cell in item.childNodes:\n                    buffer.append(self.inline_markup(cell))\n                    buffer.append("||")\n                    self.lastsegment = cell.tagName\n        return \'\'.join(buffer)\n\n\n    def toString(self):\n        """ Converts the document to a string.\n            FIXME: Result from second call differs from first call\n        """\n        body = self.content.getElementsByTagName("office:body")[0]\n        text = body.childNodes[0]\n\n        buffer = []\n\n        paragraphs = [el for el in text.childNodes\n                      if el.tagName in ["draw:page", "text:p", "text:h","text:section",\n                                        "text:list", "table:table"]]\n\n        for paragraph in paragraphs:\n            if paragraph.tagName == "text:list":\n                text = self.listToString(paragraph)\n            elif paragraph.tagName == "text:section":\n                text = self.textToString(paragraph)\n            elif paragraph.tagName == "table:table":\n                text = self.tableToString(paragraph)\n            else:\n                text = self.paragraphToString(paragraph)\n            if text:\n                buffer.append(text)\n\n        if self.footnotes:\n\n            buffer.append("----")\n            for cite, body in self.footnotes:\n                buffer.append("%s: %s" % (cite, body))\n\n\n        buffer.append("")\n        return self.compressCodeBlocks(\'\\n\'.join(buffer))\n\n\n    def textToString(self, element):\n\n        buffer = []\n\n        for node in element.childNodes:\n\n            if node.nodeType == xml.dom.Node.TEXT_NODE:\n                buffer.append(node.nodeValue)\n\n            elif node.nodeType == xml.dom.Node.ELEMENT_NODE:\n                tag = node.tagName\n\n                if tag in ("draw:text-box", "draw:frame"):\n                    buffer.append(self.textToString(node))\n\n                elif tag in ("text:p", "text:h"):\n                    text = self.paragraphToString(node)\n                    if text:\n                        buffer.append(text)\n                elif tag == "text:list":\n                    buffer.append(self.listToString(node))\n                else:\n                    method = self.elements.get(tag)\n                    if method:\n                        buffer.append(method(node))\n                    else:\n                        buffer.append(" {" + tag + "} ")\n\n        return \'\'.join(buffer)\n\n    def paragraphToString(self, paragraph, indent = 0):\n\n        dummyParaProps = ParagraphProps()\n\n        style_name = paragraph.getAttribute("text:style-name")\n        paraProps = self.paragraphStyles.get(style_name, dummyParaProps)\n        text = self.inline_markup(paragraph)\n\n        if paraProps and not paraProps.code:\n            text = text.strip()\n\n        if paragraph.tagName == "text:p" and self.lastsegment == "text:p":\n            text = "\\n" + text\n\n        self.lastsegment = paragraph.tagName\n\n        if paraProps.title:\n            self.hasTitle = 1\n            return "= " + text + " =\\n"\n\n        outlinelevel = paragraph.getAttribute("text:outline-level")\n        if outlinelevel:\n\n            level = int(outlinelevel)\n            if self.hasTitle: level += 1\n\n            if level >= 1:\n                return "=" * level + " " + text + " " + "=" * level + "\\n"\n\n        elif paraProps.code:\n            return "{{{\\n" + text + "\\n}}}\\n"\n\n        if paraProps.indented:\n            return self.wrapParagraph(text, indent = indent, blockquote = True)\n\n        else:\n            return self.wrapParagraph(text, indent = indent)\n\n\n    def wrapParagraph(self, text, indent = 0, blockquote=False):\n\n        counter = 0\n        buffer = []\n        LIMIT = 50\n\n        if blockquote:\n            buffer.append("  ")\n\n        return \'\'.join(buffer) + text\n        # Unused from here\n        for token in text.split():\n\n            if counter > LIMIT - indent:\n                buffer.append("\\n" + " "*indent)\n                if blockquote:\n                    buffer.append("  ")\n                counter = 0\n\n            buffer.append(token + " ")\n            counter += len(token)\n\n        return \'\'.join(buffer)\n',
    },
    'odf.odf2xhtml': {
        'is_package': False,
        'source': '#!/usr/bin/python\n# -*- coding: utf-8 -*-\n# Copyright (C) 2006-2010 Søren Roug, European Environment Agency\n# \n# This library is free software; you can redistribute it and/or\n# modify it under the terms of the GNU Lesser General Public\n# License as published by the Free Software Foundation; either\n# version 2.1 of the License, or (at your option) any later version.\n# \n# This library is distributed in the hope that it will be useful,\n# but WITHOUT ANY WARRANTY; without even the implied warranty of\n# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU\n# Lesser General Public License for more details.\n# \n# You should have received a copy of the GNU Lesser General Public\n# License along with this library; if not, write to the Free Software\n# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA\n#\n# Contributor(s):\n#\n#import pdb\n#pdb.set_trace()\n\nimport sys, os.path\nsys.path.append(os.path.dirname(__file__))\nfrom xml.sax import handler\nfrom xml.sax.saxutils import escape, quoteattr\nfrom xml.dom import Node\n\nfrom opendocument import load\n\nfrom odf.namespaces import ANIMNS, CHARTNS, CONFIGNS, DCNS, DR3DNS, DRAWNS, FONS, \\\n  FORMNS, MATHNS, METANS, NUMBERNS, OFFICENS, PRESENTATIONNS, SCRIPTNS, \\\n  SMILNS, STYLENS, SVGNS, TABLENS, TEXTNS, XLINKNS\n\n# Handling of styles\n#\n# First there are font face declarations. These set up a font style that will be\n# referenced from a text-property. The declaration describes the font making\n# it possible for the application to find a similar font should the system not\n# have that particular one. The StyleToCSS stores these attributes to be used\n# for the CSS2 font declaration.\n#\n# Then there are default-styles. These set defaults for various style types:\n#  "text", "paragraph", "section", "ruby", "table", "table-column", "table-row",\n#  "table-cell", "graphic", "presentation", "drawing-page", "chart".\n# Since CSS2 can\'t refer to another style, ODF2XHTML add these to all\n# styles unless overridden.\n#\n# The real styles are declared in the <style:style> element. They have a\n# family referring to the default-styles, and may have a parent style.\n# \n# Styles have scope. The same name can be used for both paragraph and\n# character etc. styles Since CSS2 has no scope we use a prefix. (Not elegant)\n# In ODF a style can have a parent, these parents can be chained.\n\nclass StyleToCSS:\n    """ The purpose of the StyleToCSS class is to contain the rules to convert\n        ODF styles to CSS2. Since it needs the generic fonts, it would probably\n        make sense to also contain the Styles in a dict as well..\n    """\n\n    def __init__(self):\n        # Font declarations\n        self.fontdict = {}\n\n        # Fill-images from presentations for backgrounds\n        self.fillimages = {}\n\n        self.ruleconversions = {\n            (DRAWNS,u\'fill-image-name\'): self.c_drawfillimage,\n            (FONS,u"background-color"): self.c_fo,\n            (FONS,u"border"): self.c_fo,\n            (FONS,u"border-bottom"): self.c_fo,\n            (FONS,u"border-left"): self.c_fo,\n            (FONS,u"border-right"): self.c_fo,\n            (FONS,u"border-top"): self.c_fo,\n            (FONS,u"color"): self.c_fo,\n            (FONS,u"font-family"): self.c_fo,\n            (FONS,u"font-size"): self.c_fo,\n            (FONS,u"font-style"): self.c_fo,\n            (FONS,u"font-variant"): self.c_fo,\n            (FONS,u"font-weight"): self.c_fo,\n            (FONS,u"line-height"): self.c_fo,\n            (FONS,u"margin"): self.c_fo,\n            (FONS,u"margin-bottom"): self.c_fo,\n            (FONS,u"margin-left"): self.c_fo,\n            (FONS,u"margin-right"): self.c_fo,\n            (FONS,u"margin-top"): self.c_fo,\n            (FONS,u"min-height"): self.c_fo,\n            (FONS,u"padding"): self.c_fo,\n            (FONS,u"padding-bottom"): self.c_fo,\n            (FONS,u"padding-left"): self.c_fo,\n            (FONS,u"padding-right"): self.c_fo,\n            (FONS,u"padding-top"): self.c_fo,\n            (FONS,u"page-width"): self.c_page_width,\n            (FONS,u"page-height"): self.c_page_height,\n            (FONS,u"text-align"): self.c_text_align,\n            (FONS,u"text-indent") :self.c_fo,\n            (TABLENS,u\'border-model\') :self.c_border_model,\n            (STYLENS,u\'column-width\') : self.c_width,\n            (STYLENS,u"font-name"): self.c_fn,\n            (STYLENS,u\'horizontal-pos\'): self.c_hp,\n            (STYLENS,u\'text-position\'): self.c_text_position,\n            (STYLENS,u\'text-line-through-style\'): self.c_text_line_through_style,\n            (STYLENS,u\'text-underline-style\'): self.c_text_underline_style,\n            (STYLENS,u\'width\') : self.c_width,\n            # FIXME Should do style:vertical-pos here\n        }\n\n    def save_font(self, name, family, generic):\n        """ It is possible that the HTML browser doesn\'t know how to\n            show a particular font. Fortunately ODF provides generic fallbacks.\n            Unfortunately they are not the same as CSS2.\n            CSS2: serif, sans-serif, cursive, fantasy, monospace\n            ODF: roman, swiss, modern, decorative, script, system\n            This method put the font and fallback into a dictionary\n        """\n        htmlgeneric = "sans-serif"\n        if   generic == "roman": htmlgeneric = "serif"\n        elif generic == "swiss": htmlgeneric = "sans-serif"\n        elif generic == "modern": htmlgeneric = "monospace"\n        elif generic == "decorative": htmlgeneric = "sans-serif"\n        elif generic == "script": htmlgeneric = "monospace"\n        elif generic == "system": htmlgeneric = "serif"\n        self.fontdict[name] = (family, htmlgeneric)\n\n    def c_drawfillimage(self, ruleset, sdict, rule, val):\n        """ Fill a figure with an image. Since CSS doesn\'t let you resize images\n            this should really be implemented as an absolutely position <img>\n            with a width and a height\n        """\n        sdict[\'background-image\'] = "url(\'%s\')" % self.fillimages[val]\n\n    def c_fo(self, ruleset, sdict, rule, val):\n        """ XSL formatting attributes """\n        selector = rule[1]\n        sdict[selector] = val\n\n    def c_border_model(self, ruleset, sdict, rule, val):\n        """ Convert to CSS2 border model """\n        if val == \'collapsing\':\n            sdict[\'border-collapse\'] =\'collapse\'\n        else:\n            sdict[\'border-collapse\'] =\'separate\'\n\n    def c_width(self, ruleset, sdict, rule, val):\n        """ Set width of box """\n        sdict[\'width\'] = val\n\n    def c_text_align(self, ruleset, sdict, rule, align):\n        """ Text align """\n        if align == "start": align = "left"\n        if align == "end": align = "right"\n        sdict[\'text-align\'] = align\n\n    def c_fn(self, ruleset, sdict, rule, fontstyle):\n        """ Generate the CSS font family\n            A generic font can be found in two ways. In a <style:font-face>\n            element or as a font-family-generic attribute in text-properties.\n        """\n        generic = ruleset.get((STYLENS,\'font-family-generic\') )\n        if generic is not None:\n            self.save_font(fontstyle, fontstyle, generic)\n        family, htmlgeneric = self.fontdict.get(fontstyle, (fontstyle, \'serif\'))\n        sdict[\'font-family\'] = \'%s, %s\'  % (family, htmlgeneric)\n\n    def c_text_position(self, ruleset, sdict, rule, tp):\n        """ Text position. This is used e.g. to make superscript and subscript\n            This attribute can have one or two values.\n\n            The first value must be present and specifies the vertical\n            text position as a percentage that relates to the current font\n            height or it takes one of the values sub or super. Negative\n            percentages or the sub value place the text below the\n            baseline. Positive percentages or the super value place\n            the text above the baseline. If sub or super is specified,\n            the application can choose an appropriate text position.\n\n            The second value is optional and specifies the font height\n            as a percentage that relates to the current font-height. If\n            this value is not specified, an appropriate font height is\n            used. Although this value may change the font height that\n            is displayed, it never changes the current font height that\n            is used for additional calculations.\n        """\n        textpos = tp.split(\' \')\n        if len(textpos) == 2 and textpos[0] != "0%":\n            # Bug in OpenOffice. If vertical-align is 0% - ignore the text size.\n            sdict[\'font-size\'] = textpos[1]\n        if textpos[0] == "super":\n            sdict[\'vertical-align\'] = "33%"\n        elif textpos[0] == "sub":\n            sdict[\'vertical-align\'] = "-33%"\n        else:\n            sdict[\'vertical-align\'] = textpos[0]\n\n    def c_hp(self, ruleset, sdict, rule, hpos):\n        #FIXME: Frames wrap-style defaults to \'parallel\', graphics to \'none\'.\n        # It is properly set in the parent-styles, but the program doesn\'t\n        # collect the information.\n        wrap = ruleset.get((STYLENS,\'wrap\'),\'parallel\')\n        # Can have: from-left, left, center, right, from-inside, inside, outside\n        if hpos == "center":\n            sdict[\'margin-left\'] = "auto"\n            sdict[\'margin-right\'] = "auto"\n#       else:\n#           # force it to be *something* then delete it\n#           sdict[\'margin-left\'] = sdict[\'margin-right\'] = \'\'\n#           del sdict[\'margin-left\'], sdict[\'margin-right\']\n\n        if hpos in ("right","outside"):\n            if wrap in ( "left", "parallel","dynamic"):\n                sdict[\'float\'] = "right"\n            elif wrap == "run-through":\n                sdict[\'position\'] = "absolute" # Simulate run-through\n                sdict[\'top\'] = "0"\n                sdict[\'right\'] = "0";\n            else: # No wrapping\n                sdict[\'margin-left\'] = "auto"\n                sdict[\'margin-right\'] = "0cm"\n        elif hpos in ("left", "inside"):\n            if wrap in ( "right", "parallel","dynamic"):\n                sdict[\'float\'] = "left"\n            elif wrap == "run-through":\n                sdict[\'position\'] = "absolute" # Simulate run-through\n                sdict[\'top\'] = "0"\n                sdict[\'left\'] = "0"\n            else: # No wrapping\n                sdict[\'margin-left\'] = "0cm"\n                sdict[\'margin-right\'] = "auto"\n        elif hpos in ("from-left", "from-inside"):\n            if wrap in ( "right", "parallel"):\n                sdict[\'float\'] = "left"\n            else:\n                sdict[\'position\'] = "relative" # No wrapping\n                if (SVGNS,\'x\') in ruleset:\n                    sdict[\'left\'] = ruleset[(SVGNS,\'x\')]\n\n    def c_page_width(self, ruleset, sdict, rule, val):\n        """ Set width of box\n            HTML doesn\'t really have a page-width. It is always 100% of the browser width\n        """\n        sdict[\'width\'] = val\n\n    def c_text_underline_style(self, ruleset, sdict, rule, val):\n        """ Set underline decoration\n            HTML doesn\'t really have a page-width. It is always 100% of the browser width\n        """\n        if val and val != "none":\n            sdict[\'text-decoration\'] = "underline"\n\n    def c_text_line_through_style(self, ruleset, sdict, rule, val):\n        """ Set underline decoration\n            HTML doesn\'t really have a page-width. It is always 100% of the browser width\n        """\n        if val and val != "none":\n            sdict[\'text-decoration\'] = "line-through"\n\n    def c_page_height(self, ruleset, sdict, rule, val):\n        """ Set height of box """\n        sdict[\'height\'] = val\n\n    def convert_styles(self, ruleset):\n        """ Rule is a tuple of (namespace, name). If the namespace is \'\' then\n            it is already CSS2\n        """\n        sdict = {}\n        procedures=[]\n        for rule,val in ruleset.items():\n            if rule[0] == \'\':\n                sdict[rule[1]] = val\n                continue\n            method = self.ruleconversions.get(rule, None )\n            if method:\n                procedures.append([method, ruleset, sdict, rule, val])\n        # this ensures that the procedures for horizontal position\n        # are run last! It is important since Python3 makes the order\n        # of dictionaries unpredictable\n        for p in filter(lambda x: x[0] != self.c_hp, procedures):\n            method, ruleset, sdict, rule, val = p\n            method(ruleset, sdict, rule, val)\n        for p in filter(lambda x: x[0] == self.c_hp, procedures):\n            method, ruleset, sdict, rule, val = p\n            method(ruleset, sdict, rule, val)\n\n        return sdict\n\n\nclass TagStack:\n    def __init__(self):\n        self.stack = []\n\n    def push(self, tag, attrs):\n        self.stack.append( (tag, attrs) )\n\n    def pop(self):\n        item = self.stack.pop()\n        return item\n\n    def stackparent(self):\n        item = self.stack[-1]\n        return item[1]\n\n    def rfindattr(self, attr):\n        """ Find a tag with the given attribute """\n        for tag, attrs in self.stack:\n            if attr in attrs:\n                return attrs[attr]\n        return None\n    def count_tags(self, tag):\n        c = 0\n        for ttag, tattrs in self.stack:\n            if ttag == tag: c = c + 1\n        return c\n\nspecial_styles = {\n   \'S-Emphasis\':\'em\',\n   \'S-Citation\':\'cite\',\n   \'S-Strong_20_Emphasis\':\'strong\',\n   \'S-Variable\':\'var\',\n   \'S-Definition\':\'dfn\',\n   \'S-Teletype\':\'tt\',\n   \'P-Heading_20_1\':\'h1\',\n   \'P-Heading_20_2\':\'h2\',\n   \'P-Heading_20_3\':\'h3\',\n   \'P-Heading_20_4\':\'h4\',\n   \'P-Heading_20_5\':\'h5\',\n   \'P-Heading_20_6\':\'h6\',\n#  \'P-Caption\':\'caption\',\n   \'P-Addressee\':\'address\',\n#  \'P-List_20_Heading\':\'dt\',\n#  \'P-List_20_Contents\':\'dd\',\n   \'P-Preformatted_20_Text\':\'pre\',\n#  \'P-Table_20_Heading\':\'th\',\n#  \'P-Table_20_Contents\':\'td\',\n#  \'P-Text_20_body\':\'p\'\n}\n\n#-----------------------------------------------------------------------------\n#\n# ODFCONTENTHANDLER\n#\n#-----------------------------------------------------------------------------\nclass ODF2XHTML(handler.ContentHandler):\n    """ The ODF2XHTML parses an ODF file and produces XHTML"""\n\n    def __init__(self, generate_css=True, embedable=False):\n        # Tags\n        self.generate_css = generate_css\n        self.elements = {\n        (DCNS, \'title\'): (self.s_processcont, self.e_dc_title),\n        (DCNS, \'language\'): (self.s_processcont, self.e_dc_contentlanguage),\n        (DCNS, \'creator\'): (self.s_processcont, self.e_dc_creator),\n        (DCNS, \'description\'): (self.s_processcont, self.e_dc_metatag),\n        (DCNS, \'date\'): (self.s_processcont, self.e_dc_metatag),\n        (DRAWNS, \'custom-shape\'): (self.s_custom_shape, self.e_custom_shape),\n        (DRAWNS, \'frame\'): (self.s_draw_frame, self.e_draw_frame),\n        (DRAWNS, \'image\'): (self.s_draw_image, None),\n        (DRAWNS, \'fill-image\'): (self.s_draw_fill_image, None),\n        (DRAWNS, "layer-set"):(self.s_ignorexml, None),\n        (DRAWNS, \'object\'): (self.s_draw_object, None),\n        (DRAWNS, \'object-ole\'): (self.s_draw_object_ole, None),\n        (DRAWNS, \'page\'): (self.s_draw_page, self.e_draw_page),\n        (DRAWNS, \'text-box\'): (self.s_draw_textbox, self.e_draw_textbox),\n        (METANS, \'creation-date\'):(self.s_processcont, self.e_dc_metatag),\n        (METANS, \'generator\'):(self.s_processcont, self.e_dc_metatag),\n        (METANS, \'initial-creator\'): (self.s_processcont, self.e_dc_metatag),\n        (METANS, \'keyword\'): (self.s_processcont, self.e_dc_metatag),\n        (NUMBERNS, "boolean-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "currency-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "date-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "number-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "text-style"):(self.s_ignorexml, None),\n        (OFFICENS, "annotation"):(self.s_ignorexml, None),\n        (OFFICENS, "automatic-styles"):(self.s_office_automatic_styles, None),\n        (OFFICENS, "document"):(self.s_office_document_content, self.e_office_document_content),\n        (OFFICENS, "document-content"):(self.s_office_document_content, self.e_office_document_content),\n        (OFFICENS, "forms"):(self.s_ignorexml, None),\n        (OFFICENS, "master-styles"):(self.s_office_master_styles, None),\n        (OFFICENS, "meta"):(self.s_ignorecont, None),\n        (OFFICENS, "presentation"):(self.s_office_presentation, self.e_office_presentation),\n        (OFFICENS, "spreadsheet"):(self.s_office_spreadsheet, self.e_office_spreadsheet),\n        (OFFICENS, "styles"):(self.s_office_styles, None),\n        (OFFICENS, "text"):(self.s_office_text, self.e_office_text),\n        (OFFICENS, "scripts"):(self.s_ignorexml, None),\n        (OFFICENS, "settings"):(self.s_ignorexml, None),\n        (PRESENTATIONNS, "notes"):(self.s_ignorexml, None),\n#       (STYLENS, "default-page-layout"):(self.s_style_default_page_layout, self.e_style_page_layout),\n        (STYLENS, "default-page-layout"):(self.s_ignorexml, None),\n        (STYLENS, "default-style"):(self.s_style_default_style, self.e_style_default_style),\n        (STYLENS, "drawing-page-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "font-face"):(self.s_style_font_face, None),\n#       (STYLENS, "footer"):(self.s_style_footer, self.e_style_footer),\n#       (STYLENS, "footer-style"):(self.s_style_footer_style, None),\n        (STYLENS, "graphic-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "handout-master"):(self.s_ignorexml, None),\n#       (STYLENS, "header"):(self.s_style_header, self.e_style_header),\n#       (STYLENS, "header-footer-properties"):(self.s_style_handle_properties, None),\n#       (STYLENS, "header-style"):(self.s_style_header_style, None),\n        (STYLENS, "master-page"):(self.s_style_master_page, None),\n        (STYLENS, "page-layout-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "page-layout"):(self.s_style_page_layout, self.e_style_page_layout),\n#       (STYLENS, "page-layout"):(self.s_ignorexml, None),\n        (STYLENS, "paragraph-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "style"):(self.s_style_style, self.e_style_style),\n        (STYLENS, "table-cell-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "table-column-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "table-properties"):(self.s_style_handle_properties, None),\n        (STYLENS, "text-properties"):(self.s_style_handle_properties, None),\n        (SVGNS, \'desc\'): (self.s_ignorexml, None),\n        (TABLENS, \'covered-table-cell\'): (self.s_ignorexml, None),\n        (TABLENS, \'table-cell\'): (self.s_table_table_cell, self.e_table_table_cell),\n        (TABLENS, \'table-column\'): (self.s_table_table_column, None),\n        (TABLENS, \'table-row\'): (self.s_table_table_row, self.e_table_table_row),\n        (TABLENS, \'table\'): (self.s_table_table, self.e_table_table),\n        (TEXTNS, \'a\'): (self.s_text_a, self.e_text_a),\n        (TEXTNS, "alphabetical-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "bibliography-configuration"):(self.s_ignorexml, None),\n        (TEXTNS, "bibliography-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, \'bookmark\'): (self.s_text_bookmark, None),\n        (TEXTNS, \'bookmark-start\'): (self.s_text_bookmark, None),\n        (TEXTNS, \'bookmark-ref\'): (self.s_text_bookmark_ref, self.e_text_a),\n        (TEXTNS, \'bookmark-ref-start\'): (self.s_text_bookmark_ref, None),\n        (TEXTNS, \'h\'): (self.s_text_h, self.e_text_h),\n        (TEXTNS, "illustration-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, \'line-break\'):(self.s_text_line_break, None),\n        (TEXTNS, "linenumbering-configuration"):(self.s_ignorexml, None),\n        (TEXTNS, "list"):(self.s_text_list, self.e_text_list),\n        (TEXTNS, "list-item"):(self.s_text_list_item, self.e_text_list_item),\n        (TEXTNS, "list-level-style-bullet"):(self.s_text_list_level_style_bullet, self.e_text_list_level_style_bullet),\n        (TEXTNS, "list-level-style-number"):(self.s_text_list_level_style_number, self.e_text_list_level_style_number),\n        (TEXTNS, "list-style"):(None, None),\n        (TEXTNS, "note"):(self.s_text_note, None),\n        (TEXTNS, "note-body"):(self.s_text_note_body, self.e_text_note_body),\n        (TEXTNS, "note-citation"):(None, self.e_text_note_citation),\n        (TEXTNS, "notes-configuration"):(self.s_ignorexml, None),\n        (TEXTNS, "object-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, \'p\'): (self.s_text_p, self.e_text_p),\n        (TEXTNS, \'s\'): (self.s_text_s, None),\n        (TEXTNS, \'span\'): (self.s_text_span, self.e_text_span),\n        (TEXTNS, \'tab\'): (self.s_text_tab, None),\n        (TEXTNS, "table-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "table-of-content-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "user-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        }\n        if embedable:\n            self.set_embedable()\n        self._resetobject()\n\n    def set_plain(self):\n        """ Tell the parser to not generate CSS """\n        self.generate_css = False\n\n    def set_embedable(self):\n        """ Tells the converter to only output the parts inside the <body>"""\n        self.elements[(OFFICENS, u"text")] = (None,None)\n        self.elements[(OFFICENS, u"spreadsheet")] = (None,None)\n        self.elements[(OFFICENS, u"presentation")] = (None,None)\n        self.elements[(OFFICENS, u"document-content")] = (None,None)\n\n\n    def add_style_file(self, stylefilename, media=None):\n        """ Add a link to an external style file.\n            Also turns of the embedding of styles in the HTML\n        """\n        self.use_internal_css = False\n        self.stylefilename = stylefilename\n        if media:\n            self.metatags.append(\'<link rel="stylesheet" type="text/css" href="%s" media="%s"/>\\n\' % (stylefilename,media))\n        else:\n            self.metatags.append(\'<link rel="stylesheet" type="text/css" href="%s"/>\\n\' % (stylefilename))\n\n    def _resetfootnotes(self):\n        # Footnotes and endnotes\n        self.notedict = {}\n        self.currentnote = 0\n        self.notebody = \'\'\n\n    def _resetobject(self):\n        self.lines = []\n        self._wfunc = self._wlines\n        self.xmlfile = \'\'\n        self.title = \'\'\n        self.language = \'\'\n        self.creator = \'\'\n        self.data = []\n        self.tagstack = TagStack()\n        self.htmlstack = []\n        self.pstack = []\n        self.processelem = True\n        self.processcont = True\n        self.listtypes = {}\n        self.headinglevels = [0, 0,0,0,0,0, 0,0,0,0,0] # level 0 to 10\n        self.use_internal_css = True\n        self.cs = StyleToCSS()\n        self.anchors = {}\n\n        # Style declarations\n        self.stylestack = []\n        self.styledict = {}\n        self.currentstyle = None\n\n        self._resetfootnotes()\n\n        # Tags from meta.xml\n        self.metatags = []\n\n\n    def writeout(self, s):\n        if s != \'\':\n            self._wfunc(s)\n\n    def writedata(self):\n        d = \'\'.join(self.data)\n        if d != \'\':\n            self.writeout(escape(d))\n\n    def opentag(self, tag, attrs={}, block=False):\n        """ Create an open HTML tag """\n        self.htmlstack.append((tag,attrs,block))\n        a = []\n        for key,val in attrs.items():\n            a.append(\'\'\'%s=%s\'\'\' % (key, quoteattr(val)))\n        if len(a) == 0:\n            self.writeout("<%s>" % tag)\n        else:\n            self.writeout("<%s %s>" % (tag, " ".join(a)))\n        if block == True:\n            self.writeout("\\n")\n\n    def closetag(self, tag, block=True):\n        """ Close an open HTML tag """\n        self.htmlstack.pop()\n        self.writeout("</%s>" % tag)\n        if block == True:\n            self.writeout("\\n")\n\n    def emptytag(self, tag, attrs={}):\n        a = []\n        for key,val in attrs.items():\n            a.append(\'\'\'%s=%s\'\'\' % (key, quoteattr(val)))\n        self.writeout("<%s %s/>\\n" % (tag, " ".join(a)))\n\n#--------------------------------------------------\n# Interface to parser\n#--------------------------------------------------\n    def characters(self, data):\n        if self.processelem and self.processcont:\n            self.data.append(data)\n\n    def startElementNS(self, tag, qname, attrs):\n        self.pstack.append( (self.processelem, self.processcont) )\n        if self.processelem:\n            method = self.elements.get(tag, (None, None) )[0]\n            if method:\n                self.handle_starttag(tag, method, attrs)\n            else:\n                self.unknown_starttag(tag,attrs)\n        self.tagstack.push( tag, attrs )\n\n    def endElementNS(self, tag, qname):\n        stag, attrs = self.tagstack.pop()\n        if self.processelem:\n            method = self.elements.get(tag, (None, None) )[1]\n            if method:\n                self.handle_endtag(tag, attrs, method)\n            else:\n                self.unknown_endtag(tag, attrs)\n        self.processelem, self.processcont = self.pstack.pop()\n\n#--------------------------------------------------\n    def handle_starttag(self, tag, method, attrs):\n        method(tag,attrs)\n\n    def handle_endtag(self, tag, attrs, method):\n        method(tag, attrs)\n\n    def unknown_starttag(self, tag, attrs):\n        pass\n\n    def unknown_endtag(self, tag, attrs):\n        pass\n\n    def s_ignorexml(self, tag, attrs):\n        """ Ignore this xml element and all children of it\n            It will automatically stop ignoring\n        """\n        self.processelem = False\n\n    def s_ignorecont(self, tag, attrs):\n        """ Stop processing the text nodes """\n        self.processcont = False\n\n    def s_processcont(self, tag, attrs):\n        """ Start processing the text nodes """\n        self.processcont = True\n\n    def classname(self, attrs):\n        """ Generate a class name from a style name """\n        c = attrs.get((TEXTNS,\'style-name\'),\'\')\n        c = c.replace(".","_")\n        return c\n\n    def get_anchor(self, name):\n        """ Create a unique anchor id for a href name """\n        if name not in self.anchors:\n            self.anchors[name] = "anchor%03d" % (len(self.anchors) + 1)\n        return self.anchors.get(name)\n\n\n#--------------------------------------------------\n\n    def purgedata(self):\n        self.data = []\n\n#-----------------------------------------------------------------------------\n#\n# Handle meta data\n#\n#-----------------------------------------------------------------------------\n    def e_dc_title(self, tag, attrs):\n        """ Get the title from the meta data and create a HTML <title>\n        """\n        self.title = \'\'.join(self.data)\n        #self.metatags.append(\'<title>%s</title>\\n\' % escape(self.title))\n        self.data = []\n\n    def e_dc_metatag(self, tag, attrs):\n        """ Any other meta data is added as a <meta> element\n        """\n        self.metatags.append(\'<meta name="%s" content=%s/>\\n\' % (tag[1], quoteattr(\'\'.join(self.data))))\n        self.data = []\n\n    def e_dc_contentlanguage(self, tag, attrs):\n        """ Set the content language. Identifies the targeted audience\n        """\n        self.language = \'\'.join(self.data)\n        self.metatags.append(\'<meta http-equiv="content-language" content="%s"/>\\n\' % escape(self.language))\n        self.data = []\n\n    def e_dc_creator(self, tag, attrs):\n        """ Set the content creator. Identifies the targeted audience\n        """\n        self.creator = \'\'.join(self.data)\n        self.metatags.append(\'<meta http-equiv="creator" content="%s"/>\\n\' % escape(self.creator))\n        self.data = []\n\n    def s_custom_shape(self, tag, attrs):\n        """ A <draw:custom-shape> is made into a <div> in HTML which is then styled\n        """\n        anchor_type = attrs.get((TEXTNS,\'anchor-type\'),\'notfound\')\n        htmltag = \'div\'\n        name = "G-" + attrs.get( (DRAWNS,\'style-name\'), "")\n        if name == \'G-\':\n            name = "PR-" + attrs.get( (PRESENTATIONNS,\'style-name\'), "")\n        name = name.replace(".","_")\n        if anchor_type == "paragraph":\n            style = \'position:absolute;\'\n        elif anchor_type == \'char\':\n            style = "position:absolute;"\n        elif anchor_type == \'as-char\':\n            htmltag = \'div\'\n            style = \'\'\n        else:\n            style = "position: absolute;"\n        if (SVGNS,"width")in attrs:\n            style = style + "width:" + attrs[(SVGNS,"width")] + ";"\n        if (SVGNS,"height") in attrs:\n            style = style + "height:" +  attrs[(SVGNS,"height")] + ";"\n        if (SVGNS,"x") in attrs:\n            style = style + "left:" +  attrs[(SVGNS,"x")] + ";"\n        if (SVGNS,"y") in attrs:\n            style = style + "top:" +  attrs[(SVGNS,"y")] + ";"\n        if self.generate_css:\n            self.opentag(htmltag, {\'class\': name, \'style\': style})\n        else:\n            self.opentag(htmltag)\n\n    def e_custom_shape(self, tag, attrs):\n        """ End the <draw:frame>\n        """\n        self.closetag(\'div\')\n\n    def s_draw_frame(self, tag, attrs):\n        """ A <draw:frame> is made into a <div> in HTML which is then styled\n        """\n        anchor_type = attrs.get((TEXTNS,\'anchor-type\'),\'notfound\')\n        htmltag = \'div\'\n        name = "G-" + attrs.get( (DRAWNS,\'style-name\'), "")\n        if name == \'G-\':\n            name = "PR-" + attrs.get( (PRESENTATIONNS,\'style-name\'), "")\n        name = name.replace(".","_")\n        if anchor_type == "paragraph":\n            style = \'position:relative;\'\n        elif anchor_type == \'char\':\n            style = "position:relative;"\n        elif anchor_type == \'as-char\':\n            htmltag = \'div\'\n            style = \'\'\n        else:\n            style = "position:absolute;"\n        if (SVGNS,"width") in attrs:\n            style = style + "width:" + attrs[(SVGNS,"width")] + ";"\n        if (SVGNS,"height") in attrs:\n            style = style + "height:" +  attrs[(SVGNS,"height")] + ";"\n        if (SVGNS,"x") in attrs:\n            style = style + "left:" +  attrs[(SVGNS,"x")] + ";"\n        if (SVGNS,"y") in attrs:\n            style = style + "top:" +  attrs[(SVGNS,"y")] + ";"\n        if self.generate_css:\n            self.opentag(htmltag, {\'class\': name, \'style\': style})\n        else:\n            self.opentag(htmltag)\n\n    def e_draw_frame(self, tag, attrs):\n        """ End the <draw:frame>\n        """\n        self.closetag(\'div\')\n\n    def s_draw_fill_image(self, tag, attrs):\n        name = attrs.get( (DRAWNS,\'name\'), "NoName")\n        imghref = attrs[(XLINKNS,"href")]\n        imghref = self.rewritelink(imghref)\n        self.cs.fillimages[name] = imghref\n\n    def rewritelink(self, imghref):\n        """ Intended to be overloaded if you don\'t store your pictures\n            in a Pictures subfolder\n        """\n        return imghref\n\n    def s_draw_image(self, tag, attrs):\n        """ A <draw:image> becomes an <img/> element\n        """\n        parent = self.tagstack.stackparent()\n        anchor_type = parent.get((TEXTNS,\'anchor-type\'))\n        imghref = attrs[(XLINKNS,"href")]\n        imghref = self.rewritelink(imghref)\n        htmlattrs = {\'alt\':"", \'src\':imghref }\n        if self.generate_css:\n            if anchor_type != "char":\n                htmlattrs[\'style\'] = "display: block;"\n        self.emptytag(\'img\', htmlattrs)\n\n    def s_draw_object(self, tag, attrs):\n        """ A <draw:object> is embedded object in the document (e.g. spreadsheet in presentation).\n        """\n        objhref = attrs[(XLINKNS,"href")]\n        # Remove leading "./": from "./Object 1" to "Object 1"\n        # objhref = objhref [2:]\n       \n        # Not using os.path.join since it fails to find the file on Windows.\n        # objcontentpath = \'/\'.join([objhref, \'content.xml\'])\n        ################################\n        # fixed: self.document.childnodes ===>  self.document.childobjects\n        # 2016-02-19 G.K.\n        ################################\n        for c in self.document.childobjects:\n            if c.folder == objhref:\n                self._walknode(c.topnode)\n\n    def s_draw_object_ole(self, tag, attrs):\n        """ A <draw:object-ole> is embedded OLE object in the document (e.g. MS Graph).\n        """\n        class_id = attrs[(DRAWNS,"class-id")]\n        if class_id and class_id.lower() == "00020803-0000-0000-c000-000000000046": ## Microsoft Graph 97 Chart\n            tagattrs = { \'name\':\'object_ole_graph\', \'class\':\'ole-graph\' }\n            self.opentag(\'a\', tagattrs)\n            self.closetag(\'a\', tagattrs)\n\n    def s_draw_page(self, tag, attrs):\n        """ A <draw:page> is a slide in a presentation. We use a <fieldset> element in HTML.\n            Therefore if you convert a ODP file, you get a series of <fieldset>s.\n            Override this for your own purpose.\n        """\n        name = attrs.get( (DRAWNS,\'name\'), "NoName")\n        stylename = attrs.get( (DRAWNS,\'style-name\'), "")\n        stylename = stylename.replace(".","_")\n        masterpage = attrs.get( (DRAWNS,\'master-page-name\'),"")\n        masterpage = masterpage.replace(".","_")\n        if self.generate_css:\n            self.opentag(\'fieldset\', {\'class\':"DP-%s MP-%s" % (stylename, masterpage) })\n        else:\n            self.opentag(\'fieldset\')\n        self.opentag(\'legend\')\n        self.writeout(escape(name))\n        self.closetag(\'legend\')\n\n    def e_draw_page(self, tag, attrs):\n        self.closetag(\'fieldset\')\n\n    def s_draw_textbox(self, tag, attrs):\n        style = \'\'\n        if (FONS,"min-height") in attrs:\n            style = style + "min-height:" +  attrs[(FONS,"min-height")] + ";"\n        self.opentag(\'div\')\n#       self.opentag(\'div\', {\'style\': style})\n\n    def e_draw_textbox(self, tag, attrs):\n        """ End the <draw:text-box>\n        """\n        self.closetag(\'div\')\n\n    def html_body(self, tag, attrs):\n        self.writedata()\n        if self.generate_css and self.use_internal_css:\n            self.opentag(\'style\', {\'type\':"text/css"}, True)\n            self.writeout(\'/*<![CDATA[*/\\n\')\n            self.generate_stylesheet()\n            self.writeout(\'/*]]>*/\\n\')\n            self.closetag(\'style\')\n        self.purgedata()\n        self.closetag(\'head\')\n        self.opentag(\'body\', block=True)\n\n    default_styles = """\nimg { width: 100%; height: 100%; }\n* { padding: 0; margin: 0;  background-color:white; }\nbody { margin: 0 1em; }\nol, ul { padding-left: 2em; }\n"""\n\n    def generate_stylesheet(self):\n        for name in self.stylestack:\n            styles = self.styledict.get(name)\n            # Preload with the family\'s default style\n            if \'__style-family\'in styles and styles[\'__style-family\'] in self.styledict:\n                familystyle = self.styledict[styles[\'__style-family\']].copy()\n                del styles[\'__style-family\']\n                for style, val in styles.items():\n                    familystyle[style] = val\n                styles = familystyle\n            # Resolve the remaining parent styles\n            while \'__parent-style-name\' in styles and styles[\'__parent-style-name\'] in self.styledict:\n                parentstyle = self.styledict[styles[\'__parent-style-name\']].copy()\n                del styles[\'__parent-style-name\']\n                for style, val in styles.items():\n                    parentstyle[style] = val\n                styles = parentstyle\n            self.styledict[name] = styles\n        # Write the styles to HTML\n        self.writeout(self.default_styles)\n        for name in self.stylestack:\n            styles = self.styledict.get(name)\n            css2 = self.cs.convert_styles(styles)\n            self.writeout("%s {\\n" % name)\n            for style, val in css2.items():\n                self.writeout("\\t%s: %s;\\n" % (style, val) )\n            self.writeout("}\\n")\n\n    def generate_footnotes(self):\n        if self.currentnote == 0:\n            return\n        if self.generate_css:\n            self.opentag(\'ol\', {\'style\':\'border-top: 1px solid black\'}, True)\n        else:\n            self.opentag(\'ol\')\n        for key in range(1,self.currentnote+1):\n            note = self.notedict[key]\n#       for key,note in self.notedict.items():\n            self.opentag(\'li\', { \'id\':"footnote-%d" % key })\n#           self.opentag(\'sup\')\n#           self.writeout(escape(note[\'citation\']))\n#           self.closetag(\'sup\', False)\n            self.writeout(note[\'body\'])\n            self.closetag(\'li\')\n        self.closetag(\'ol\')\n\n    def s_office_automatic_styles(self, tag, attrs):\n        if self.xmlfile == \'styles.xml\':\n            self.autoprefix = "A"\n        else:\n            self.autoprefix = ""\n\n    def s_office_document_content(self, tag, attrs):\n        """ First tag in the content.xml file"""\n        self.writeout(\'<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" \')\n        self.writeout(\'"http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">\\n\')\n        self.opentag(\'html\', {\'xmlns\':"http://www.w3.org/1999/xhtml"}, True)\n        self.opentag(\'head\', block=True)\n        self.emptytag(\'meta\', { \'http-equiv\':"Content-Type", \'content\':"text/html;charset=UTF-8"})\n        for metaline in self.metatags:\n            self.writeout(metaline)\n        self.writeout(\'<title>%s</title>\\n\' % escape(self.title))\n\n    def e_office_document_content(self, tag, attrs):\n        """ Last tag """\n        self.closetag(\'html\')\n\n    def s_office_master_styles(self, tag, attrs):\n        """ """\n\n    def s_office_presentation(self, tag, attrs):\n        """ For some odd reason, OpenOffice Impress doesn\'t define a default-style\n            for the \'paragraph\'. We therefore force a standard when we see\n            it is a presentation\n        """\n        self.styledict[\'p\'] = {(FONS,u\'font-size\'): u"24pt" }\n        self.styledict[\'presentation\'] = {(FONS,u\'font-size\'): u"24pt" }\n        self.html_body(tag, attrs)\n\n    def e_office_presentation(self, tag, attrs):\n        self.generate_footnotes()\n        self.closetag(\'body\')\n\n    def s_office_spreadsheet(self, tag, attrs):\n        self.html_body(tag, attrs)\n\n    def e_office_spreadsheet(self, tag, attrs):\n        self.generate_footnotes()\n        self.closetag(\'body\')\n\n    def s_office_styles(self, tag, attrs):\n        self.autoprefix = ""\n\n    def s_office_text(self, tag, attrs):\n        """ OpenDocument text """\n        self.styledict[\'frame\'] = { (STYLENS,\'wrap\'): u\'parallel\'}\n        self.html_body(tag, attrs)\n\n    def e_office_text(self, tag, attrs):\n        self.generate_footnotes()\n        self.closetag(\'body\')\n\n    def s_style_handle_properties(self, tag, attrs):\n        """ Copy all attributes to a struct.\n            We will later convert them to CSS2\n        """\n        for key,attr in attrs.items():\n            self.styledict[self.currentstyle][key] = attr\n\n\n    familymap = {\'frame\':\'frame\', \'paragraph\':\'p\', \'presentation\':\'presentation\',\n        \'text\':\'span\',\'section\':\'div\',\n        \'table\':\'table\',\'table-cell\':\'td\',\'table-column\':\'col\',\n        \'table-row\':\'tr\',\'graphic\':\'graphic\' }\n\n    def s_style_default_style(self, tag, attrs):\n        """ A default style is like a style on an HTML tag\n        """\n        family = attrs[(STYLENS,\'family\')]\n        htmlfamily = self.familymap.get(family,\'unknown\')\n        self.currentstyle = htmlfamily\n#       self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n\n    def e_style_default_style(self, tag, attrs):\n        self.currentstyle = None\n\n    def s_style_font_face(self, tag, attrs):\n        """ It is possible that the HTML browser doesn\'t know how to\n            show a particular font. Luckily ODF provides generic fallbacks\n            Unfortunately they are not the same as CSS2.\n            CSS2: serif, sans-serif, cursive, fantasy, monospace\n            ODF: roman, swiss, modern, decorative, script, system\n        """\n        name = attrs[(STYLENS,"name")]\n        family = attrs[(SVGNS,"font-family")]\n        generic = attrs.get( (STYLENS,\'font-family-generic\'),"" )\n        self.cs.save_font(name, family, generic)\n\n    def s_style_footer(self, tag, attrs):\n        self.opentag(\'div\', { \'id\':"footer" })\n        self.purgedata()\n\n    def e_style_footer(self, tag, attrs):\n        self.writedata()\n        self.closetag(\'div\')\n        self.purgedata()\n\n    def s_style_footer_style(self, tag, attrs):\n        self.currentstyle = "@print #footer"\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n\n    def s_style_header(self, tag, attrs):\n        self.opentag(\'div\', { \'id\':"header" })\n        self.purgedata()\n\n    def e_style_header(self, tag, attrs):\n        self.writedata()\n        self.closetag(\'div\')\n        self.purgedata()\n\n    def s_style_header_style(self, tag, attrs):\n        self.currentstyle = "@print #header"\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n\n    def s_style_default_page_layout(self, tag, attrs):\n        """ Collect the formatting for the default page layout style.\n        """\n        self.currentstyle = "@page"\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n\n    def s_style_page_layout(self, tag, attrs):\n        """ Collect the formatting for the page layout style.\n            This won\'t work in CSS 2.1, as page identifiers are not allowed.\n            It is legal in CSS3, but the rest of the application doesn\'t specify when to use what page layout\n        """\n        name = attrs[(STYLENS,\'name\')]\n        name = name.replace(".","_")\n        self.currentstyle = ".PL-" + name\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n\n    def e_style_page_layout(self, tag, attrs):\n        """ End this style\n        """\n        self.currentstyle = None\n\n    def s_style_master_page(self, tag, attrs):\n        """ Collect the formatting for the page layout style.\n        """\n        name = attrs[(STYLENS,\'name\')]\n        name = name.replace(".","_")\n\n        self.currentstyle = ".MP-" + name\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {(\'\',\'position\'):\'relative\'}\n        # Then load the pagelayout style if we find it\n        pagelayout = attrs.get( (STYLENS,\'page-layout-name\'), None)\n        if pagelayout:\n            pagelayout = ".PL-" + pagelayout\n            if pagelayout in self.styledict:\n                styles = self.styledict[pagelayout]\n                for style, val in styles.items():\n                    self.styledict[self.currentstyle][style] = val\n            else:\n                self.styledict[self.currentstyle][\'__parent-style-name\'] = pagelayout\n        self.s_ignorexml(tag, attrs)\n\n    # Short prefixes for class selectors\n    _familyshort = {\'drawing-page\':\'DP\', \'paragraph\':\'P\', \'presentation\':\'PR\',\n        \'text\':\'S\', \'section\':\'D\',\n         \'table\':\'T\', \'table-cell\':\'TD\', \'table-column\':\'TC\',\n         \'table-row\':\'TR\', \'graphic\':\'G\' }\n\n    def s_style_style(self, tag, attrs):\n        """ Collect the formatting for the style.\n            Styles have scope. The same name can be used for both paragraph and\n            character styles Since CSS has no scope we use a prefix. (Not elegant)\n            In ODF a style can have a parent, these parents can be chained.\n            We may not have encountered the parent yet, but if we have, we resolve it.\n        """\n        name = attrs[(STYLENS,\'name\')]\n        name = name.replace(".","_")\n        family = attrs[(STYLENS,\'family\')]\n        htmlfamily = self.familymap.get(family,\'unknown\')\n        sfamily = self._familyshort.get(family,\'X\')\n        name = "%s%s-%s" % (self.autoprefix, sfamily, name)\n        parent = attrs.get( (STYLENS,\'parent-style-name\') )\n        self.currentstyle = special_styles.get(name,"."+name)\n        self.stylestack.append(self.currentstyle)\n        if self.currentstyle not in self.styledict:\n            self.styledict[self.currentstyle] = {}\n\n        self.styledict[self.currentstyle][\'__style-family\'] = htmlfamily\n\n        # Then load the parent style if we find it\n        if parent:\n            parent = "%s-%s" % (sfamily, parent)\n            parent = special_styles.get(parent, "."+parent)\n            if parent in self.styledict:\n                styles = self.styledict[parent]\n                for style, val in styles.items():\n                    self.styledict[self.currentstyle][style] = val\n            else:\n                self.styledict[self.currentstyle][\'__parent-style-name\'] = parent\n\n    def e_style_style(self, tag, attrs):\n        """ End this style\n        """\n        self.currentstyle = None\n\n    def s_table_table(self, tag, attrs):\n        """ Start a table\n        """\n        c = attrs.get( (TABLENS,\'style-name\'), None)\n        if c and self.generate_css:\n            c = c.replace(".","_")\n            self.opentag(\'table\',{ \'class\': "T-%s" % c })\n        else:\n            self.opentag(\'table\')\n        self.purgedata()\n\n    def e_table_table(self, tag, attrs):\n        """ End a table\n        """\n        self.writedata()\n        self.closetag(\'table\')\n        self.purgedata()\n\n    def s_table_table_cell(self, tag, attrs):\n        """ Start a table cell """\n        #FIXME: number-columns-repeated § 8.1.3\n        #repeated = int(attrs.get( (TABLENS,\'number-columns-repeated\'), 1))\n        htmlattrs = {}\n        rowspan = attrs.get( (TABLENS,\'number-rows-spanned\') )\n        if rowspan:\n            htmlattrs[\'rowspan\'] = rowspan\n        colspan = attrs.get( (TABLENS,\'number-columns-spanned\') )\n        if colspan:\n            htmlattrs[\'colspan\'] = colspan\n\n        c = attrs.get( (TABLENS,\'style-name\') )\n        if c:\n            htmlattrs[\'class\'] = \'TD-%s\' % c.replace(".","_")\n        self.opentag(\'td\', htmlattrs)\n        self.purgedata()\n\n    def e_table_table_cell(self, tag, attrs):\n        """ End a table cell """\n        self.writedata()\n        self.closetag(\'td\')\n        self.purgedata()\n\n    def s_table_table_column(self, tag, attrs):\n        """ Start a table column """\n        c = attrs.get( (TABLENS,\'style-name\'), None)\n        repeated = int(attrs.get( (TABLENS,\'number-columns-repeated\'), 1))\n        htmlattrs = {}\n        if c:\n            htmlattrs[\'class\'] = "TC-%s" % c.replace(".","_")\n        for x in range(repeated):\n            self.emptytag(\'col\', htmlattrs)\n        self.purgedata()\n\n    def s_table_table_row(self, tag, attrs):\n        """ Start a table row """\n        #FIXME: table:number-rows-repeated\n        c = attrs.get( (TABLENS,\'style-name\'), None)\n        htmlattrs = {}\n        if c:\n            htmlattrs[\'class\'] = "TR-%s" % c.replace(".","_")\n        self.opentag(\'tr\', htmlattrs)\n        self.purgedata()\n\n    def e_table_table_row(self, tag, attrs):\n        """ End a table row """\n        self.writedata()\n        self.closetag(\'tr\')\n        self.purgedata()\n\n    def s_text_a(self, tag, attrs):\n        """ Anchors start """\n        self.writedata()\n        href = attrs[(XLINKNS,"href")].split("|")[0]\n        if href[0] == "#":\n            href = "#" + self.get_anchor(href[1:])\n        self.opentag(\'a\', {\'href\':href})\n        self.purgedata()\n\n    def e_text_a(self, tag, attrs):\n        """ End an anchor or bookmark reference """\n        self.writedata()\n        self.closetag(\'a\', False)\n        self.purgedata()\n\n    def s_text_bookmark(self, tag, attrs):\n        """ Bookmark definition """\n        name = attrs[(TEXTNS,\'name\')]\n        html_id = self.get_anchor(name)\n        self.writedata()\n        self.opentag(\'span\', {\'id\':html_id})\n        self.closetag(\'span\', False)\n        self.purgedata()\n\n    def s_text_bookmark_ref(self, tag, attrs):\n        """ Bookmark reference """\n        name = attrs[(TEXTNS,\'ref-name\')]\n        html_id = "#" + self.get_anchor(name)\n        self.writedata()\n        self.opentag(\'a\', {\'href\':html_id})\n        self.purgedata()\n\n    def s_text_h(self, tag, attrs):\n        """ Headings start """\n        level = int(attrs[(TEXTNS,\'outline-level\')])\n        if level > 6: level = 6 # Heading levels go only to 6 in XHTML\n        if level < 1: level = 1\n        self.headinglevels[level] = self.headinglevels[level] + 1\n        name = self.classname(attrs)\n        for x in range(level + 1,10):\n            self.headinglevels[x] = 0\n        special = special_styles.get("P-"+name)\n        if special or not self.generate_css:\n            self.opentag(\'h%s\' % level)\n        else:\n            self.opentag(\'h%s\' % level, {\'class\':"P-%s" % name })\n        self.purgedata()\n\n    def e_text_h(self, tag, attrs):\n        """ Headings end\n            Side-effect: If there is no title in the metadata, then it is taken\n            from the first heading of any level.\n        """\n        self.writedata()\n        level = int(attrs[(TEXTNS,\'outline-level\')])\n        if level > 6: level = 6 # Heading levels go only to 6 in XHTML\n        if level < 1: level = 1\n        lev = self.headinglevels[1:level+1]\n        outline = \'.\'.join(map(str,lev) )\n        heading = \'\'.join(self.data)\n        if self.title == \'\': self.title = heading\n        anchor = self.get_anchor("%s.%s" % ( outline, heading))\n        self.opentag(\'a\', {\'id\': anchor} )\n        self.closetag(\'a\', False)\n        self.closetag(\'h%s\' % level)\n        self.purgedata()\n\n    def s_text_line_break(self, tag, attrs):\n        """ Force a line break (<br/>) """\n        self.writedata()\n        self.emptytag(\'br\')\n        self.purgedata()\n\n    def s_text_list(self, tag, attrs):\n        """ Start a list (<ul> or <ol>)\n            To know which level we\'re at, we have to count the number\n            of <text:list> elements on the tagstack.\n        """\n        name = attrs.get( (TEXTNS,\'style-name\') )\n        level = self.tagstack.count_tags(tag) + 1\n        if name:\n            name = name.replace(".","_")\n        else:\n            # FIXME: If a list is contained in a table cell or text box,\n            # the list level must return to 1, even though the table or\n            # textbox itself may be nested within another list.\n            name = self.tagstack.rfindattr( (TEXTNS,\'style-name\') )\n        list_class = "%s_%d" % (name, level)\n        if self.generate_css:\n            self.opentag(\'%s\' % self.listtypes.get(list_class,\'ul\'), {\'class\': list_class })\n        else:\n            self.opentag(\'%s\' % self.listtypes.get(list_class,\'ul\'))\n        self.purgedata()\n\n    def e_text_list(self, tag, attrs):\n        """ End a list """\n        self.writedata()\n        name = attrs.get( (TEXTNS,\'style-name\') )\n        level = self.tagstack.count_tags(tag) + 1\n        if name:\n            name = name.replace(".","_")\n        else:\n            # FIXME: If a list is contained in a table cell or text box,\n            # the list level must return to 1, even though the table or\n            # textbox itself may be nested within another list.\n            name = self.tagstack.rfindattr( (TEXTNS,\'style-name\') )\n        list_class = "%s_%d" % (name, level)\n        self.closetag(self.listtypes.get(list_class,\'ul\'))\n        self.purgedata()\n\n    def s_text_list_item(self, tag, attrs):\n        """ Start list item """\n        self.opentag(\'li\')\n        self.purgedata()\n\n    def e_text_list_item(self, tag, attrs):\n        """ End list item """\n        self.writedata()\n        self.closetag(\'li\')\n        self.purgedata()\n\n    def s_text_list_level_style_bullet(self, tag, attrs):\n        """ CSS doesn\'t have the ability to set the glyph\n            to a particular character, so we just go through\n            the available glyphs\n        """\n        name = self.tagstack.rfindattr( (STYLENS,\'name\') )\n        level = attrs[(TEXTNS,\'level\')]\n        self.prevstyle = self.currentstyle\n        list_class = "%s_%s" % (name, level)\n        self.listtypes[list_class] = \'ul\'\n        self.currentstyle = ".%s_%s" % ( name.replace(".","_"), level)\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n\n        level = int(level)\n        listtype = ("square", "disc", "circle")[level % 3]\n        self.styledict[self.currentstyle][(\'\',\'list-style-type\')] = listtype\n\n    def e_text_list_level_style_bullet(self, tag, attrs):\n        self.currentstyle = self.prevstyle\n        del self.prevstyle\n\n    def s_text_list_level_style_number(self, tag, attrs):\n        name = self.tagstack.stackparent()[(STYLENS,\'name\')]\n        level = attrs[(TEXTNS,\'level\')]\n        num_format = attrs.get( (STYLENS,\'name\'),"1")\n        list_class = "%s_%s" % (name, level)\n        self.prevstyle = self.currentstyle\n        self.currentstyle = ".%s_%s" % ( name.replace(".","_"), level)\n        self.listtypes[list_class] = \'ol\'\n        self.stylestack.append(self.currentstyle)\n        self.styledict[self.currentstyle] = {}\n        if   num_format == "1": listtype = "decimal"\n        elif num_format == "I": listtype = "upper-roman"\n        elif num_format == "i": listtype = "lower-roman"\n        elif num_format == "A": listtype = "upper-alpha"\n        elif num_format == "a": listtype = "lower-alpha"\n        else: listtype = "decimal"\n        self.styledict[self.currentstyle][(\'\',\'list-style-type\')] = listtype\n\n    def e_text_list_level_style_number(self, tag, attrs):\n        self.currentstyle = self.prevstyle\n        del self.prevstyle\n\n    def s_text_note(self, tag, attrs):\n        self.writedata()\n        self.purgedata()\n        self.currentnote = self.currentnote + 1\n        self.notedict[self.currentnote] = {}\n        self.notebody = []\n\n    def e_text_note(self, tag, attrs):\n        pass\n\n    def collectnote(self,s):\n        if s != \'\':\n            self.notebody.append(s)\n\n    def s_text_note_body(self, tag, attrs):\n        self._orgwfunc = self._wfunc\n        self._wfunc = self.collectnote\n\n    def e_text_note_body(self, tag, attrs):\n        self._wfunc = self._orgwfunc\n        self.notedict[self.currentnote][\'body\'] = \'\'.join(self.notebody)\n        self.notebody = \'\'\n        del self._orgwfunc\n\n    def e_text_note_citation(self, tag, attrs):\n        mark = \'\'.join(self.data)\n        self.notedict[self.currentnote][\'citation\'] = mark\n        self.opentag(\'a\',{ \'href\': "#footnote-%s" % self.currentnote })\n        self.opentag(\'sup\')\n#        self.writeout( escape(mark) )\n        # Since HTML only knows about endnotes, there is too much risk that the\n        # marker is reused in the source. Therefore we force numeric markers\n        if sys.version_info[0]==3:\n            self.writeout(self.currentnote)\n        else:\n            self.writeout(unicode(self.currentnote))\n        self.closetag(\'sup\')\n        self.closetag(\'a\')\n\n    def s_text_p(self, tag, attrs):\n        """ Paragraph\n        """\n        htmlattrs = {}\n        specialtag = "p"\n        c = attrs.get( (TEXTNS,\'style-name\'), None)\n        if c:\n            c = c.replace(".","_")\n            specialtag = special_styles.get("P-"+c)\n            if specialtag is None:\n                specialtag = \'p\'\n                if self.generate_css:\n                    htmlattrs[\'class\'] = "P-%s" % c\n        self.opentag(specialtag, htmlattrs)\n        self.purgedata()\n\n    def e_text_p(self, tag, attrs):\n        """ End Paragraph\n        """\n        specialtag = "p"\n        c = attrs.get( (TEXTNS,\'style-name\'), None)\n        if c:\n            c = c.replace(".","_")\n            specialtag = special_styles.get("P-"+c)\n            if specialtag is None:\n                specialtag = \'p\'\n        self.writedata()\n        self.closetag(specialtag)\n        self.purgedata()\n\n    def s_text_s(self, tag, attrs):\n        """ Generate a number of spaces. ODF has an element; HTML uses &nbsp;\n            We use &#160; so we can send the output through an XML parser if we desire to\n        """\n        c = attrs.get( (TEXTNS,\'c\'),"1")\n        for x in range(int(c)):\n            self.writeout(\'&#160;\')\n\n    def s_text_span(self, tag, attrs):\n        """ The <text:span> element matches the <span> element in HTML. It is\n            typically used to properties of the text.\n        """\n        self.writedata()\n        c = attrs.get( (TEXTNS,\'style-name\'), None)\n        htmlattrs = {}\n        if c:\n            c = c.replace(".","_")\n            special = special_styles.get("S-"+c)\n            if special is None and self.generate_css:\n                htmlattrs[\'class\'] = "S-%s" % c\n        self.opentag(\'span\', htmlattrs)\n        self.purgedata()\n\n    def e_text_span(self, tag, attrs):\n        """ End the <text:span> """\n        self.writedata()\n        self.closetag(\'span\', False)\n        self.purgedata()\n\n    def s_text_tab(self, tag, attrs):\n        """ Move to the next tabstop. We ignore this in HTML\n        """\n        self.writedata()\n        self.writeout(\' \')\n        self.purgedata()\n\n    def s_text_x_source(self, tag, attrs):\n        """ Various indexes and tables of contents. We ignore those.\n        """\n        self.writedata()\n        self.purgedata()\n        self.s_ignorexml(tag, attrs)\n\n    def e_text_x_source(self, tag, attrs):\n        """ Various indexes and tables of contents. We ignore those.\n        """\n        self.writedata()\n        self.purgedata()\n\n\n#-----------------------------------------------------------------------------\n#\n# Reading the file\n#\n#-----------------------------------------------------------------------------\n\n    def load(self, odffile):\n        """\n        Loads a document into the parser and parses it.\n        The argument can either be a filename or a document in memory.\n        @param odffile if the type is unicode string: name of a file; else\n        it must be an open file type\n        """\n        assert(type(odffile)==type(u"") or \'rb\' in repr(odffile) or \'BufferedReader\' in repr(odffile)  or \'BytesIO\' in repr(odffile))\n\n        self.lines = []\n        self._wfunc = self._wlines\n        self.document = load(odffile)\n        self._walknode(self.document.topnode)\n\n    def _walknode(self, node):\n        if node.nodeType == Node.ELEMENT_NODE:\n            self.startElementNS(node.qname, node.tagName, node.attributes)\n            for c in node.childNodes:\n                self._walknode(c)\n            self.endElementNS(node.qname, node.tagName)\n        if node.nodeType == Node.TEXT_NODE or node.nodeType == Node.CDATA_SECTION_NODE:\n            if sys.version_info[0]==3:\n                self.characters(str(node))\n            else:\n                self.characters(unicode(node))\n\n\n    def odf2xhtml(self, odffile):\n        """\n        Load a file and return the XHTML\n        @param odffile if the type is unicode string: name of a file; else\n        it must be an open file type\n        @return XHTML code as a a unicode string\n        """\n        assert(type(odffile)==type(u"") or \'rb\' in repr(odffile) or \'BufferedReader\' in repr(odffile)  or \'BytesIO\' in repr(odffile))\n\n\n        self.load(odffile)\n\n        result=self.xhtml()\n        assert(type(result)==type(u""))\n        return result\n\n    def _wlines(self,s):\n        if s != \'\': self.lines.append(s)\n\n    def xhtml(self):\n        """ Returns the xhtml\n        """\n        return \'\'.join(self.lines)\n\n    def _writecss(self, s):\n        if s != \'\': self._csslines.append(s)\n\n    def _writenothing(self, s):\n        pass\n\n    def css(self):\n        """ Returns the CSS content """\n        self._csslines = []\n        self._wfunc = self._writecss\n        self.generate_stylesheet()\n        res = \'\'.join(self._csslines)\n        self._wfunc = self._wlines\n        del self._csslines\n        return res\n\n    def save(self, outputfile, addsuffix=False):\n        """ Save the HTML under the filename.\n            If the filename is \'-\' then save to stdout\n            We have the last style filename in self.stylefilename\n        """\n        if outputfile == \'-\':\n            outputfp = sys.stdout\n        else:\n            if addsuffix:\n                outputfile = outputfile + ".html"\n            outputfp = file(outputfile, "w")\n        outputfp.write(self.xhtml().encode(\'us-ascii\',\'xmlcharrefreplace\'))\n        outputfp.close()\n\n\nclass ODF2XHTMLembedded(ODF2XHTML):\n    """ The ODF2XHTML parses an ODF file and produces XHTML"""\n\n    def __init__(self, lines, generate_css=True, embedable=False):\n        self._resetobject()\n        self.lines = lines\n\n        # Tags\n        self.generate_css = generate_css\n        self.elements = {\n#        (DCNS, \'title\'): (self.s_processcont, self.e_dc_title),\n#        (DCNS, \'language\'): (self.s_processcont, self.e_dc_contentlanguage),\n#        (DCNS, \'creator\'): (self.s_processcont, self.e_dc_metatag),\n#        (DCNS, \'description\'): (self.s_processcont, self.e_dc_metatag),\n#        (DCNS, \'date\'): (self.s_processcont, self.e_dc_metatag),\n        (DRAWNS, \'frame\'): (self.s_draw_frame, self.e_draw_frame),\n        (DRAWNS, \'image\'): (self.s_draw_image, None),\n        (DRAWNS, \'fill-image\'): (self.s_draw_fill_image, None),\n        (DRAWNS, "layer-set"):(self.s_ignorexml, None),\n        (DRAWNS, \'page\'): (self.s_draw_page, self.e_draw_page),\n        (DRAWNS, \'object\'): (self.s_draw_object, None),\n        (DRAWNS, \'object-ole\'): (self.s_draw_object_ole, None),\n        (DRAWNS, \'text-box\'): (self.s_draw_textbox, self.e_draw_textbox),\n#        (METANS, \'creation-date\'):(self.s_processcont, self.e_dc_metatag),\n#        (METANS, \'generator\'):(self.s_processcont, self.e_dc_metatag),\n#        (METANS, \'initial-creator\'): (self.s_processcont, self.e_dc_metatag),\n#        (METANS, \'keyword\'): (self.s_processcont, self.e_dc_metatag),\n        (NUMBERNS, "boolean-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "currency-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "date-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "number-style"):(self.s_ignorexml, None),\n        (NUMBERNS, "text-style"):(self.s_ignorexml, None),\n#        (OFFICENS, "automatic-styles"):(self.s_office_automatic_styles, None),\n#        (OFFICENS, "document-content"):(self.s_office_document_content, self.e_office_document_content),\n        (OFFICENS, "forms"):(self.s_ignorexml, None),\n#        (OFFICENS, "master-styles"):(self.s_office_master_styles, None),\n        (OFFICENS, "meta"):(self.s_ignorecont, None),\n#        (OFFICENS, "presentation"):(self.s_office_presentation, self.e_office_presentation),\n#        (OFFICENS, "spreadsheet"):(self.s_office_spreadsheet, self.e_office_spreadsheet),\n#        (OFFICENS, "styles"):(self.s_office_styles, None),\n#        (OFFICENS, "text"):(self.s_office_text, self.e_office_text),\n        (OFFICENS, "scripts"):(self.s_ignorexml, None),\n        (PRESENTATIONNS, "notes"):(self.s_ignorexml, None),\n##       (STYLENS, "default-page-layout"):(self.s_style_default_page_layout, self.e_style_page_layout),\n#        (STYLENS, "default-page-layout"):(self.s_ignorexml, None),\n#        (STYLENS, "default-style"):(self.s_style_default_style, self.e_style_default_style),\n#        (STYLENS, "drawing-page-properties"):(self.s_style_handle_properties, None),\n#        (STYLENS, "font-face"):(self.s_style_font_face, None),\n##       (STYLENS, "footer"):(self.s_style_footer, self.e_style_footer),\n##       (STYLENS, "footer-style"):(self.s_style_footer_style, None),\n#        (STYLENS, "graphic-properties"):(self.s_style_handle_properties, None),\n#        (STYLENS, "handout-master"):(self.s_ignorexml, None),\n##       (STYLENS, "header"):(self.s_style_header, self.e_style_header),\n##       (STYLENS, "header-footer-properties"):(self.s_style_handle_properties, None),\n##       (STYLENS, "header-style"):(self.s_style_header_style, None),\n#        (STYLENS, "master-page"):(self.s_style_master_page, None),\n#        (STYLENS, "page-layout-properties"):(self.s_style_handle_properties, None),\n##       (STYLENS, "page-layout"):(self.s_style_page_layout, self.e_style_page_layout),\n#        (STYLENS, "page-layout"):(self.s_ignorexml, None),\n#        (STYLENS, "paragraph-properties"):(self.s_style_handle_properties, None),\n#        (STYLENS, "style"):(self.s_style_style, self.e_style_style),\n#        (STYLENS, "table-cell-properties"):(self.s_style_handle_properties, None),\n#        (STYLENS, "table-column-properties"):(self.s_style_handle_properties, None),\n#        (STYLENS, "table-properties"):(self.s_style_handle_properties, None),\n#        (STYLENS, "text-properties"):(self.s_style_handle_properties, None),\n        (SVGNS, \'desc\'): (self.s_ignorexml, None),\n        (TABLENS, \'covered-table-cell\'): (self.s_ignorexml, None),\n        (TABLENS, \'table-cell\'): (self.s_table_table_cell, self.e_table_table_cell),\n        (TABLENS, \'table-column\'): (self.s_table_table_column, None),\n        (TABLENS, \'table-row\'): (self.s_table_table_row, self.e_table_table_row),\n        (TABLENS, \'table\'): (self.s_table_table, self.e_table_table),\n        (TEXTNS, \'a\'): (self.s_text_a, self.e_text_a),\n        (TEXTNS, "alphabetical-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "bibliography-configuration"):(self.s_ignorexml, None),\n        (TEXTNS, "bibliography-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, \'h\'): (self.s_text_h, self.e_text_h),\n        (TEXTNS, "illustration-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, \'line-break\'):(self.s_text_line_break, None),\n        (TEXTNS, "linenumbering-configuration"):(self.s_ignorexml, None),\n        (TEXTNS, "list"):(self.s_text_list, self.e_text_list),\n        (TEXTNS, "list-item"):(self.s_text_list_item, self.e_text_list_item),\n        (TEXTNS, "list-level-style-bullet"):(self.s_text_list_level_style_bullet, self.e_text_list_level_style_bullet),\n        (TEXTNS, "list-level-style-number"):(self.s_text_list_level_style_number, self.e_text_list_level_style_number),\n        (TEXTNS, "list-style"):(None, None),\n        (TEXTNS, "note"):(self.s_text_note, None),\n        (TEXTNS, "note-body"):(self.s_text_note_body, self.e_text_note_body),\n        (TEXTNS, "note-citation"):(None, self.e_text_note_citation),\n        (TEXTNS, "notes-configuration"):(self.s_ignorexml, None),\n        (TEXTNS, "object-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, \'p\'): (self.s_text_p, self.e_text_p),\n        (TEXTNS, \'s\'): (self.s_text_s, None),\n        (TEXTNS, \'span\'): (self.s_text_span, self.e_text_span),\n        (TEXTNS, \'tab\'): (self.s_text_tab, None),\n        (TEXTNS, "table-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "table-of-content-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "user-index-source"):(self.s_text_x_source, self.e_text_x_source),\n        (TEXTNS, "page-number"):(None, None),\n        }\n\n',
    },
    'odf.odfmanifest': {
        'is_package': False,
        'source': r'''
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
# 
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
# 
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
# 
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#
from __future__ import print_function
# This script lists the content of the manifest.xml file
import zipfile
from defusedxml.sax import make_parser
from xml.sax import handler
from xml.sax.xmlreader import InputSource
import xml.sax.saxutils
try:
    from cStringIO import StringIO
except ImportError:
    from io import StringIO

MANIFESTNS="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"

#-----------------------------------------------------------------------------
#
# ODFMANIFESTHANDLER
#
#-----------------------------------------------------------------------------

class ODFManifestHandler(handler.ContentHandler):
    """ The ODFManifestHandler parses a manifest file and produces a list of
        content """

    def __init__(self):
        self.manifest = {}

        # Tags
        # FIXME: Also handle encryption data
        self.elements = {
        (MANIFESTNS, 'file-entry'): (self.s_file_entry, self.donothing),
        }

    def handle_starttag(self, tag, method, attrs):
        method(tag,attrs)

    def handle_endtag(self, tag, method):
        method(tag)

    def startElementNS(self, tag, qname, attrs):
        method = self.elements.get(tag, (None, None))[0]
        if method:
            self.handle_starttag(tag, method, attrs)
        else:
            self.unknown_starttag(tag,attrs)

    def endElementNS(self, tag, qname):
        method = self.elements.get(tag, (None, None))[1]
        if method:
            self.handle_endtag(tag, method)
        else:
            self.unknown_endtag(tag)

    def unknown_starttag(self, tag, attrs):
        pass

    def unknown_endtag(self, tag):
        pass

    def donothing(self, tag, attrs=None):
        pass

    def s_file_entry(self, tag, attrs):
        m = attrs.get((MANIFESTNS, 'media-type'),"application/octet-stream")
        p = attrs.get((MANIFESTNS, 'full-path'))
        self.manifest[p] = { 'media-type':m, 'full-path':p }


#-----------------------------------------------------------------------------
#
# Reading the file
#
#-----------------------------------------------------------------------------

def manifestlist(manifestxml):
    odhandler = ODFManifestHandler()
    parser = make_parser()
    parser.setFeature(handler.feature_namespaces, 1)
    parser.setContentHandler(odhandler)
    parser.setErrorHandler(handler.ErrorHandler())

    inpsrc = InputSource()
    if not isinstance(manifestxml, str):
        manifestxml=manifestxml.decode("utf-8")
    inpsrc.setByteStream(StringIO(manifestxml))
    parser.parse(inpsrc)

    return odhandler.manifest

def odfmanifest(odtfile):
    z = zipfile.ZipFile(odtfile)
    manifest = z.read('META-INF/manifest.xml')
    z.close()
    return manifestlist(manifest)

if __name__ == "__main__":
    import sys
    result = odfmanifest(sys.argv[1])
    for file in result.values():
        print ("%-40s %-40s" % (file['media-type'], file['full-path']))


''',
    },
    'odf.office': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import OFFICENS
from odf.element import Element
from odf.draw import StyleRefElement

# Autogenerated
def Annotation(**args):
    return StyleRefElement(qname = (OFFICENS,'annotation'), **args)

def AnnotationEnd(**args):
    return StyleRefElement(qname = (OFFICENS,'annotation-end'), **args)

def AutomaticStyles(**args):
    return Element(qname = (OFFICENS, 'automatic-styles'), **args)

def BinaryData(**args):
    return Element(qname = (OFFICENS,'binary-data'), **args)

def Body(**args):
    return Element(qname = (OFFICENS, 'body'), **args)

def ChangeInfo(**args):
    return Element(qname = (OFFICENS,'change-info'), **args)

def Chart(**args):
    return Element(qname = (OFFICENS,'chart'), **args)

def DdeSource(**args):
    return Element(qname = (OFFICENS,'dde-source'), **args)

def Document(version="1.2", **args):
    return Element(qname = (OFFICENS,'document'), version=version, **args)

def DocumentContent(version="1.2", **args):
    return Element(qname = (OFFICENS, 'document-content'), version=version, **args)

def DocumentMeta(version="1.2", **args):
    return Element(qname = (OFFICENS, 'document-meta'), version=version, **args)

def DocumentSettings(version="1.2", **args):
    return Element(qname = (OFFICENS, 'document-settings'), version=version, **args)

def DocumentStyles(version="1.2", **args):
    return Element(qname = (OFFICENS, 'document-styles'), version=version, **args)

def Drawing(**args):
    return Element(qname = (OFFICENS,'drawing'), **args)

def EventListeners(**args):
    return Element(qname = (OFFICENS,'event-listeners'), **args)

def FontFaceDecls(**args):
    return Element(qname = (OFFICENS, 'font-face-decls'), **args)

def Forms(**args):
    return Element(qname = (OFFICENS,'forms'), **args)

def Image(**args):
    return Element(qname = (OFFICENS,'image'), **args)

def MasterStyles(**args):
    return Element(qname = (OFFICENS, 'master-styles'), **args)

def Meta(**args):
    return Element(qname = (OFFICENS, 'meta'), **args)

def Presentation(**args):
    return Element(qname = (OFFICENS,'presentation'), **args)

def Script(**args):
    return Element(qname = (OFFICENS, 'script'), **args)

def Scripts(**args):
    return Element(qname = (OFFICENS, 'scripts'), **args)

def Settings(**args):
    return Element(qname = (OFFICENS, 'settings'), **args)

def Spreadsheet(**args):
    return Element(qname = (OFFICENS, 'spreadsheet'), **args)

def Styles(**args):
    return Element(qname = (OFFICENS, 'styles'), **args)

def Text(**args):
    return Element(qname = (OFFICENS, 'text'), **args)

# Autogenerated end

""",
    },
    'odf.opendocument': {
        'is_package': False,
        'source': '# -*- coding: utf-8 -*-\n# Copyright (C) 2006-2010 Søren Roug, European Environment Agency\n#\n# This library is free software; you can redistribute it and/or\n# modify it under the terms of the GNU Lesser General Public\n# License as published by the Free Software Foundation; either\n# version 2.1 of the License, or (at your option) any later version.\n#\n# This library is distributed in the hope that it will be useful,\n# but WITHOUT ANY WARRANTY; without even the implied warranty of\n# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU\n# Lesser General Public License for more details.\n#\n# You should have received a copy of the GNU Lesser General Public\n# License along with this library; if not, write to the Free Software\n# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301 USA\n#\n# Contributor(s):\n#\n# Copyright (C) 2014 Georges Khaznadar <georgesk@debian.org>\n#     migration to Python3, JavaDOC comments and automatic\n#      build of documentation\n#\n\n__doc__="""Use OpenDocument to generate your documents."""\n\nimport zipfile, time, uuid, sys, mimetypes, copy, os.path\n\n# to allow Python3 to access modules in the same path\nsys.path.append(os.path.dirname(__file__))\n\n# using BytesIO provides a cleaner interface than StringIO\n# with both Python2 and Python3: the programmer must care to\n# convert strings or unicode to bytes, which is valid for Python 2 and 3.\nfrom io import StringIO, BytesIO\n\nfrom odf.namespaces import *\nimport odf.manifest as manifest\nimport odf.meta as meta\nfrom odf.office import *\nimport odf.element as element\nfrom odf.attrconverters import make_NCName\nfrom xml.sax.xmlreader import InputSource\nfrom odf.odfmanifest import manifestlist\nimport codecs\n\nif sys.version_info[0] == 3:\n    unicode=str # unicode function does not exist\n\n__version__= TOOLSVERSION\n\n_XMLPROLOGUE = u"<?xml version=\'1.0\' encoding=\'UTF-8\'?>\\n"\n\n#####\n# file permission as an integer value.\n# The following syntax would be invalid for Python3:\n# UNIXPERMS = 0100644 << 16L  # -rw-r--r--\n#\n# So it has been precomputed:\n# 2175008768 is the same value as 0100644 << 16L == -rw-r--r--\n####\nUNIXPERMS = 2175008768\n\nIS_FILENAME = 0\nIS_IMAGE = 1\n# We need at least Python 2.2\nassert sys.version_info[0]>=2 and sys.version_info[1] >= 2\n\n#sys.setrecursionlimit(100)\n#The recursion limit is set conservative so mistakes like\n# s=content() s.addElement(s) won\'t eat up too much processor time.\n\n###############\n# mime-types => file extensions\n###############\nodmimetypes = {\n u\'application/vnd.oasis.opendocument.text\':                  u\'.odt\',\n u\'application/vnd.oasis.opendocument.text-template\':         u\'.ott\',\n u\'application/vnd.oasis.opendocument.graphics\':              u\'.odg\',\n u\'application/vnd.oasis.opendocument.graphics-template\':     u\'.otg\',\n u\'application/vnd.oasis.opendocument.presentation\':          u\'.odp\',\n u\'application/vnd.oasis.opendocument.presentation-template\': u\'.otp\',\n u\'application/vnd.oasis.opendocument.spreadsheet\':           u\'.ods\',\n u\'application/vnd.oasis.opendocument.spreadsheet-template\':  u\'.ots\',\n u\'application/vnd.oasis.opendocument.chart\':                 u\'.odc\',\n u\'application/vnd.oasis.opendocument.chart-template\':        u\'.otc\',\n u\'application/vnd.oasis.opendocument.image\':                 u\'.odi\',\n u\'application/vnd.oasis.opendocument.image-template\':        u\'.oti\',\n u\'application/vnd.oasis.opendocument.formula\':               u\'.odf\',\n u\'application/vnd.oasis.opendocument.formula-template\':      u\'.otf\',\n u\'application/vnd.oasis.opendocument.text-master\':           u\'.odm\',\n u\'application/vnd.oasis.opendocument.text-web\':              u\'.oth\',\n}\n\nclass OpaqueObject:\n    """\n    just a record to bear a filename, a mediatype and a bytes content\n    """\n    def __init__(self, filename, mediatype, content=None):\n        """\n        the constructor\n        @param filename a unicode string\n        @param mediatype a unicode string\n        @param content a byte string or None\n        """\n        assert(type(filename)==type(u""))\n        assert(type(mediatype)==type(u""))\n        assert(type(content)==type(b"") or content == None)\n\n        self.mediatype = mediatype\n        self.filename = filename\n        self.content = content\n\nclass OpenDocument:\n    """\n    A class to hold the content of an OpenDocument document\n    Use the xml method to write the XML\n    source to the screen or to a file.\n    Example of use: d = OpenDocument(mimetype); fd.write(d.xml())\n    """\n    thumbnail = None\n\n    def __init__(self, mimetype, add_generator=True):\n        """\n        the constructor\n        @param mimetype a unicode string\n        @param add_generator a boolean\n        """\n        assert(type(mimetype)==type(u""))\n        assert(isinstance(add_generator,True.__class__))\n\n        self.mimetype = mimetype\n        self.childobjects = []\n        self._extra = []\n        self.folder = u"" # Always empty for toplevel documents\n        self.topnode = Document(mimetype=self.mimetype)\n        self.topnode.ownerDocument = self\n\n        self.clear_caches()\n\n        self.Pictures = {}\n        self.meta = Meta()\n        self.topnode.addElement(self.meta)\n        if add_generator:\n            self.meta.addElement(meta.Generator(text=TOOLSVERSION))\n        self.scripts = Scripts()\n        self.topnode.addElement(self.scripts)\n        self.fontfacedecls = FontFaceDecls()\n        self.topnode.addElement(self.fontfacedecls)\n        self.settings = Settings()\n        self.topnode.addElement(self.settings)\n        self.styles = Styles()\n        self.topnode.addElement(self.styles)\n        self.automaticstyles = AutomaticStyles()\n        self.topnode.addElement(self.automaticstyles)\n        self.masterstyles = MasterStyles()\n        self.topnode.addElement(self.masterstyles)\n        self.body = Body()\n        self.topnode.addElement(self.body)\n\n    def rebuild_caches(self, node=None):\n        if node is None: node = self.topnode\n        self.build_caches(node)\n        for e in node.childNodes:\n            if e.nodeType == element.Node.ELEMENT_NODE:\n                self.rebuild_caches(e)\n\n    def clear_caches(self):\n        """\n        Clears internal caches\n        """\n        self.element_dict = {}\n        self._styles_dict = {}\n        self._styles_ooo_fix = {}\n\n    def build_caches(self, elt):\n        """\n        Builds internal caches; called from element.py\n        @param elt an element.Element instance\n        """\n        # assert(isinstance(elt, element.Element))\n        # why do I need this more intricated assertion?\n        # with Python3, the type of elt pops out as odf.element.Element\n        # in one test ???\n        import odf.element\n        assert(isinstance(elt, element.Element) or isinstance(elt, odf.element.Element) )\n\n        if elt.qname not in self.element_dict:\n            self.element_dict[elt.qname] = []\n        self.element_dict[elt.qname].append(elt)\n        if elt.qname == (STYLENS, u\'style\'):\n            self.__register_stylename(elt) # Add to style dictionary\n        styleref = elt.getAttrNS(TEXTNS,u\'style-name\')\n        if styleref is not None and styleref in self._styles_ooo_fix:\n            elt.setAttrNS(TEXTNS,u\'style-name\', self._styles_ooo_fix[styleref])\n\n    def remove_from_caches(self, elt):\n        """\n        Updates internal caches when an element has been removed\n        @param elt an element.Element instance\n        """\n        # See remark in build_caches about the following assertion\n        import odf.element\n        assert(isinstance(elt, element.Element) or isinstance(elt, odf.element.Element))\n\n        self.element_dict[elt.qname].remove(elt)\n        for e in elt.childNodes:\n            if e.nodeType == element.Node.ELEMENT_NODE:\n                self.remove_from_caches(e)\n\n        if elt.qname == (STYLENS, u\'style\'):\n            del self._styles_dict[elt.getAttrNS(STYLENS, u\'name\')]\n\n    def __register_stylename(self, elt):\n        \'\'\'\n        Register a style. But there are three style dictionaries:\n        office:styles, office:automatic-styles and office:master-styles\n        Chapter 14.\n        @param elt an element.Element instance\n        \'\'\'\n        assert(isinstance(elt, element.Element))\n\n        name = elt.getAttrNS(STYLENS, u\'name\')\n        if name is None:\n            return\n        if elt.parentNode.qname in ((OFFICENS,u\'styles\'), (OFFICENS,u\'automatic-styles\')):\n            if name in self._styles_dict:\n                newname = u\'M\'+name # Rename style\n                self._styles_ooo_fix[name] = newname\n                # From here on all references to the old name will refer to the new one\n                name = newname\n                elt.setAttrNS(STYLENS, u\'name\', name)\n            self._styles_dict[name] = elt\n\n    def toXml(self, filename=u\'\'):\n        """\n        converts the document to a valid Xml format.\n        @param filename unicode string: the name of a file, defaults to\n        an empty string.\n        @return if filename is not empty, the XML code will be written into it\n        and the method returns None; otherwise the method returns a StringIO\n        containing valid XML.\n        Then a ".getvalue()" should return a unicode string.\n        """\n        assert(type(filename)==type(u""))\n\n        result=None\n        xml=StringIO()\n        if sys.version_info[0]==2:\n            xml.write(_XMLPROLOGUE)\n        else:\n            xml.write(_XMLPROLOGUE)\n        self.body.toXml(0, xml)\n        if not filename:\n            result=xml.getvalue()\n        else:\n            f=codecs.open(filename,\'w\', encoding=\'utf-8\')\n            f.write(xml.getvalue())\n            f.close()\n        return result\n\n    def xml(self):\n        """\n        Generates the full document as an XML "file"\n        @return a bytestream in UTF-8 encoding\n        """\n        self.__replaceGenerator()\n        xml=StringIO()\n        if sys.version_info[0]==2:\n            xml.write(_XMLPROLOGUE)\n        else:\n            xml.write(_XMLPROLOGUE)\n        self.topnode.toXml(0, xml)\n        return xml.getvalue().encode("utf-8")\n\n\n    def contentxml(self):\n        """\n        Generates the content.xml file\n        @return a bytestream in UTF-8 encoding\n        """\n        xml=StringIO()\n        xml.write(_XMLPROLOGUE)\n        x = DocumentContent()\n        x.write_open_tag(0, xml)\n        if self.scripts.hasChildNodes():\n            self.scripts.toXml(1, xml)\n        if self.fontfacedecls.hasChildNodes():\n            self.fontfacedecls.toXml(1, xml)\n        a = AutomaticStyles()\n        stylelist = self._used_auto_styles([self.styles, self.automaticstyles, self.body])\n        if len(stylelist) > 0:\n            a.write_open_tag(1, xml)\n            for s in stylelist:\n                s.toXml(2, xml)\n            a.write_close_tag(1, xml)\n        else:\n            a.toXml(1, xml)\n        self.body.toXml(1, xml)\n        x.write_close_tag(0, xml)\n        return xml.getvalue().encode("utf-8")\n\n    def __manifestxml(self):\n        """\n        Generates the manifest.xml file;\n        The self.manifest isn\'t avaible unless the document is being saved\n        @return a unicode string\n        """\n        xml=StringIO()\n        xml.write(_XMLPROLOGUE)\n        self.manifest.toXml(0,xml)\n        result=xml.getvalue()\n        assert(type(result)==type(u""))\n        return result\n\n    def metaxml(self):\n        """\n        Generates the meta.xml file\n        @return a unicode string\n        """\n        self.__replaceGenerator()\n        x = DocumentMeta()\n        x.addElement(self.meta)\n        xml=StringIO()\n        xml.write(_XMLPROLOGUE)\n        x.toXml(0,xml)\n        result=xml.getvalue()\n        assert(type(result)==type(u""))\n        return result\n\n    def settingsxml(self):\n        """\n        Generates the settings.xml file\n        @return a unicode string\n        """\n        x = DocumentSettings()\n        x.addElement(self.settings)\n        xml=StringIO()\n        if sys.version_info[0]==2:\n            xml.write(_XMLPROLOGUE)\n        else:\n            xml.write(_XMLPROLOGUE)\n        x.toXml(0,xml)\n        result=xml.getvalue()\n        assert(type(result)==type(u""))\n        return result\n\n    def _parseoneelement(self, top, stylenamelist):\n        """\n        Finds references to style objects in master-styles\n        and add the style name to the style list if not already there.\n        Recursive\n        @return the list of style names as unicode strings\n        """\n        for e in top.childNodes:\n            if e.nodeType == element.Node.ELEMENT_NODE:\n                for styleref in (\n                        (CHARTNS,u\'style-name\'),\n                        (DRAWNS,u\'style-name\'),\n                        (DRAWNS,u\'text-style-name\'),\n                        (PRESENTATIONNS,u\'style-name\'),\n                        (STYLENS,u\'data-style-name\'),\n                        (STYLENS,u\'list-style-name\'),\n                        (STYLENS,u\'page-layout-name\'),\n                        (STYLENS,u\'style-name\'),\n                        (TABLENS,u\'default-cell-style-name\'),\n                        (TABLENS,u\'style-name\'),\n                        (TEXTNS,u\'style-name\') ):\n                    if e.getAttrNS(styleref[0],styleref[1]):\n                        stylename = e.getAttrNS(styleref[0],styleref[1])\n                        if stylename not in stylenamelist:\n                            # due to the polymorphism of e.getAttrNS(),\n                            # a unicode type is enforced for elements\n                            stylenamelist.append(unicode(stylename))\n                stylenamelist = self._parseoneelement(e, stylenamelist)\n        return stylenamelist\n\n    def _used_auto_styles(self, segments):\n        """\n        Loop through the masterstyles elements, and find the automatic\n        styles that are used. These will be added to the automatic-styles\n        element in styles.xml\n        @return a list of element.Element instances\n        """\n        stylenamelist = []\n        for top in segments:\n            stylenamelist = self._parseoneelement(top, stylenamelist)\n        stylelist = []\n        for e in self.automaticstyles.childNodes:\n            if isinstance(e, element.Element) and e.getAttrNS(STYLENS,u\'name\') in stylenamelist:\n                stylelist.append(e)\n\n        # check the type of the returned data\n        ok=True\n        for e in stylelist: ok = ok and isinstance(e, element.Element)\n        assert(ok)\n\n        return stylelist\n\n    def stylesxml(self):\n        """\n        Generates the styles.xml file\n        @return valid XML code as a unicode string\n        """\n        xml=StringIO()\n        xml.write(_XMLPROLOGUE)\n        x = DocumentStyles()\n        x.write_open_tag(0, xml)\n        if self.fontfacedecls.hasChildNodes():\n            self.fontfacedecls.toXml(1, xml)\n        self.styles.toXml(1, xml)\n        a = AutomaticStyles()\n        a.write_open_tag(1, xml)\n        for s in self._used_auto_styles([self.masterstyles]):\n            s.toXml(2, xml)\n        a.write_close_tag(1, xml)\n        if self.masterstyles.hasChildNodes():\n            self.masterstyles.toXml(1, xml)\n        x.write_close_tag(0, xml)\n        result = xml.getvalue()\n\n        assert(type(result)==type(u""))\n\n        return result\n\n    def addPicture(self, filename, mediatype=None, content=None):\n        """\n        Add a picture\n        It uses the same convention as OOo, in that it saves the picture in\n        the zipfile in the subdirectory \'Pictures\'\n        If passed a file ptr, mediatype must be set\n        @param filename unicode string: name of a file for Pictures\n        @param mediatype unicode string: name of a media, None by default\n        @param content bytes: content of media, None by default\n        @return a unicode string: the file name of the media, eventually\n        created on the fly\n        """\n        if content is None:\n            if mediatype is None:\n                mediatype, encoding = mimetypes.guess_type(filename)\n            if mediatype is None:\n                mediatype = u\'\'\n                try: ext = filename[filename.rindex(u\'.\'):]\n                except: ext=u\'\'\n            else:\n                ext = mimetypes.guess_extension(mediatype)\n            manifestfn = u"Pictures/%s%s" % (uuid.uuid4().hex.upper(), ext)\n            self.Pictures[manifestfn] = (IS_FILENAME, filename, mediatype)\n            content=b""  # this value is only use by the assert further\n            filename=u"" # this value is only use by the assert further\n        else:\n            manifestfn = filename\n            self.Pictures[manifestfn] = (IS_IMAGE, content, mediatype)\n\n        assert(type(filename)==type(u""))\n        assert(type(content) == type(b""))\n\n        return manifestfn\n\n    def addPictureFromFile(self, filename, mediatype=None):\n        """\n        Add a picture\n        It uses the same convention as OOo, in that it saves the picture in\n        the zipfile in the subdirectory \'Pictures\'.\n        If mediatype is not given, it will be guessed from the filename\n        extension.\n        @param filesname unicode string: name of an image file\n        @param mediatype unicode string: type of media, dfaults to None\n        @return a unicode string, the name of the created file\n        """\n        if mediatype is None:\n            mediatype, encoding = mimetypes.guess_type(filename)\n        if mediatype is None:\n            mediatype = u\'\'\n            try: ext = filename[filename.rindex(u\'.\'):]\n            except ValueError: ext=u\'\'\n        else:\n            ext = mimetypes.guess_extension(mediatype)\n        manifestfn = u"Pictures/%s%s" % (uuid.uuid4().hex.upper(), ext)\n        self.Pictures[manifestfn] = (IS_FILENAME, filename, mediatype)\n\n        assert(type(filename)==type(u""))\n        assert(type(mediatype)==type(u""))\n\n        return manifestfn\n\n    def addPictureFromString(self, content, mediatype):\n        """\n        Add a picture from contents given as a Byte string.\n        It uses the same convention as OOo, in that it saves the picture in\n        the zipfile in the subdirectory \'Pictures\'. The content variable\n        is a string that contains the binary image data. The mediatype\n        indicates the image format.\n        @param content bytes: content of media\n        @param mediatype unicode string: name of a media\n        @return a unicode string, the name of the created file\n        """\n        assert(type(content)==type(b""))\n        assert(type(mediatype)==type(u""))\n\n        ext = mimetypes.guess_extension(mediatype)\n        manifestfn = u"Pictures/%s%s" % (uuid.uuid4().hex.upper(), ext)\n        self.Pictures[manifestfn] = (IS_IMAGE, content, mediatype)\n        return manifestfn\n\n    def addThumbnail(self, filecontent=None):\n        """\n        Add a fixed thumbnail\n        The thumbnail in the library is big, so this is pretty useless.\n        @param filecontent bytes: the content of a file; defaults to None\n        """\n        assert(type(filecontent)==type(b""))\n\n        if filecontent is None:\n            import thumbnail\n            self.thumbnail = thumbnail.thumbnail()\n        else:\n            self.thumbnail = filecontent\n\n    def addObject(self, document, objectname=None):\n        """\n        Adds an object (subdocument). The object must be an OpenDocument class\n        @param document OpenDocument instance\n        @param objectname unicode string: the name of an object to add\n        @return a unicode string: the folder name in the zipfile the object is\n        stored in.\n        """\n        assert(isinstance(document, OpenDocument))\n        assert(type(objectname)==type(u"") or objectname == None)\n\n        self.childobjects.append(document)\n        if objectname is None:\n            document.folder = u"%s/Object %d" % (self.folder, len(self.childobjects))\n        else:\n            document.folder = objectname\n        return u".%s" % document.folder\n\n    def _savePictures(self, anObject, folder):\n        """\n        saves pictures contained in an object\n        @param anObject instance of OpenDocument containing pictures\n        @param folder unicode string: place to save pictures\n        """\n        assert(isinstance(anObject, OpenDocument))\n        assert(type(folder)==type(u""))\n\n        hasPictures = False\n        for arcname, picturerec in anObject.Pictures.items():\n            what_it_is, fileobj, mediatype = picturerec\n            self.manifest.addElement(manifest.FileEntry(fullpath=u"%s%s" % ( folder ,arcname), mediatype=mediatype))\n            hasPictures = True\n            if what_it_is == IS_FILENAME:\n                self._z.write(fileobj, folder + arcname, zipfile.ZIP_STORED)\n            else:\n                zi = zipfile.ZipInfo(str(arcname), self._now)\n                zi.compress_type = zipfile.ZIP_STORED\n                zi.external_attr = UNIXPERMS\n                self._z.writestr(zi, fileobj)\n        # According to section 17.7.3 in ODF 1.1, the pictures folder should not have a manifest entry\n#       if hasPictures:\n#           self.manifest.addElement(manifest.FileEntry(fullpath="%sPictures/" % folder, mediatype=""))\n        # Look in subobjects\n        subobjectnum = 1\n        for subobject in anObject.childobjects:\n            self._savePictures(subobject, u\'%sObject %d/\' % (folder, subobjectnum))\n            subobjectnum += 1\n\n    def __replaceGenerator(self):\n        """\n        Removes a previous \'generator\' stance and declares TOOLSVERSION\n        as the new generator.\n        Section 3.1.1: The application MUST NOT export the original identifier\n        belonging to the application that created the document.\n        """\n        for m in self.meta.childNodes[:]:\n            if m.qname == (METANS, u\'generator\'):\n                self.meta.removeChild(m)\n        self.meta.addElement(meta.Generator(text=TOOLSVERSION))\n\n    def save(self, outputfile, addsuffix=False):\n        """\n        Save the document under the filename.\n        If the filename is \'-\' then save to stdout\n        @param outputfile unicode string: the special name \'-\' is for stdout;\n        as an alternative, it can be an io.ByteIO instance which contains\n        the ZIP content.\n        @param addsuffix boolean: whether to add a suffix or not; defaults to False\n        """\n\n        if outputfile == u\'-\':\n            outputfp = zipfile.ZipFile(sys.stdout,"w")\n        else:\n            if addsuffix:\n                outputfile = outputfile + odmimetypes.get(self.mimetype,u\'.xxx\')\n            outputfp = zipfile.ZipFile(outputfile, "w")\n        self.__zipwrite(outputfp)\n        outputfp.close()\n\n    def write(self, outputfp):\n        """\n        User API to write the ODF file to an open file descriptor\n        Writes the ZIP format\n        @param outputfp open file descriptor\n        """\n        zipoutputfp = zipfile.ZipFile(outputfp,"w")\n        self.__zipwrite(zipoutputfp)\n\n    def __zipwrite(self, outputfp):\n        """\n        Write the document to an open file pointer\n        This is where the real work is done\n        @param outputfp instance of zipfile.ZipFile\n        """\n        assert(isinstance(outputfp, zipfile.ZipFile))\n\n        self._z = outputfp\n        self._now = time.localtime()[:6]\n        self.manifest = manifest.Manifest()\n\n        # Write mimetype\n        zi = zipfile.ZipInfo(\'mimetype\', self._now)\n        zi.compress_type = zipfile.ZIP_STORED\n        zi.external_attr = UNIXPERMS\n        self._z.writestr(zi, self.mimetype.encode("utf-8"))\n\n        self._saveXmlObjects(self,u"")\n\n        # Write pictures\n        self._savePictures(self,u"")\n\n        # Write the thumbnail\n        if self.thumbnail is not None:\n            self.manifest.addElement(manifest.FileEntry(fullpath=u"Thumbnails/", mediatype=u\'\'))\n            self.manifest.addElement(manifest.FileEntry(fullpath=u"Thumbnails/thumbnail.png", mediatype=u\'\'))\n            zi = zipfile.ZipInfo(u"Thumbnails/thumbnail.png", self._now)\n            zi.compress_type = zipfile.ZIP_DEFLATED\n            zi.external_attr = UNIXPERMS\n            self._z.writestr(zi, self.thumbnail)\n\n        # Write any extra files\n        for op in self._extra:\n            if op.filename == u"META-INF/documentsignatures.xml": continue # Don\'t save signatures\n            self.manifest.addElement(manifest.FileEntry(fullpath=op.filename, mediatype=op.mediatype))\n            if sys.version_info[0]==3:\n                zi = zipfile.ZipInfo(op.filename, self._now)\n            else:\n                zi = zipfile.ZipInfo(op.filename.encode(\'utf-8\'), self._now)\n            zi.compress_type = zipfile.ZIP_DEFLATED\n            zi.external_attr = UNIXPERMS\n            if op.content is not None:\n                self._z.writestr(zi, op.content)\n        # Write manifest\n        zi = zipfile.ZipInfo(u"META-INF/manifest.xml", self._now)\n        zi.compress_type = zipfile.ZIP_DEFLATED\n        zi.external_attr = UNIXPERMS\n        self._z.writestr(zi, self.__manifestxml() )\n        del self._z\n        del self._now\n        del self.manifest\n\n\n    def _saveXmlObjects(self, anObject, folder):\n        """\n        save xml objects of an opendocument to some folder\n        @param anObject instance of OpenDocument\n        @param folder unicode string place to save xml objects\n        """\n        assert(isinstance(anObject, OpenDocument))\n        assert(type(folder)==type(u""))\n\n        if self == anObject:\n            self.manifest.addElement(manifest.FileEntry(fullpath=u"/", mediatype=anObject.mimetype))\n        else:\n            self.manifest.addElement(manifest.FileEntry(fullpath=folder, mediatype=anObject.mimetype))\n        # Write styles\n        self.manifest.addElement(manifest.FileEntry(fullpath=u"%sstyles.xml" % folder, mediatype=u"text/xml"))\n        zi = zipfile.ZipInfo(u"%sstyles.xml" % folder, self._now)\n        zi.compress_type = zipfile.ZIP_DEFLATED\n        zi.external_attr = UNIXPERMS\n        self._z.writestr(zi, anObject.stylesxml().encode("utf-8") )\n\n        # Write content\n        self.manifest.addElement(manifest.FileEntry(fullpath=u"%scontent.xml" % folder, mediatype=u"text/xml"))\n        zi = zipfile.ZipInfo(u"%scontent.xml" % folder, self._now)\n        zi.compress_type = zipfile.ZIP_DEFLATED\n        zi.external_attr = UNIXPERMS\n        self._z.writestr(zi, anObject.contentxml() )\n\n        # Write settings\n        if anObject.settings.hasChildNodes():\n            self.manifest.addElement(manifest.FileEntry(fullpath=u"%ssettings.xml" % folder, mediatype=u"text/xml"))\n            zi = zipfile.ZipInfo(u"%ssettings.xml" % folder, self._now)\n            zi.compress_type = zipfile.ZIP_DEFLATED\n            zi.external_attr = UNIXPERMS\n            self._z.writestr(zi, anObject.settingsxml().encode("utf-8") )\n\n        # Write meta\n        if self == anObject:\n            self.manifest.addElement(manifest.FileEntry(fullpath=u"meta.xml", mediatype=u"text/xml"))\n            zi = zipfile.ZipInfo(u"meta.xml", self._now)\n            zi.compress_type = zipfile.ZIP_DEFLATED\n            zi.external_attr = UNIXPERMS\n            self._z.writestr(zi, anObject.metaxml().encode("utf-8") )\n\n        # Write subobjects\n        subobjectnum = 1\n        for subobject in anObject.childobjects:\n            self._saveXmlObjects(subobject, u\'%sObject %d/\' % (folder, subobjectnum))\n            subobjectnum += 1\n\n# Document\'s DOM methods\n    def createElement(self, elt):\n        """\n        Inconvenient interface to create an element, but follows XML-DOM.\n        Does not allow attributes as argument, therefore can\'t check grammar.\n        @param elt element.Element instance\n        @return an element.Element instance whose grammar is not checked\n        """\n        assert(isinstance(elt, element.Element))\n\n        # this old code is ambiguous: is \'element\' the module or is it the\n        # local variable? To disambiguate this, the local variable has been\n        # renamed to \'elt\'\n        #return element(check_grammar=False)\n        return elt(check_grammar=False)\n\n    def createTextNode(self, data):\n        """\n        Method to create a text node\n        @param data unicode string to include in the Text element\n        @return an instance of element.Text\n        """\n        assert(type(data)==type(u""))\n\n        return element.Text(data)\n\n    def createCDATASection(self, data):\n        """\n        Method to create a CDATA section\n        @param data unicode string to include in the CDATA element\n        @return an instance of element.CDATASection\n        """\n        assert(type(data)==type(u""))\n\n        return element.CDATASection(cdata)\n\n    def getMediaType(self):\n        """\n        Returns the media type\n        @result a unicode string\n        """\n        assert (type(self.mimetype)==type(u""))\n\n        return self.mimetype\n\n    def getStyleByName(self, name):\n        """\n        Finds a style object based on the name\n        @param name unicode string the name of style to search\n        @return a syle as an element.Element instance\n        """\n        assert(type(name)==type(u""))\n\n        ncname = make_NCName(name)\n        if self._styles_dict == {}:\n            self.rebuild_caches()\n        result=self._styles_dict.get(ncname, None)\n\n        assert(isinstance(result, element.Element))\n        return result\n\n    def getElementsByType(self, elt):\n        """\n        Gets elements based on the type, which is function from\n        text.py, draw.py etc.\n        @param elt instance of a function which returns an element.Element\n        @return a list of istances of element.Element\n        """\n        import types\n        assert(isinstance (elt, types.FunctionType))\n\n        obj = elt(check_grammar=False)\n        assert (isinstance(obj, element.Element))\n\n        if self.element_dict == {}:\n            self.rebuild_caches()\n\n        # This previous code was ambiguous\n        # was "element" the module name or the local variable?\n        # the local variable is renamed to "elt" to disambiguate the code\n        #return self.element_dict.get(obj.qname, [])\n\n        result=self.element_dict.get(obj.qname, [])\n\n        ok=True\n        for e in result: ok = ok and isinstance(e, element.Element)\n        assert(ok)\n\n        return result\n\n# Convenience functions\ndef OpenDocumentChart():\n    """\n    Creates a chart document\n    @return an OpenDocument instance with chart mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.chart\')\n    doc.chart = Chart()\n    doc.body.addElement(doc.chart)\n    return doc\n\ndef OpenDocumentDrawing():\n    """\n    Creates a drawing document\n    @return an OpenDocument instance with drawing mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.graphics\')\n    doc.drawing = Drawing()\n    doc.body.addElement(doc.drawing)\n    return doc\n\ndef OpenDocumentImage():\n    """\n    Creates an image document\n    @return an OpenDocument instance with image mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.image\')\n    doc.image = Image()\n    doc.body.addElement(doc.image)\n    return doc\n\ndef OpenDocumentPresentation():\n    """\n    Creates a presentation document\n    @return an OpenDocument instance with presentation mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.presentation\')\n    doc.presentation = Presentation()\n    doc.body.addElement(doc.presentation)\n    return doc\n\ndef OpenDocumentSpreadsheet():\n    """\n    Creates a spreadsheet document\n    @return an OpenDocument instance with spreadsheet mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.spreadsheet\')\n    doc.spreadsheet = Spreadsheet()\n    doc.body.addElement(doc.spreadsheet)\n    return doc\n\ndef OpenDocumentText():\n    """\n    Creates a text document\n    @return an OpenDocument instance with text mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.text\')\n    doc.text = Text()\n    doc.body.addElement(doc.text)\n    return doc\n\ndef OpenDocumentTextMaster():\n    """\n    Creates a text master document\n    @return an OpenDocument instance with master mimetype\n    """\n    doc = OpenDocument(u\'application/vnd.oasis.opendocument.text-master\')\n    doc.text = Text()\n    doc.body.addElement(doc.text)\n    return doc\n\ndef __loadxmlparts(z, manifest, doc, objectpath):\n    """\n    Parses a document from its zipfile\n    @param z an instance of zipfile.ZipFile\n    @param manifest Manifest data structured in a dictionary\n    @param doc instance of OpenDocument to feed in\n    @param objectpath unicode string: path to an object\n    """\n    assert(isinstance(z, zipfile.ZipFile))\n    assert(type(manifest)==type(dict()))\n    assert(isinstance(doc, OpenDocument))\n    assert(type(objectpath)==type(u""))\n\n    from odf.load import LoadParser\n    from defusedxml.sax import make_parser\n    from xml.sax import handler\n\n    for xmlfile in (objectpath+u\'settings.xml\', objectpath+u\'meta.xml\', objectpath+u\'content.xml\', objectpath+u\'styles.xml\'):\n        if xmlfile not in manifest:\n            continue\n        ##########################################################\n        # this one is added to debug the bad behavior with Python2\n        # which raises exceptions of type SAXParseException\n        from xml.sax._exceptions import SAXParseException\n        ##########################################################\n        try:\n            xmlpart = z.read(xmlfile).decode("utf-8")\n            doc._parsing = xmlfile\n\n            parser = make_parser()\n            parser.setFeature(handler.feature_namespaces, 1)\n            parser.setFeature(handler.feature_external_ges, 0)\n            parser.setContentHandler(LoadParser(doc))\n            parser.setErrorHandler(handler.ErrorHandler())\n\n            inpsrc = InputSource()\n            #################\n            # There may be a SAXParseException triggered because of\n            # a missing xmlns prefix like meta, config, etc.\n            # So i add such declarations when needed (GK, 2014/10/21).\n            # Is there any option to prevent xmlns checks by SAX?\n            xmlpart=__fixXmlPart(xmlpart)\n\n            inpsrc.setByteStream(BytesIO(xmlpart.encode("utf-8")))\n            parser.parse(inpsrc)\n            del doc._parsing\n        except KeyError as v: pass\n        except SAXParseException:\n            print (u"====== SAX FAILED TO PARSE ==========\\n", xmlpart)\n\ndef __fixXmlPart(xmlpart):\n    """\n    fixes an xml code when it does not contain a set of requested\n    "xmlns:whatever" declarations.\n    added by G.K. on 2014/10/21\n    @param xmlpart unicode string: some XML code\n    @return fixed XML code\n    """\n    result=xmlpart\n    requestedPrefixes = (u\'meta\', u\'config\', u\'dc\', u\'style\',\n                         u\'svg\', u\'fo\',u\'draw\', u\'table\',u\'form\')\n    for prefix in requestedPrefixes:\n        if u\' xmlns:{prefix}\'.format(prefix=prefix) not in xmlpart:\n            ###########################################\n            # fixed a bug triggered by math elements\n            # Notice: math elements are creectly exported to XHTML\n            #         and best viewed with MathJax javascript.\n            # 2016-02-19 G.K.\n            ###########################################\n            try:\n                pos=result.index(u" xmlns:")\n                toInsert=u\' xmlns:{prefix}="urn:oasis:names:tc:opendocument:xmlns:{prefix}:1.0"\'.format(prefix=prefix)\n                result=result[:pos]+toInsert+result[pos:]\n            except:\n                pass\n    return result\n\n\ndef __detectmimetype(zipfd, odffile):\n    """\n    detects the mime-type of an ODF file\n    @param zipfd an open zipfile.ZipFile instance\n    @param odffile this parameter is not used\n    @return a mime-type as a unicode string\n    """\n    assert(isinstance(zipfd, zipfile.ZipFile))\n\n    try:\n        mimetype = zipfd.read(\'mimetype\').decode("utf-8")\n        return mimetype\n    except:\n        pass\n    # Fall-through to next mechanism\n    manifestpart = zipfd.read(\'META-INF/manifest.xml\')\n    manifest =  manifestlist(manifestpart)\n    for mentry,mvalue in manifest.items():\n        if mentry == "/":\n            assert(type(mvalue[\'media-type\'])==type(u""))\n            return mvalue[\'media-type\']\n\n    # Fall-through to last mechanism\n    return u\'application/vnd.oasis.opendocument.text\'\n\ndef load(odffile):\n    """\n    Load an ODF file into memory\n    @param odffile unicode string: name of a file, or as an alternative,\n    an open readable stream\n    @return a reference to the structure (an OpenDocument instance)\n    """\n    z = zipfile.ZipFile(odffile)\n    mimetype = __detectmimetype(z, odffile)\n    doc = OpenDocument(mimetype, add_generator=False)\n\n    # Look in the manifest file to see if which of the four files there are\n    manifestpart = z.read(\'META-INF/manifest.xml\')\n    manifest =  manifestlist(manifestpart)\n    __loadxmlparts(z, manifest, doc, u\'\')\n    for mentry,mvalue in manifest.items():\n        if mentry[:9] == u"Pictures/" and len(mentry) > 9:\n            doc.addPicture(mvalue[\'full-path\'], mvalue[\'media-type\'], z.read(mentry))\n        elif mentry == u"Thumbnails/thumbnail.png":\n            doc.addThumbnail(z.read(mentry))\n        elif mentry in (u\'settings.xml\', u\'meta.xml\', u\'content.xml\', u\'styles.xml\'):\n            pass\n        # Load subobjects into structure\n        elif mentry[:7] == u"Object " and len(mentry) < 11 and mentry[-1] == u"/":\n            subdoc = OpenDocument(mvalue[\'media-type\'], add_generator=False)\n            doc.addObject(subdoc, u"/" + mentry[:-1])\n            __loadxmlparts(z, manifest, subdoc, mentry)\n        elif mentry[:7] == u"Object ":\n            pass # Don\'t load subobjects as opaque objects\n        else:\n            if mvalue[\'full-path\'][-1] == u\'/\':\n                doc._extra.append(OpaqueObject(mvalue[\'full-path\'], mvalue[\'media-type\'], None))\n            else:\n                doc._extra.append(OpaqueObject(mvalue[\'full-path\'], mvalue[\'media-type\'], z.read(mentry)))\n            # Add the SUN junk here to the struct somewhere\n            # It is cached data, so it can be out-of-date\n    z.close()\n    b = doc.getElementsByType(Body)\n    if mimetype[:39] == u\'application/vnd.oasis.opendocument.text\':\n        doc.text = b[0].firstChild\n    elif mimetype[:43] == u\'application/vnd.oasis.opendocument.graphics\':\n        doc.graphics = b[0].firstChild\n    elif mimetype[:47] == u\'application/vnd.oasis.opendocument.presentation\':\n        doc.presentation = b[0].firstChild\n    elif mimetype[:46] == u\'application/vnd.oasis.opendocument.spreadsheet\':\n        doc.spreadsheet = b[0].firstChild\n    elif mimetype[:40] == u\'application/vnd.oasis.opendocument.chart\':\n        doc.chart = b[0].firstChild\n    elif mimetype[:40] == u\'application/vnd.oasis.opendocument.image\':\n        doc.image = b[0].firstChild\n    elif mimetype[:42] == u\'application/vnd.oasis.opendocument.formula\':\n        doc.formula = b[0].firstChild\n\n    return doc\n\n# vim: set expandtab sw=4 :\n',
    },
    'odf.presentation': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import PRESENTATIONNS
from odf.element import Element

# ODF 1.0 section 9.6 and 9.7
# Autogenerated
def AnimationGroup(**args):
    return Element(qname = (PRESENTATIONNS,'animation-group'), **args)

def Animations(**args):
    return Element(qname = (PRESENTATIONNS,'animations'), **args)

def DateTime(**args):
    return Element(qname = (PRESENTATIONNS,'date-time'), **args)

def DateTimeDecl(**args):
    return Element(qname = (PRESENTATIONNS,'date-time-decl'), **args)

def Dim(**args):
    return Element(qname = (PRESENTATIONNS,'dim'), **args)

def EventListener(**args):
    return Element(qname = (PRESENTATIONNS,'event-listener'), **args)

def Footer(**args):
    return Element(qname = (PRESENTATIONNS,'footer'), **args)

def FooterDecl(**args):
    return Element(qname = (PRESENTATIONNS,'footer-decl'), **args)

def Header(**args):
    return Element(qname = (PRESENTATIONNS,'header'), **args)

def HeaderDecl(**args):
    return Element(qname = (PRESENTATIONNS,'header-decl'), **args)

def HideShape(**args):
    return Element(qname = (PRESENTATIONNS,'hide-shape'), **args)

def HideText(**args):
    return Element(qname = (PRESENTATIONNS,'hide-text'), **args)

def Notes(**args):
    return Element(qname = (PRESENTATIONNS,'notes'), **args)

def Placeholder(**args):
    return Element(qname = (PRESENTATIONNS,'placeholder'), **args)

def Play(**args):
    return Element(qname = (PRESENTATIONNS,'play'), **args)

def Settings(**args):
    return Element(qname = (PRESENTATIONNS,'settings'), **args)

def Show(**args):
    return Element(qname = (PRESENTATIONNS,'show'), **args)

def ShowShape(**args):
    return Element(qname = (PRESENTATIONNS,'show-shape'), **args)

def ShowText(**args):
    return Element(qname = (PRESENTATIONNS,'show-text'), **args)

def Sound(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (PRESENTATIONNS,'sound'), **args)


""",
    },
    'odf.script': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import SCRIPTNS
from odf.element import Element

# ODF 1.0 section 12.4.1
# The <script:event-listener> element binds an event to a macro.

# Autogenerated
def EventListener(**args):
    return Element(qname = (SCRIPTNS,'event-listener'), **args)


""",
    },
    'odf.style': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import STYLENS
from odf.element import Element

def StyleElement(**args):
    e = Element(**args)
    if args.get('check_grammar', True) == True:
        if 'displayname' not in args:
            e.setAttrNS(STYLENS,'display-name', args.get('name'))
    return e

# Autogenerated
def BackgroundImage(**args):
    return Element(qname = (STYLENS,'background-image'), **args)

def ChartProperties(**args):
    return Element(qname = (STYLENS,'chart-properties'), **args)

def Column(**args):
    return Element(qname = (STYLENS,'column'), **args)

def ColumnSep(**args):
    return Element(qname = (STYLENS,'column-sep'), **args)

def Columns(**args):
    return Element(qname = (STYLENS,'columns'), **args)

def DefaultPageLayout(**args):
    return Element(qname = (STYLENS,'default-page-layout'), **args)

def DefaultStyle(**args):
    return Element(qname = (STYLENS,'default-style'), **args)

def DrawingPageProperties(**args):
    return Element(qname = (STYLENS,'drawing-page-properties'), **args)

def DropCap(**args):
    return Element(qname = (STYLENS,'drop-cap'), **args)

def FontFace(**args):
    return Element(qname = (STYLENS,'font-face'), **args)

def Footer(**args):
    return Element(qname = (STYLENS,'footer'), **args)

def FooterLeft(**args):
    return Element(qname = (STYLENS,'footer-left'), **args)

def FooterStyle(**args):
    return Element(qname = (STYLENS,'footer-style'), **args)

def FootnoteSep(**args):
    return Element(qname = (STYLENS,'footnote-sep'), **args)

def GraphicProperties(**args):
    return Element(qname = (STYLENS,'graphic-properties'), **args)

def HandoutMaster(**args):
    return Element(qname = (STYLENS,'handout-master'), **args)

def Header(**args):
    return Element(qname = (STYLENS,'header'), **args)

def HeaderFooterProperties(**args):
    return Element(qname = (STYLENS,'header-footer-properties'), **args)

def HeaderLeft(**args):
    return Element(qname = (STYLENS,'header-left'), **args)

def HeaderStyle(**args):
    return Element(qname = (STYLENS,'header-style'), **args)

def ListLevelLabelAlignment(**args):
    return Element(qname = (STYLENS,'list-level-label-alignment'), **args)

def ListLevelProperties(**args):
    return Element(qname = (STYLENS,'list-level-properties'), **args)

def Map(**args):
    return Element(qname = (STYLENS,'map'), **args)

def MasterPage(**args):
    return StyleElement(qname = (STYLENS,'master-page'), **args)

def PageLayout(**args):
    return Element(qname = (STYLENS,'page-layout'), **args)

def PageLayoutProperties(**args):
    return Element(qname = (STYLENS,'page-layout-properties'), **args)

def ParagraphProperties(**args):
    return Element(qname = (STYLENS,'paragraph-properties'), **args)

def PresentationPageLayout(**args):
    return StyleElement(qname = (STYLENS,'presentation-page-layout'), **args)

def RegionCenter(**args):
    return Element(qname = (STYLENS,'region-center'), **args)

def RegionLeft(**args):
    return Element(qname = (STYLENS,'region-left'), **args)

def RegionRight(**args):
    return Element(qname = (STYLENS,'region-right'), **args)

def RubyProperties(**args):
    return Element(qname = (STYLENS,'ruby-properties'), **args)

def SectionProperties(**args):
    return Element(qname = (STYLENS,'section-properties'), **args)

def Style(**args):
    return StyleElement(qname = (STYLENS,'style'), **args)

def TabStop(**args):
    return Element(qname = (STYLENS,'tab-stop'), **args)

def TabStops(**args):
    return Element(qname = (STYLENS,'tab-stops'), **args)

def TableCellProperties(**args):
    return Element(qname = (STYLENS,'table-cell-properties'), **args)

def TableColumnProperties(**args):
    return Element(qname = (STYLENS,'table-column-properties'), **args)

def TableProperties(**args):
    return Element(qname = (STYLENS,'table-properties'), **args)

def TableRowProperties(**args):
    return Element(qname = (STYLENS,'table-row-properties'), **args)

def TextProperties(**args):
    return Element(qname = (STYLENS,'text-properties'), **args)


""",
    },
    'odf.svg': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import SVGNS
from odf.element import Element
from odf.draw import DrawElement

# Autogenerated
def DefinitionSrc(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (SVGNS,'definition-src'), **args)

def Desc(**args):
    return Element(qname = (SVGNS,'desc'), **args)

def FontFaceFormat(**args):
    return Element(qname = (SVGNS,'font-face-format'), **args)

def FontFaceName(**args):
    return Element(qname = (SVGNS,'font-face-name'), **args)

def FontFaceSrc(**args):
    return Element(qname = (SVGNS,'font-face-src'), **args)

def FontFaceUri(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (SVGNS,'font-face-uri'), **args)

def Lineargradient(**args):
    return DrawElement(qname = (SVGNS,'linearGradient'), **args)

def Radialgradient(**args):
    return DrawElement(qname = (SVGNS,'radialGradient'), **args)

def Stop(**args):
    return Element(qname = (SVGNS,'stop'), **args)

def Title(**args):
    return Element(qname = (SVGNS,'title'), **args)

""",
    },
    'odf.table': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import TABLENS
from odf.element import Element


# Autogenerated
def Background(**args):
    return Element(qname = (TABLENS,'background'), **args)

def Body(**args):
    return Element(qname = (TABLENS,'body'), **args)

def CalculationSettings(**args):
    return Element(qname = (TABLENS,'calculation-settings'), **args)

def CellAddress(**args):
    return Element(qname = (TABLENS,'cell-address'), **args)

def CellContentChange(**args):
    return Element(qname = (TABLENS,'cell-content-change'), **args)

def CellContentDeletion(**args):
    return Element(qname = (TABLENS,'cell-content-deletion'), **args)

def CellRangeSource(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (TABLENS,'cell-range-source'), **args)

def ChangeDeletion(**args):
    return Element(qname = (TABLENS,'change-deletion'), **args)

def ChangeTrackTableCell(**args):
    return Element(qname = (TABLENS,'change-track-table-cell'), **args)

def Consolidation(**args):
    return Element(qname = (TABLENS,'consolidation'), **args)

def ContentValidation(**args):
    return Element(qname = (TABLENS,'content-validation'), **args)

def ContentValidations(**args):
    return Element(qname = (TABLENS,'content-validations'), **args)

def CoveredTableCell(**args):
    return Element(qname = (TABLENS,'covered-table-cell'), **args)

def CutOffs(**args):
    return Element(qname = (TABLENS,'cut-offs'), **args)

def DataPilotDisplayInfo(**args):
    return Element(qname = (TABLENS,'data-pilot-display-info'), **args)

def DataPilotField(**args):
    return Element(qname = (TABLENS,'data-pilot-field'), **args)

def DataPilotFieldReference(**args):
    return Element(qname = (TABLENS,'data-pilot-field-reference'), **args)

def DataPilotGroup(**args):
    return Element(qname = (TABLENS,'data-pilot-group'), **args)

def DataPilotGroupMember(**args):
    return Element(qname = (TABLENS,'data-pilot-group-member'), **args)

def DataPilotGroups(**args):
    return Element(qname = (TABLENS,'data-pilot-groups'), **args)

def DataPilotLayoutInfo(**args):
    return Element(qname = (TABLENS,'data-pilot-layout-info'), **args)

def DataPilotLevel(**args):
    return Element(qname = (TABLENS,'data-pilot-level'), **args)

def DataPilotMember(**args):
    return Element(qname = (TABLENS,'data-pilot-member'), **args)

def DataPilotMembers(**args):
    return Element(qname = (TABLENS,'data-pilot-members'), **args)

def DataPilotSortInfo(**args):
    return Element(qname = (TABLENS,'data-pilot-sort-info'), **args)

def DataPilotSubtotal(**args):
    return Element(qname = (TABLENS,'data-pilot-subtotal'), **args)

def DataPilotSubtotals(**args):
    return Element(qname = (TABLENS,'data-pilot-subtotals'), **args)

def DataPilotTable(**args):
    return Element(qname = (TABLENS,'data-pilot-table'), **args)

def DataPilotTables(**args):
    return Element(qname = (TABLENS,'data-pilot-tables'), **args)

def DatabaseRange(**args):
    return Element(qname = (TABLENS,'database-range'), **args)

def DatabaseRanges(**args):
    return Element(qname = (TABLENS,'database-ranges'), **args)

def DatabaseSourceQuery(**args):
    return Element(qname = (TABLENS,'database-source-query'), **args)

def DatabaseSourceSql(**args):
    return Element(qname = (TABLENS,'database-source-sql'), **args)

def DatabaseSourceTable(**args):
    return Element(qname = (TABLENS,'database-source-table'), **args)

def DdeLink(**args):
    return Element(qname = (TABLENS,'dde-link'), **args)

def DdeLinks(**args):
    return Element(qname = (TABLENS,'dde-links'), **args)

def Deletion(**args):
    return Element(qname = (TABLENS,'deletion'), **args)

def Deletions(**args):
    return Element(qname = (TABLENS,'deletions'), **args)

def Dependencies(**args):
    return Element(qname = (TABLENS,'dependencies'), **args)

def Dependency(**args):
    return Element(qname = (TABLENS,'dependency'), **args)

def Desc(**args):
    return Element(qname = (TABLENS,'desc'), **args)

def Detective(**args):
    return Element(qname = (TABLENS,'detective'), **args)

def ErrorMacro(**args):
    return Element(qname = (TABLENS,'error-macro'), **args)

def ErrorMessage(**args):
    return Element(qname = (TABLENS,'error-message'), **args)

def EvenColumns(**args):
    return Element(qname = (TABLENS,'even-columns'), **args)

def EvenRows(**args):
    return Element(qname = (TABLENS,'even-rows'), **args)

def Filter(**args):
    return Element(qname = (TABLENS,'filter'), **args)

def FilterAnd(**args):
    return Element(qname = (TABLENS,'filter-and'), **args)

def FilterCondition(**args):
    return Element(qname = (TABLENS,'filter-condition'), **args)

def FilterOr(**args):
    return Element(qname = (TABLENS,'filter-or'), **args)

def FilterSetItem(**args):
    return Element(qname = (TABLENS,'filter-set-item'), **args)

def FirstColumn(**args):
    return Element(qname = (TABLENS,'first-column'), **args)

def FirstRow(**args):
    return Element(qname = (TABLENS,'first-row'), **args)

def HelpMessage(**args):
    return Element(qname = (TABLENS,'help-message'), **args)

def HighlightedRange(**args):
    return Element(qname = (TABLENS,'highlighted-range'), **args)

def Insertion(**args):
    return Element(qname = (TABLENS,'insertion'), **args)

def InsertionCutOff(**args):
    return Element(qname = (TABLENS,'insertion-cut-off'), **args)

def Iteration(**args):
    return Element(qname = (TABLENS,'iteration'), **args)

def LabelRange(**args):
    return Element(qname = (TABLENS,'label-range'), **args)

def LabelRanges(**args):
    return Element(qname = (TABLENS,'label-ranges'), **args)

def LastColumn(**args):
    return Element(qname = (TABLENS,'last-column'), **args)

def LastRow(**args):
    return Element(qname = (TABLENS,'last-row'), **args)

def Movement(**args):
    return Element(qname = (TABLENS,'movement'), **args)

def MovementCutOff(**args):
    return Element(qname = (TABLENS,'movement-cut-off'), **args)

def NamedExpression(**args):
    return Element(qname = (TABLENS,'named-expression'), **args)

def NamedExpressions(**args):
    return Element(qname = (TABLENS,'named-expressions'), **args)

def NamedRange(**args):
    return Element(qname = (TABLENS,'named-range'), **args)

def NullDate(**args):
    return Element(qname = (TABLENS,'null-date'), **args)

def OddColumns(**args):
    return Element(qname = (TABLENS,'odd-columns'), **args)

def OddRows(**args):
    return Element(qname = (TABLENS,'odd-rows'), **args)

def Operation(**args):
    return Element(qname = (TABLENS,'operation'), **args)

def Previous(**args):
    return Element(qname = (TABLENS,'previous'), **args)

def Scenario(**args):
    return Element(qname = (TABLENS,'scenario'), **args)

def Shapes(**args):
    return Element(qname = (TABLENS,'shapes'), **args)

def Sort(**args):
    return Element(qname = (TABLENS,'sort'), **args)

def SortBy(**args):
    return Element(qname = (TABLENS,'sort-by'), **args)

def SortGroups(**args):
    return Element(qname = (TABLENS,'sort-groups'), **args)

def SourceCellRange(**args):
    return Element(qname = (TABLENS,'source-cell-range'), **args)

def SourceRangeAddress(**args):
    return Element(qname = (TABLENS,'source-range-address'), **args)

def SourceService(**args):
    return Element(qname = (TABLENS,'source-service'), **args)

def SubtotalField(**args):
    return Element(qname = (TABLENS,'subtotal-field'), **args)

def SubtotalRule(**args):
    return Element(qname = (TABLENS,'subtotal-rule'), **args)

def SubtotalRules(**args):
    return Element(qname = (TABLENS,'subtotal-rules'), **args)

def Table(**args):
    return Element(qname = (TABLENS,'table'), **args)

def TableCell(**args):
    return Element(qname = (TABLENS,'table-cell'), **args)

def TableColumn(**args):
    return Element(qname = (TABLENS,'table-column'), **args)

def TableColumnGroup(**args):
    return Element(qname = (TABLENS,'table-column-group'), **args)

def TableColumns(**args):
    return Element(qname = (TABLENS,'table-columns'), **args)

def TableHeaderColumns(**args):
    return Element(qname = (TABLENS,'table-header-columns'), **args)

def TableHeaderRows(**args):
    return Element(qname = (TABLENS,'table-header-rows'), **args)

def TableRow(**args):
    return Element(qname = (TABLENS,'table-row'), **args)

def TableRowGroup(**args):
    return Element(qname = (TABLENS,'table-row-group'), **args)

def TableRows(**args):
    return Element(qname = (TABLENS,'table-rows'), **args)

def TableSource(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (TABLENS,'table-source'), **args)

def TableTemplate(**args):
    return Element(qname = (TABLENS,'table-template'), **args)

def TargetRangeAddress(**args):
    return Element(qname = (TABLENS,'target-range-address'), **args)

def Title(**args):
    return Element(qname = (TABLENS,'title'), **args)

def TrackedChanges(**args):
    return Element(qname = (TABLENS,'tracked-changes'), **args)


""",
    },
    'odf.teletype': {
        'is_package': False,
        'source': r'''
# -*- coding: utf-8 -*-
#
#   Create and extract text from ODF, handling whitespace correctly.
#   Copyright (C) 2008 J. David Eisenberg
#
#   This program is free software; you can redistribute it and/or modify
#   it under the terms of the GNU General Public License as published by
#   the Free Software Foundation; either version 2 of the License, or
#   (at your option) any later version.
#
#   This program is distributed in the hope that it will be useful,
#   but WITHOUT ANY WARRANTY; without even the implied warranty of
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#   GNU General Public License for more details.
#
#   You should have received a copy of the GNU General Public License along
#   with this program; if not, write to the Free Software Foundation, Inc.,
#   51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.


"""
Class for handling whitespace properly in OpenDocument.

While it is possible to use getTextContent() and setTextContent()
to extract or create ODF content, these won't extract or create
the appropriate <text:s>, <text:tab>, or <text:line-break>
elements.  This module takes care of that problem.
"""

from odf.element import Node
import odf.opendocument
from odf.text import S,LineBreak,Tab

class WhitespaceText(object):

    def __init__(self):
        self.textBuffer = []
        self.spaceCount = 0

    def addTextToElement(self, odfElement, s):
        """ Process an input string, inserting
            <text:tab> elements for '\t',
            <text:line-break> elements for '\n', and
            <text:s> elements for runs of more than one blank.
            These will be added to the given element.
        """
        i = 0
        ch = ' '

        # When we encounter a tab or newline, we can immediately
        # dump any accumulated text and then emit the appropriate
        # ODF element.
        #
        # When we encounter a space, we add it to the text buffer,
        # and then collect more spaces.  If there are more spaces
        # after the first one, we dump the text buffer and then
        # then emit the appropriate <text:s> element.

        while i < len(s):
            ch = s[i]
            if ch == '\t':
                self._emitTextBuffer(odfElement)
                odfElement.addElement(Tab())
                i += 1
            elif ch == '\n':
                self._emitTextBuffer(odfElement);
                odfElement.addElement(LineBreak())
                i += 1
            elif ch == ' ':
                self.textBuffer.append(' ')
                i += 1
                self.spaceCount = 0
                while i < len(s) and (s[i] == ' '):
                    self.spaceCount += 1
                    i += 1
                if self.spaceCount > 0:
                    self._emitTextBuffer(odfElement)
                    self._emitSpaces(odfElement)
            else:
                self.textBuffer.append(ch)
                i += 1

        self._emitTextBuffer(odfElement)

    def _emitTextBuffer(self, odfElement):
        """ Creates a Text Node whose contents are the current textBuffer.
            Side effect: clears the text buffer.
        """
        if len(self.textBuffer) > 0:
            odfElement.addText(''.join(self.textBuffer))
        self.textBuffer = []


    def _emitSpaces(self, odfElement):
        """ Creates a <text:s> element for the current spaceCount.
            Side effect: sets spaceCount back to zero
        """
        if self.spaceCount > 0:
            spaceElement = S(c=self.spaceCount)
            odfElement.addElement(spaceElement)
        self.spaceCount = 0

def addTextToElement(odfElement, s):
    wst = WhitespaceText()
    wst.addTextToElement(odfElement, s)

def extractText(odfElement):
    """ Extract text content from an Element, with whitespace represented
        properly. Returns the text, with tabs, spaces, and newlines
        correctly evaluated. This method recursively descends through the
        children of the given element, accumulating text and "unwrapping"
        <text:s>, <text:tab>, and <text:line-break> elements along the way.
    """
    result = [];

    if len(odfElement.childNodes) != 0:
        for child in odfElement.childNodes:
            if child.nodeType == Node.TEXT_NODE:
                result.append(child.data)
            elif child.nodeType == Node.ELEMENT_NODE:
                subElement = child
                tagName = subElement.qname;
                if tagName == (u"urn:oasis:names:tc:opendocument:xmlns:text:1.0", u"line-break"):
                    result.append("\n")
                elif tagName == (u"urn:oasis:names:tc:opendocument:xmlns:text:1.0", u"tab"):
                    result.append("\t")
                elif tagName == (u"urn:oasis:names:tc:opendocument:xmlns:text:1.0", u"s"):
                    c = subElement.getAttribute('c')
                    if c:
                        spaceCount =  int(c)
                    else:
                        spaceCount = 1

                    result.append(" " * spaceCount)
                else:
                    result.append(extractText(subElement))
    return ''.join(result)

''',
    },
    'odf.text': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2013 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#
import re, sys, os.path
sys.path.append(os.path.dirname(__file__))


from odf.namespaces import TEXTNS
from odf.element import Element
from odf.style import StyleElement

# Autogenerated
def A(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (TEXTNS,'a'), **args)

def AlphabeticalIndex(**args):
    return Element(qname = (TEXTNS,'alphabetical-index'), **args)

def AlphabeticalIndexAutoMarkFile(**args):
    args.setdefault('type', 'simple')
    return Element(qname = (TEXTNS,'alphabetical-index-auto-mark-file'), **args)

def AlphabeticalIndexEntryTemplate(**args):
    return Element(qname = (TEXTNS,'alphabetical-index-entry-template'), **args)

def AlphabeticalIndexMark(**args):
    return Element(qname = (TEXTNS,'alphabetical-index-mark'), **args)

def AlphabeticalIndexMarkEnd(**args):
    return Element(qname = (TEXTNS,'alphabetical-index-mark-end'), **args)

def AlphabeticalIndexMarkStart(**args):
    return Element(qname = (TEXTNS,'alphabetical-index-mark-start'), **args)

def AlphabeticalIndexSource(**args):
    return Element(qname = (TEXTNS,'alphabetical-index-source'), **args)

def AuthorInitials(**args):
    return Element(qname = (TEXTNS,'author-initials'), **args)

def AuthorName(**args):
    return Element(qname = (TEXTNS,'author-name'), **args)

def Bibliography(**args):
    return Element(qname = (TEXTNS,'bibliography'), **args)

def BibliographyConfiguration(**args):
    return Element(qname = (TEXTNS,'bibliography-configuration'), **args)

def BibliographyEntryTemplate(**args):
    return Element(qname = (TEXTNS,'bibliography-entry-template'), **args)

def BibliographyMark(**args):
    return Element(qname = (TEXTNS,'bibliography-mark'), **args)

def BibliographySource(**args):
    return Element(qname = (TEXTNS,'bibliography-source'), **args)

def Bookmark(**args):
    return Element(qname = (TEXTNS,'bookmark'), **args)

def BookmarkEnd(**args):
    return Element(qname = (TEXTNS,'bookmark-end'), **args)

def BookmarkRef(**args):
    return Element(qname = (TEXTNS,'bookmark-ref'), **args)

def BookmarkStart(**args):
    return Element(qname = (TEXTNS,'bookmark-start'), **args)

def Change(**args):
    return Element(qname = (TEXTNS,'change'), **args)

def ChangeEnd(**args):
    return Element(qname = (TEXTNS,'change-end'), **args)

def ChangeStart(**args):
    return Element(qname = (TEXTNS,'change-start'), **args)

def ChangedRegion(**args):
    return Element(qname = (TEXTNS,'changed-region'), **args)

def Chapter(**args):
    return Element(qname = (TEXTNS,'chapter'), **args)

def CharacterCount(**args):
    return Element(qname = (TEXTNS,'character-count'), **args)

def ConditionalText(**args):
    return Element(qname = (TEXTNS,'conditional-text'), **args)

def CreationDate(**args):
    return Element(qname = (TEXTNS,'creation-date'), **args)

def CreationTime(**args):
    return Element(qname = (TEXTNS,'creation-time'), **args)

def Creator(**args):
    return Element(qname = (TEXTNS,'creator'), **args)

def DatabaseDisplay(**args):
    return Element(qname = (TEXTNS,'database-display'), **args)

def DatabaseName(**args):
    return Element(qname = (TEXTNS,'database-name'), **args)

def DatabaseNext(**args):
    return Element(qname = (TEXTNS,'database-next'), **args)

def DatabaseRowNumber(**args):
    return Element(qname = (TEXTNS,'database-row-number'), **args)

def DatabaseRowSelect(**args):
    return Element(qname = (TEXTNS,'database-row-select'), **args)

def Date(**args):
    return Element(qname = (TEXTNS,'date'), **args)

def DdeConnection(**args):
    return Element(qname = (TEXTNS,'dde-connection'), **args)

def DdeConnectionDecl(**args):
    return Element(qname = (TEXTNS,'dde-connection-decl'), **args)

def DdeConnectionDecls(**args):
    return Element(qname = (TEXTNS,'dde-connection-decls'), **args)

def Deletion(**args):
    return Element(qname = (TEXTNS,'deletion'), **args)

def Description(**args):
    return Element(qname = (TEXTNS,'description'), **args)

def EditingCycles(**args):
    return Element(qname = (TEXTNS,'editing-cycles'), **args)

def EditingDuration(**args):
    return Element(qname = (TEXTNS,'editing-duration'), **args)

def ExecuteMacro(**args):
    return Element(qname = (TEXTNS,'execute-macro'), **args)

def Expression(**args):
    return Element(qname = (TEXTNS,'expression'), **args)

def FileName(**args):
    return Element(qname = (TEXTNS,'file-name'), **args)

def FormatChange(**args):
    return Element(qname = (TEXTNS,'format-change'), **args)

def H(**args):
    return Element(qname = (TEXTNS, 'h'), **args)

def HiddenParagraph(**args):
    return Element(qname = (TEXTNS,'hidden-paragraph'), **args)

def HiddenText(**args):
    return Element(qname = (TEXTNS,'hidden-text'), **args)

def IllustrationIndex(**args):
    return Element(qname = (TEXTNS,'illustration-index'), **args)

def IllustrationIndexEntryTemplate(**args):
    return Element(qname = (TEXTNS,'illustration-index-entry-template'), **args)

def IllustrationIndexSource(**args):
    return Element(qname = (TEXTNS,'illustration-index-source'), **args)

def ImageCount(**args):
    return Element(qname = (TEXTNS,'image-count'), **args)

def IndexBody(**args):
    return Element(qname = (TEXTNS,'index-body'), **args)

def IndexEntryBibliography(**args):
    return Element(qname = (TEXTNS,'index-entry-bibliography'), **args)

def IndexEntryChapter(**args):
    return Element(qname = (TEXTNS,'index-entry-chapter'), **args)

def IndexEntryLinkEnd(**args):
    return Element(qname = (TEXTNS,'index-entry-link-end'), **args)

def IndexEntryLinkStart(**args):
    return Element(qname = (TEXTNS,'index-entry-link-start'), **args)

def IndexEntryPageNumber(**args):
    return Element(qname = (TEXTNS,'index-entry-page-number'), **args)

def IndexEntrySpan(**args):
    return Element(qname = (TEXTNS,'index-entry-span'), **args)

def IndexEntryTabStop(**args):
    return Element(qname = (TEXTNS,'index-entry-tab-stop'), **args)

def IndexEntryText(**args):
    return Element(qname = (TEXTNS,'index-entry-text'), **args)

def IndexSourceStyle(**args):
    return Element(qname = (TEXTNS,'index-source-style'), **args)

def IndexSourceStyles(**args):
    return Element(qname = (TEXTNS,'index-source-styles'), **args)

def IndexTitle(**args):
    return Element(qname = (TEXTNS,'index-title'), **args)

def IndexTitleTemplate(**args):
    return Element(qname = (TEXTNS,'index-title-template'), **args)

def InitialCreator(**args):
    return Element(qname = (TEXTNS,'initial-creator'), **args)

def Insertion(**args):
    return Element(qname = (TEXTNS,'insertion'), **args)

def Keywords(**args):
    return Element(qname = (TEXTNS,'keywords'), **args)

def LineBreak(**args):
    return Element(qname = (TEXTNS,'line-break'), **args)

def LinenumberingConfiguration(**args):
    return Element(qname = (TEXTNS,'linenumbering-configuration'), **args)

def LinenumberingSeparator(**args):
    return Element(qname = (TEXTNS,'linenumbering-separator'), **args)

def List(**args):
    return Element(qname = (TEXTNS,'list'), **args)

def ListHeader(**args):
    return Element(qname = (TEXTNS,'list-header'), **args)

def ListItem(**args):
    return Element(qname = (TEXTNS,'list-item'), **args)

def ListLevelStyleBullet(**args):
    return Element(qname = (TEXTNS,'list-level-style-bullet'), **args)

def ListLevelStyleImage(**args):
    return Element(qname = (TEXTNS,'list-level-style-image'), **args)

def ListLevelStyleNumber(**args):
    return Element(qname = (TEXTNS,'list-level-style-number'), **args)

def ListStyle(**args):
    return StyleElement(qname = (TEXTNS,'list-style'), **args)

def Measure(**args):
    return Element(qname = (TEXTNS,'measure'), **args)

def Meta(**args):
    return Element(qname = (TEXTNS,'meta'), **args)

def MetaField(**args):
    return Element(qname = (TEXTNS,'meta-field'), **args)

def ModificationDate(**args):
    return Element(qname = (TEXTNS,'modification-date'), **args)

def ModificationTime(**args):
    return Element(qname = (TEXTNS,'modification-time'), **args)

def Note(**args):
    return Element(qname = (TEXTNS,'note'), **args)

def NoteBody(**args):
    return Element(qname = (TEXTNS,'note-body'), **args)

def NoteCitation(**args):
    return Element(qname = (TEXTNS,'note-citation'), **args)

def NoteContinuationNoticeBackward(**args):
    return Element(qname = (TEXTNS,'note-continuation-notice-backward'), **args)

def NoteContinuationNoticeForward(**args):
    return Element(qname = (TEXTNS,'note-continuation-notice-forward'), **args)

def NoteRef(**args):
    return Element(qname = (TEXTNS,'note-ref'), **args)

def NotesConfiguration(**args):
    return Element(qname = (TEXTNS,'notes-configuration'), **args)

def Number(**args):
    return Element(qname = (TEXTNS,'number'), **args)

def NumberedParagraph(**args):
    return Element(qname = (TEXTNS,'numbered-paragraph'), **args)

def ObjectCount(**args):
    return Element(qname = (TEXTNS,'object-count'), **args)

def ObjectIndex(**args):
    return Element(qname = (TEXTNS,'object-index'), **args)

def ObjectIndexEntryTemplate(**args):
    return Element(qname = (TEXTNS,'object-index-entry-template'), **args)

def ObjectIndexSource(**args):
    return Element(qname = (TEXTNS,'object-index-source'), **args)

def OutlineLevelStyle(**args):
    return Element(qname = (TEXTNS,'outline-level-style'), **args)

def OutlineStyle(**args):
    return Element(qname = (TEXTNS,'outline-style'), **args)

def P(**args):
    return Element(qname = (TEXTNS, 'p'), **args)

def Page(**args):
    return Element(qname = (TEXTNS,'page'), **args)

def PageContinuation(**args):
    return Element(qname = (TEXTNS,'page-continuation'), **args)

def PageCount(**args):
    return Element(qname = (TEXTNS,'page-count'), **args)

def PageNumber(**args):
    return Element(qname = (TEXTNS,'page-number'), **args)

def PageSequence(**args):
    return Element(qname = (TEXTNS,'page-sequence'), **args)

def PageVariableGet(**args):
    return Element(qname = (TEXTNS,'page-variable-get'), **args)

def PageVariableSet(**args):
    return Element(qname = (TEXTNS,'page-variable-set'), **args)

def ParagraphCount(**args):
    return Element(qname = (TEXTNS,'paragraph-count'), **args)

def Placeholder(**args):
    return Element(qname = (TEXTNS,'placeholder'), **args)

def PrintDate(**args):
    return Element(qname = (TEXTNS,'print-date'), **args)

def PrintTime(**args):
    return Element(qname = (TEXTNS,'print-time'), **args)

def PrintedBy(**args):
    return Element(qname = (TEXTNS,'printed-by'), **args)

def ReferenceMark(**args):
    return Element(qname = (TEXTNS,'reference-mark'), **args)

def ReferenceMarkEnd(**args):
    return Element(qname = (TEXTNS,'reference-mark-end'), **args)

def ReferenceMarkStart(**args):
    return Element(qname = (TEXTNS,'reference-mark-start'), **args)

def ReferenceRef(**args):
    return Element(qname = (TEXTNS,'reference-ref'), **args)

def Ruby(**args):
    return Element(qname = (TEXTNS,'ruby'), **args)

def RubyBase(**args):
    return Element(qname = (TEXTNS,'ruby-base'), **args)

def RubyText(**args):
    return Element(qname = (TEXTNS,'ruby-text'), **args)

def S(**args):
    return Element(qname = (TEXTNS,'s'), **args)

def Script(**args):
    return Element(qname = (TEXTNS,'script'), **args)

def Section(**args):
    return Element(qname = (TEXTNS,'section'), **args)

def SectionSource(**args):
    return Element(qname = (TEXTNS,'section-source'), **args)

def SenderCity(**args):
    return Element(qname = (TEXTNS,'sender-city'), **args)

def SenderCompany(**args):
    return Element(qname = (TEXTNS,'sender-company'), **args)

def SenderCountry(**args):
    return Element(qname = (TEXTNS,'sender-country'), **args)

def SenderEmail(**args):
    return Element(qname = (TEXTNS,'sender-email'), **args)

def SenderFax(**args):
    return Element(qname = (TEXTNS,'sender-fax'), **args)

def SenderFirstname(**args):
    return Element(qname = (TEXTNS,'sender-firstname'), **args)

def SenderInitials(**args):
    return Element(qname = (TEXTNS,'sender-initials'), **args)

def SenderLastname(**args):
    return Element(qname = (TEXTNS,'sender-lastname'), **args)

def SenderPhonePrivate(**args):
    return Element(qname = (TEXTNS,'sender-phone-private'), **args)

def SenderPhoneWork(**args):
    return Element(qname = (TEXTNS,'sender-phone-work'), **args)

def SenderPosition(**args):
    return Element(qname = (TEXTNS,'sender-position'), **args)

def SenderPostalCode(**args):
    return Element(qname = (TEXTNS,'sender-postal-code'), **args)

def SenderStateOrProvince(**args):
    return Element(qname = (TEXTNS,'sender-state-or-province'), **args)

def SenderStreet(**args):
    return Element(qname = (TEXTNS,'sender-street'), **args)

def SenderTitle(**args):
    return Element(qname = (TEXTNS,'sender-title'), **args)

def Sequence(**args):
    return Element(qname = (TEXTNS,'sequence'), **args)

def SequenceDecl(**args):
    return Element(qname = (TEXTNS,'sequence-decl'), **args)

def SequenceDecls(**args):
    return Element(qname = (TEXTNS,'sequence-decls'), **args)

def SequenceRef(**args):
    return Element(qname = (TEXTNS,'sequence-ref'), **args)

def SheetName(**args):
    return Element(qname = (TEXTNS,'sheet-name'), **args)

def SoftPageBreak(**args):
    return Element(qname = (TEXTNS,'soft-page-break'), **args)

def SortKey(**args):
    return Element(qname = (TEXTNS,'sort-key'), **args)

def Span(**args):
    return Element(qname = (TEXTNS,'span'), **args)

def Subject(**args):
    return Element(qname = (TEXTNS,'subject'), **args)

def Tab(**args):
    return Element(qname = (TEXTNS,'tab'), **args)

def TableCount(**args):
    return Element(qname = (TEXTNS,'table-count'), **args)

def TableFormula(**args):
    return Element(qname = (TEXTNS,'table-formula'), **args)

def TableIndex(**args):
    return Element(qname = (TEXTNS,'table-index'), **args)

def TableIndexEntryTemplate(**args):
    return Element(qname = (TEXTNS,'table-index-entry-template'), **args)

def TableIndexSource(**args):
    return Element(qname = (TEXTNS,'table-index-source'), **args)

def TableOfContent(**args):
    return Element(qname = (TEXTNS,'table-of-content'), **args)

def TableOfContentEntryTemplate(**args):
    return Element(qname = (TEXTNS,'table-of-content-entry-template'), **args)

def TableOfContentSource(**args):
    return Element(qname = (TEXTNS,'table-of-content-source'), **args)

def TemplateName(**args):
    return Element(qname = (TEXTNS,'template-name'), **args)

def TextInput(**args):
    return Element(qname = (TEXTNS,'text-input'), **args)

def Time(**args):
    return Element(qname = (TEXTNS,'time'), **args)

def Title(**args):
    return Element(qname = (TEXTNS,'title'), **args)

def TocMark(**args):
    return Element(qname = (TEXTNS,'toc-mark'), **args)

def TocMarkEnd(**args):
    return Element(qname = (TEXTNS,'toc-mark-end'), **args)

def TocMarkStart(**args):
    return Element(qname = (TEXTNS,'toc-mark-start'), **args)

def TrackedChanges(**args):
    return Element(qname = (TEXTNS,'tracked-changes'), **args)

def UserDefined(**args):
    return Element(qname = (TEXTNS,'user-defined'), **args)

def UserFieldDecl(**args):
    return Element(qname = (TEXTNS,'user-field-decl'), **args)

def UserFieldDecls(**args):
    return Element(qname = (TEXTNS,'user-field-decls'), **args)

def UserFieldGet(**args):
    return Element(qname = (TEXTNS,'user-field-get'), **args)

def UserFieldInput(**args):
    return Element(qname = (TEXTNS,'user-field-input'), **args)

def UserIndex(**args):
    return Element(qname = (TEXTNS,'user-index'), **args)

def UserIndexEntryTemplate(**args):
    return Element(qname = (TEXTNS,'user-index-entry-template'), **args)

def UserIndexMark(**args):
    return Element(qname = (TEXTNS,'user-index-mark'), **args)

def UserIndexMarkEnd(**args):
    return Element(qname = (TEXTNS,'user-index-mark-end'), **args)

def UserIndexMarkStart(**args):
    return Element(qname = (TEXTNS,'user-index-mark-start'), **args)

def UserIndexSource(**args):
    return Element(qname = (TEXTNS,'user-index-source'), **args)

def VariableDecl(**args):
    return Element(qname = (TEXTNS,'variable-decl'), **args)

def VariableDecls(**args):
    return Element(qname = (TEXTNS,'variable-decls'), **args)

def VariableGet(**args):
    return Element(qname = (TEXTNS,'variable-get'), **args)

def VariableInput(**args):
    return Element(qname = (TEXTNS,'variable-input'), **args)

def VariableSet(**args):
    return Element(qname = (TEXTNS,'variable-set'), **args)

def WordCount(**args):
    return Element(qname = (TEXTNS,'word-count'), **args)


""",
    },
    'odf.thumbnail': {
        'is_package': False,
        'source': r'''
#!/usr/bin/python
# -*- coding: utf-8 -*-
# This contains a 104x128 px thumbnail in PNG format
# Downloaded from http://da.libreoffice.org/assets/Uploads/Da-Projekt_Billeder/IconLibreOffice.png
# Copyright information: Unless otherwise specified, all text and images on the
# http://da.libreoffice.org website are licensed under the Creative Commons
# Attribution-Share Alike 3.0 License. (As of 2013-01-09)

# Alternative download: http://commons.wikimedia.org/wiki/File:LibreOffice_icon_3.3.1_48_px.svg
import base64

iconstr = """\
iVBORw0KGgoAAAANSUhEUgAAAGgAAACACAYAAADnCyxOAAAABHNCSVQICAgIfAhkiAAAAAlwSFlz
AAAG7AAABuwBHnU4NQAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAwBSURB
VHic7Z1bbFTHHca/2UsxYDs4IRiv7Y0xUNRUCCmAFBXiCwaiRhFIxjdsjG18g8YvlfoQqepDH/yA
hKI+EUIIKqVxwiVAuapUlh+qqi+8VH2hfUgDUpuHVnhjs17vnsv0oZ3lnNkzs2fXu2eX3flJR7ue
nbN7znz7/ec/M+esCaUUk5OTtQDOAOgAUAdnqKBc9lqm5bl8Lwrg7wD+BODWJ5988lDymY6cPHmy
F8AIpfRHAL6X6f5ZQAE8JYQ80DTtF5999tkiGR8fXw/gLwBCHhxAITkD4MPz588bbipPTEx8H8Bf
4Y0wKRBC/qbr+m4fgGmUvjgA8DMADyYmJta4qUwI6UCBxAEASuk2v9//Sx+Adwt1EAXgAIC7bkSi
lM5TSlHg7VCAUvqG9cCmp6dBqayLeLmYm5vD7OystagdwN3x8fH3P/300yXRfpWVlb9bWFh4Qgh5
Q1THAzb7+BJKKUzTLJmtra0N+/bt40+TiSR00kcffRQzTXPCNE0tx42eET7eVgxN0xCNRrG0tISl
pSXEYjHEYjEsLy9jeXkZ8Xgc8XgciUQCiUQCmqbZNl3Xoes6DMOAYRjJBuM/J99QStHe3o729nY+
fLRTSu+OjY0JRbp48eJD0zR7NU3TChXmyOjoqK21pqenYZomotEompqa8t6A+cAwDEQiEZimmSwj
hGB2dhZzc3N89TkA71+4cEEY7oaGhjoBfBkMBoN5OWAJthDHf7MJIS/lFggEUFNTA5/PZzu3jo4O
tLW18W3QDkDqpEuXLt0wTbM/Ho977qSUEOcU7l5G/H4/ampqQAhJno9pmujo6EBLS4tjuBsdHRWK
dPny5euGYRyLx+Oe9kkpDioVgYAXIvFOOnDgAFpbW/nq7QCkIs3MzFyllB6PxWJ6fo44FWGSUCqw
cMc7af/+/di7d6+Tk+6dOHFCKNLnn3/+pWEYQ9FoVPckxPEHUIpiBQIBvPrqqykiHTx4EHv27OHP
tQ2AVKQrV67MGIYxEo1G8+6kkg5xVmQivf322xmLdO3atd/quj76/PlzV3N72SJMEkqRYDCI1157
DYQQW/l7772HXbt22cZplNI2Sum9kZERoUhfffXVbzRNG1tYWDA8C3GMUhZp/fr1AOwR49ChQ9i5
c6dt7IT/O0km0s2bN39tGMbEwsJCXpxkc1AhRvqFgDmJP/fDhw9jx44dMAwjxUnDw8MykS4ahnEy
Eonk3EllkSQ4sWrVKrz++uu286SUoqurC9u3b3d0kkykW7duXdB1/Sfz8/OmqE42CGcS+DhdijCR
rI4xTRM9PT148803YRi2qNUG4P7w8PBa0fvduXPnPKX0g2fPnuVMJGGSwH2DSpaKigps2LAhRaTe
3l5s27YNuq5b26WVUnpvaGhIKNLt27fPUUqnnj17ZuY8xDEHWR/LYauoqEBtbS3f96C/vx9bt27l
ndQKQCrSnTt3PqaUTkUikRV/y8veQQyrSFYGBwfR3NwMTdNSnHT8+PF0In3w3XffrchJwjS7HPog
ntWrV2Pjxo28GBgaGsKmTZug67aJg1YAUpFu3759zjTNqcXFxay/7cKZhHJzEGP16tWoq6tDIpGw
tcfIyAjC4XA2In0MYOr58+dZNWhZzSS4Zc2aNQiFQkgkErbysbGxZDkf7gYHB4Ui3bx5MylSTpIE
BbB27VqEQiHE43Fbg01OTqKurg6aZlsWagUgFenGjRsfA/ggFotl1MgqSZBQWVmJUCiE5eVlW9uc
OnUKtbW1Tk66f+zYMZlI50zTnIrFYjTrJEGFOTtVVVWor69HLBZLESkUCvFOagGQTqSzAKbi8bir
BhYmCYoXVFdXo7GxEbFYLFlGCMHk5CTq6+v5xCGtSNevXz8LYCqRSKRt6JJfUc0V1dXVCIfDNicB
wPj4OOrr6/nUvIVSen9gYEAo0rVr185SSqcSiYQ03EkX7NRm36qqqtDY2Gjrk4D/ZXcNDQ2OTkon
EoApXddFzoiU7FU9+YKFu3g8bisfGxtDY2MjP3fXQim939/fLxTp6tWrZymlU4ZhODnpYdkt2OUC
Fu747O7EiRMIh8P8dFELADciHTJN84ml+F8AfqWyuCypqqpCOBxOGScNDw+jqamJH6akFenKlSt3
5+fnf0gp3U8pfd/n8/1gZmbmz6Svry+pRDQaxenTp2EYBqLRKN566608nmJpsLi4iCdPnoC/Kvjy
5cv45ptvbNfkAfgjgB9/8cUXUbfvrxy0QpiTuAQBg4OD2LRpE++kdwA8OHr0qNBJPMI+SOEeljjw
s+ADAwNobm7mv+xMpEo37y28aEQ5KDNY4sBlcTh69Ciam5v5tn2HUnq/r68vrUiODlICZQcTiZ/H
7Ovrw5YtW5yclFYk4UyCEig7qqur0dDQkBKNuru7kyLxTurt7RWKVDaX/nrJK6+8gsbGxpTrDLu6
urB161a++jsAhCIJszj+G6C2zGZcWLizvk4IwZEjR0QiPXASSU315ABR+7G5O75Njxw54hTu9lJK
H/T09NhEEq6oKoFyA0vBAbuQnZ2d2LJlC199LwCbSMpBHuAkEgB0dnYmx0m8k7q7uysBNZPgGUwk
6+VshBB0dXVh8+bNfPW9AM4Bae7yVuQWlt0BSOmTmpub+eoD3d3dPSrEeYxMpKamJl6DU9I0W5Ef
1q1bh3A4nBLuuru7+apNAT5zUw7yhnXr1oFSiqdPn8I0zeQPcHDt3qDS7AJSU1PjOE6yEAjwJcpF
3lJTU4Ovv/4aPp8v2ebW0KdCXJHBt7tKs4sAmTGEDlJZnHewJMFJJDWTUATIkjO1HlQEyH6fIiXE
OT1X5Be+a7FeqiW8JkHhHVb3ZJQklLJQxXSTtJMGjICoYqlTTOcq6/vVZGkRkFGSoLI477G2Nd+1
qHFQESBr84ASovDIkjPloCJA1gcJBVJ4h2yCQIW4IkB2PYhwwY49V+SfrJYblDjeYU0MpMsN1jUJ
hXeIxAHULZBFQUbjIOUg71HjoCJH1tYqzS4CnG6WY6iBahEgTbPT7aBYOZksDqadSVDC5J507el6
HKQoDK7TbDWTUBhcp9miqxsV+UXW1upfAxQBst9IEiYJykHekdVygxLJO1wv2ClhCoermQTV7xQG
mTHUVE+RwNqdN4kaBxUZrkOcCnfeYTWFclCRIrpYR6XZecZNJJLVsQmkpnpyj5t2zCjE8Tuqvii/
8MsMGa+oKiflDydxpA6y7mjdQTkp91jFsbat6wU7/iJG5aTcIQtrvBGkl13xOyuRVo6TONa/eaTL
Dfw9+yrcrQwncZzuBXY1DgKQ/E9ShBDbpkTKHFFYY23sykF8Bd5BDCVSZojEYQ6SdSVpB6qifkeJ
5A6ZOE7TbDzSJIHFR5U4ZIcbcfg+yPVMAiEkaT/mFieUk5xxI47MBAzpZKn1g5RI7nErjlO9tDMJ
TvZTIrknU3H4EJeRg6yqKpHSk61zrM9THCT7QOtlqEokOdmKw/dBPGlnEtgPnvKDVP55OQ9mV9Ln
ZDQO4ne2jnIZykl2ViIOAP7/fqcgDHGsgdlsghIplZWK43RfUEYXjRiGoUQSkCtx+Lk4aRbHrwGx
Psjn8ymRLORKHNa+MqQO0nU96SImDnu0JgRs30wTh5dRsFyLw9rXdZJgxTAMm8IiJwF2l7h1kqhe
seJWHKc6QKo4pmlC13VbfZ60d3nzFsy1SC8L+RDHKcXOeKpHiZRfcdg40/pZrkOcNQ0sV5Hy7RzZ
7Y9AmhVV1gexVdVyE8kLcZwGqhk5iIlSbiJ5IY51Kk1EyjjICp+nl4tIXoqT0YoqXzEYDCISicDn
88Hn88Hv9yefE0Lg9/uTZX6/P1lGCEnWYY/sOYDkexTjJV1ODc83plMfYp0ZYI8shFlnZFgZex4M
BlOSBCvSEPf48eO8nby1EZyeW8tkr1nLGOyErQNpUVk2r3uJ9PfiNE3z8FAK1wg8IuG9/GxGioP8
fj8qKio8OyDFC/x+f0oZ2b179z8BhLw6CFG4Svfotq4oJGXzWOjwBuDfAUrp7wGMePmpxRLKMqUA
IW/WB+BDAN96/cmKtHwL4KeEUopdu3bVAjgD4CCADYU9rrLnHwD+AODnjx49+s9/AexyD6bH7vMJ
AAAAAElFTkSuQmCC\
"""

def thumbnail():
    icon = base64.b64decode(iconstr)
    return icon

if __name__ == "__main__":
    icon = thumbnail()
    f = open("thumbnail.png","wb")
    f.write(icon)
    f.close()

''',
    },
    'odf.userfield': {
        'is_package': False,
        'source': r'''
#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2009 Søren Roug, European Environment Agency
#
# This is free software.  You may redistribute it under the terms
# of the Apache license and the GNU General Public License Version
# 2 or at your option any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public
# License along with this program; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s): Michael Howitz, gocept gmbh & co. kg
#
# $Id: userfield.py 447 2008-07-10 20:01:30Z roug $

"""Class to show and manipulate user fields in odf documents."""

import sys
import zipfile

from odf.text import UserFieldDecl
from odf.namespaces import OFFICENS
from odf.opendocument import load
import io, sys

if sys.version_info[0]==3:
    unicode=str

OUTENCODING = "utf-8"


# OpenDocument v.1.0 section 6.7.1
VALUE_TYPES = {
    u'float': (OFFICENS, u'value'),
    u'percentage': (OFFICENS, u'value'),
    u'currency': (OFFICENS, u'value'),
    u'date': (OFFICENS, u'date-value'),
    u'time': (OFFICENS, u'time-value'),
    u'boolean': (OFFICENS, u'boolean-value'),
    u'string': (OFFICENS, u'string-value'),
    }


class UserFields(object):
    """List, view and manipulate user fields."""

    # these attributes can be a filename or a file like object
    src_file = None
    dest_file = None

    def __init__(self, src=None, dest=None):
        """Constructor

        @param src open file in binary mode: source document,
        or filename as a unicode string, or None for stdin.
        @param dest opendile in binary mode: destination document,
        or filename as a unicode string, or None for stdout.
        """
        assert(src==None or 'rb' in repr(src) or 'BufferedReader' in repr(src) or 'BytesIO' in repr(src) or type(src)==type(u""))
        assert(dest==None or 'wb' in repr(dest) or 'BufferedWriter' in repr(dest) or 'BytesIO' in repr(dest) or type(dest)==type(u""))
        self.src_file = src
        self.dest_file = dest
        self.document = None

    def loaddoc(self):
        if (sys.version_info[0]==3 and (isinstance(self.src_file, str) or (isinstance(self.src_file, io.IOBase)))) or (sys.version_info[0]==2 and isinstance(self.src_file, basestring)):
            # src_file is a filename, check if it is a zip-file
            if not zipfile.is_zipfile(self.src_file):
                raise TypeError(u"%s is no odt file." % self.src_file)
        elif self.src_file is None:
            # use stdin if no file given
            self.src_file = sys.stdin

        self.document = load(self.src_file)

    def savedoc(self):
        # write output
        if self.dest_file is None:
            # use stdout if no filename given
            self.document.save(u'-')
        else:
            self.document.save(self.dest_file)

    def list_fields(self):
        """List (extract) all known user-fields.

        @return list of user-field names as unicode strings.
        """
        return [x[0] for x in self.list_fields_and_values()]

    def list_fields_and_values(self, field_names=None):
        """List (extract) user-fields with type and value.

        @param field_names list of field names as unicode strings
        to show, or None for all.

        @return list of tuples (<field name>, <field type>, <value>)
        as type (unicode string, stringified type, unicode string).

        """
        self.loaddoc()
        found_fields = []
        all_fields = self.document.getElementsByType(UserFieldDecl)
        for f in all_fields:
            value_type = f.getAttribute(u'valuetype')
            if value_type == u'string':
                value = f.getAttribute(u'stringvalue')
            else:
                value = f.getAttribute(u'value')
            field_name = f.getAttribute(u'name')

            if field_names is None or field_name in field_names:
                found_fields.append((field_name,
                                     value_type,
                                     value))
        return found_fields

    def list_values(self, field_names):
        """Extract the contents of given field names from the file.

        @param field_names list of field names as unicode strings

        @return list of field values as unicode strings.

        """
        return [x[2] for x in self.list_fields_and_values(field_names)]

    def get(self, field_name):
        """Extract the contents of this field from the file.
        @param field_name unicode string: name of a field
        @return field value as a unicode string or None if field does not exist.

        """
        assert(type(field_name)==type(u""))
        values = self.list_values([field_name])
        if not values:
            return None
        return values[0]

    def get_type_and_value(self, field_name):
        """Extract the type and contents of this field from the file.
        @param field_name unicode string: name of a field
        @return tuple (<type>, <field-value>) as a pair of unicode strings
        or None if field does not exist.

        """
        assert(type(field_name)==type(u""))
        fields = self.list_fields_and_values([field_name])
        if not fields:
            return None
        field_name, value_type, value = fields[0]
        return value_type, value

    def update(self, data):
        """Set the value of user fields. The field types will be the same.

        data ... dict, with field name as key, field value as value

        Returns None

        """
        self.loaddoc()
        all_fields = self.document.getElementsByType(UserFieldDecl)
        for f in all_fields:
            field_name = f.getAttribute(u'name')
            if field_name in data:
                value_type = f.getAttribute(u'valuetype')
                value = data.get(field_name)
                if value_type == u'string':
                    f.setAttribute(u'stringvalue', value)
                else:
                    f.setAttribute(u'value', value)
        self.savedoc()


''',
    },
    'odf.xforms': {
        'is_package': False,
        'source': r"""
# -*- coding: utf-8 -*-
# Copyright (C) 2006-2007 Søren Roug, European Environment Agency
#
# This library is free software; you can redistribute it and/or
# modify it under the terms of the GNU Lesser General Public
# License as published by the Free Software Foundation; either
# version 2.1 of the License, or (at your option) any later version.
#
# This library is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
# Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public
# License along with this library; if not, write to the Free Software
# Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
#
# Contributor(s):
#

from odf.namespaces import XFORMSNS
from odf.element import Element

# ODF 1.0 section 11.2
# XForms is designed to be embedded in another XML format.
# Autogenerated
def Model(**args):
    return Element(qname = (XFORMSNS,'model'), **args)

def Instance(**args):
    return Element(qname = (XFORMSNS,'instance'), **args)

def Bind(**args):
    return Element(qname = (XFORMSNS,'bind'), **args)

""",
    },
}


class _BundledLoader(importlib.abc.Loader):
    def __init__(self, fullname, source, is_package):
        self.fullname = fullname
        self.source = source
        self.is_package = is_package

    def create_module(self, spec):
        return None

    def exec_module(self, module):
        module_path = self.fullname.replace(".", "/") + ".py"
        module.__file__ = "<%s:%s>" % (_BUNDLE_TAG, module_path)
        code = compile(
            self.source,
            module.__file__,
            "exec",
            dont_inherit=True,
        )
        exec(code, module.__dict__)
        if self.is_package:
            package_path = self.fullname.replace(".", "/")
            module.__path__ = ["<%s:%s>" % (_BUNDLE_TAG, package_path)]


class _BundledFinder(importlib.abc.MetaPathFinder):
    def __init__(self, prefixes):
        self._prefixes = prefixes

    def _owns(self, fullname):
        for prefix in self._prefixes:
            if fullname == prefix or fullname.startswith(prefix + "."):
                return True
        return False

    def find_spec(self, fullname, path, target=None):
        if not self._owns(fullname):
            return None
        item = _SOURCES.get(fullname)
        if item is None:
            return None
        loader = _BundledLoader(fullname, item["source"], item["is_package"])
        return importlib.util.spec_from_loader(
            fullname,
            loader,
            is_package=item["is_package"],
        )


def install():
    """Зарегистрировать bundled-пакеты в sys.meta_path."""
    global _FINDER
    if _FINDER is not None:
        return
    _FINDER = _BundledFinder(_PREFIXES)
    sys.meta_path.insert(0, _FINDER)


install()

