# -*- coding: utf-8 -*-
"""
openpyxl + et_xmlfile в одном файле для LibreOffice pythonpath.

Версии: et_xmlfile 2.0.0, openpyxl 3.1.5
Сборка: python macro-lib/bundle_openpyxl.py

Перед import целевых пакетов выполните:
  import openpyxl_bundled
"""
from __future__ import annotations
MACRO_VERSION = "3.10.708"
import importlib.abc
import importlib.util
import sys

_FINDER = None
_BUNDLE_TAG = "openpyxl_bundled"

_PREFIXES = ('et_xmlfile', 'openpyxl')
_SOURCES = {
    'et_xmlfile': {
        'is_package': True,
        'source': r"""
from .xmlfile import xmlfile

# constants
__version__ = '2.0.0'
__author__ = 'See AUTHORS.txt'
__license__ = 'MIT'
__author_email__ = 'charlie.clark@clark-consulting.eu'
__url__ = 'https://foss.heptapod.net/openpyxl/et_xmlfile'

""",
    },
    'et_xmlfile.incremental_tree': {
        'is_package': False,
        'source': r'''
# Code modified from cPython's Lib/xml/etree/ElementTree.py
# The write() code is modified to allow specifying a particular namespace
# uri -> prefix mapping.
#
# ---------------------------------------------------------------------
# Licensed to PSF under a Contributor Agreement.
# See https://www.python.org/psf/license for licensing details.
#
# ElementTree
# Copyright (c) 1999-2008 by Fredrik Lundh.  All rights reserved.
#
# fredrik@pythonware.com
# http://www.pythonware.com
# --------------------------------------------------------------------
# The ElementTree toolkit is
#
# Copyright (c) 1999-2008 by Fredrik Lundh
#
# By obtaining, using, and/or copying this software and/or its
# associated documentation, you agree that you have read, understood,
# and will comply with the following terms and conditions:
#
# Permission to use, copy, modify, and distribute this software and
# its associated documentation for any purpose and without fee is
# hereby granted, provided that the above copyright notice appears in
# all copies, and that both that copyright notice and this permission
# notice appear in supporting documentation, and that the name of
# Secret Labs AB or the author not be used in advertising or publicity
# pertaining to distribution of the software without specific, written
# prior permission.
#
# SECRET LABS AB AND THE AUTHOR DISCLAIMS ALL WARRANTIES WITH REGARD
# TO THIS SOFTWARE, INCLUDING ALL IMPLIED WARRANTIES OF MERCHANT-
# ABILITY AND FITNESS.  IN NO EVENT SHALL SECRET LABS AB OR THE AUTHOR
# BE LIABLE FOR ANY SPECIAL, INDIRECT OR CONSEQUENTIAL DAMAGES OR ANY
# DAMAGES WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS,
# WHETHER IN AN ACTION OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS
# ACTION, ARISING OUT OF OR IN CONNECTION WITH THE USE OR PERFORMANCE
# OF THIS SOFTWARE.
# --------------------------------------------------------------------
import contextlib
import io

import xml.etree.ElementTree as ET


def current_global_nsmap():
    return {
        prefix: uri for uri, prefix in ET._namespace_map.items()
    }


class IncrementalTree(ET.ElementTree):

    def write(
        self,
        file_or_filename,
        encoding=None,
        xml_declaration=None,
        default_namespace=None,
        method=None,
        *,
        short_empty_elements=True,
        nsmap=None,
        root_ns_only=False,
        minimal_ns_only=False,
    ):
        """Write element tree to a file as XML.

        Arguments:
          *file_or_filename* -- file name or a file object opened for writing

          *encoding* -- the output encoding (default: US-ASCII)

          *xml_declaration* -- bool indicating if an XML declaration should be
                               added to the output. If None, an XML declaration
                               is added if encoding IS NOT either of:
                               US-ASCII, UTF-8, or Unicode

          *default_namespace* -- sets the default XML namespace (for "xmlns").
                                 Takes precedence over any default namespace
                                 provided in nsmap or
                                 xml.etree.ElementTree.register_namespace().

          *method* -- either "xml" (default), "html, "text", or "c14n"

          *short_empty_elements* -- controls the formatting of elements
                                    that contain no content. If True (default)
                                    they are emitted as a single self-closed
                                    tag, otherwise they are emitted as a pair
                                    of start/end tags

          *nsmap* -- a mapping of namespace prefixes to URIs. These take
                     precedence over any mappings registered using
                     xml.etree.ElementTree.register_namespace(). The
                     default_namespace argument, if supplied, takes precedence
                     over any default namespace supplied in nsmap. All supplied
                     namespaces will be declared on the root element, even if
                     unused in the document.

          *root_ns_only* -- bool indicating namespace declrations should only
                            be written on the root element.  This requires two
                            passes of the xml tree adding additional time to
                            the writing process. This is primarily meant to
                            mimic xml.etree.ElementTree's behaviour.

          *minimal_ns_only* -- bool indicating only namespaces that were used
                               to qualify elements or attributes should be
                               declared. All namespace declarations will be
                               written on the root element regardless of the
                               value of the root_ns_only arg. Requires two
                               passes of the xml tree adding additional time to
                               the writing process.

        """
        if not method:
            method = "xml"
        elif method not in ("text", "xml", "html"):
            raise ValueError("unknown method %r" % method)
        if not encoding:
            encoding = "us-ascii"

        with _get_writer(file_or_filename, encoding) as (write, declared_encoding):
            if method == "xml" and (
                xml_declaration
                or (
                    xml_declaration is None
                    and encoding.lower() != "unicode"
                    and declared_encoding.lower() not in ("utf-8", "us-ascii")
                )
            ):
                write("<?xml version='1.0' encoding='%s'?>\n" % (declared_encoding,))
            if method == "text":
                ET._serialize_text(write, self._root)
            else:
                if method == "xml":
                    is_html = False
                else:
                    is_html = True
                if nsmap:
                    if None in nsmap:
                        raise ValueError(
                            'Found None as default nsmap prefix in nsmap. '
                            'Use "" as the default namespace prefix.'
                        )
                    new_nsmap = nsmap.copy()
                else:
                    new_nsmap = {}
                if default_namespace:
                    new_nsmap[""] = default_namespace
                if root_ns_only or minimal_ns_only:
                    # _namespaces returns a mapping of only the namespaces that
                    # were used.
                    new_nsmap = _namespaces(
                        self._root,
                        default_namespace,
                        new_nsmap,
                    )
                    if not minimal_ns_only:
                        if nsmap:
                            # We want all namespaces defined in the provided
                            # nsmap to be declared regardless of whether
                            # they've been used.
                            new_nsmap.update(nsmap)
                        if default_namespace:
                            new_nsmap[""] = default_namespace
                global_nsmap = {
                    prefix: uri for uri, prefix in ET._namespace_map.items()
                }
                if None in global_nsmap:
                    raise ValueError(
                        'Found None as default nsmap prefix in nsmap registered with '
                        'register_namespace. Use "" for the default namespace prefix.'
                    )
                nsmap_scope = {}
                _serialize_ns_xml(
                    write,
                    self._root,
                    nsmap_scope,
                    global_nsmap,
                    is_html=is_html,
                    is_root=True,
                    short_empty_elements=short_empty_elements,
                    new_nsmap=new_nsmap,
                )


def _make_new_ns_prefix(
    nsmap_scope,
    global_prefixes,
    local_nsmap=None,
    default_namespace=None,
):
    i = len(nsmap_scope)
    if default_namespace is not None and "" not in nsmap_scope:
        # Keep the same numbering scheme as python which assumes the default
        # namespace is present if supplied.
        i += 1

    while True:
        prefix = f"ns{i}"
        if (
            prefix not in nsmap_scope
            and prefix not in global_prefixes
            and (
                not local_nsmap or prefix not in local_nsmap
            )
        ):
            return prefix
        i += 1


def _get_or_create_prefix(
    uri,
    nsmap_scope,
    global_nsmap,
    new_namespace_prefixes,
    uri_to_prefix,
    for_default_namespace_attr_prefix=False,
):
    """Find a prefix that doesn't conflict with the ns scope or create a new prefix

    This function mutates nsmap_scope, global_nsmap, new_namespace_prefixes and
    uri_to_prefix. It is intended to keep state in _serialize_ns_xml consistent
    while deduplicating the house keeping code or updating these dictionaries.
    """
    # Check if we can reuse an existing (global) prefix within the current
    # namespace scope. There maybe many prefixes pointing to a single URI by
    # this point and we need to select a prefix that is not in use in the
    # current scope.
    for global_prefix, global_uri in global_nsmap.items():
        if uri == global_uri and global_prefix not in nsmap_scope:
            prefix = global_prefix
            break
    else:  # no break
        # We couldn't find a suitable existing prefix for this namespace scope,
        # let's create a new one.
        prefix = _make_new_ns_prefix(nsmap_scope, global_prefixes=global_nsmap)
        global_nsmap[prefix] = uri
    nsmap_scope[prefix] = uri
    if not for_default_namespace_attr_prefix:
        # Don't override the actual default namespace prefix
        uri_to_prefix[uri] = prefix
    if prefix != "xml":
        new_namespace_prefixes.add(prefix)
    return prefix


def _find_default_namespace_attr_prefix(
    default_namespace,
    nsmap,
    local_nsmap,
    global_prefixes,
    provided_default_namespace=None,
):
    # Search the provided nsmap for any prefixes for this uri that aren't the
    # default namespace ""
    for prefix, uri in nsmap.items():
        if uri == default_namespace and prefix != "":
            return prefix

    for prefix, uri in local_nsmap.items():
        if uri == default_namespace and prefix != "":
            return prefix

    # _namespace_map is a 1:1 mapping of uri -> prefix
    prefix = ET._namespace_map.get(default_namespace)
    if prefix and prefix not in nsmap:
        return prefix

    return _make_new_ns_prefix(
        nsmap,
        global_prefixes,
        local_nsmap,
        provided_default_namespace,
    )


def process_attribs(
    elem,
    is_nsmap_scope_changed,
    default_ns_attr_prefix,
    nsmap_scope,
    global_nsmap,
    new_namespace_prefixes,
    uri_to_prefix,
):
    item_parts = []
    for k, v in elem.items():
        if isinstance(k, ET.QName):
            k = k.text
        try:
            if k[:1] == "{":
                uri_and_name = k[1:].rsplit("}", 1)
                try:
                    prefix = uri_to_prefix[uri_and_name[0]]
                except KeyError:
                    if not is_nsmap_scope_changed:
                        # We're about to mutate the these dicts so
                        # let's copy them first. We don't have to
                        # recompute other mappings as we're looking up
                        # or creating a new prefix
                        nsmap_scope = nsmap_scope.copy()
                        uri_to_prefix = uri_to_prefix.copy()
                        is_nsmap_scope_changed = True
                    prefix = _get_or_create_prefix(
                        uri_and_name[0],
                        nsmap_scope,
                        global_nsmap,
                        new_namespace_prefixes,
                        uri_to_prefix,
                    )

                if not prefix:
                    if default_ns_attr_prefix:
                        prefix = default_ns_attr_prefix
                    else:
                        for prefix, known_uri in nsmap_scope.items():
                            if known_uri == uri_and_name[0] and prefix != "":
                                default_ns_attr_prefix = prefix
                                break
                        else:  # no break
                            if not is_nsmap_scope_changed:
                                # We're about to mutate the these dicts so
                                # let's copy them first. We don't have to
                                # recompute other mappings as we're looking up
                                # or creating a new prefix
                                nsmap_scope = nsmap_scope.copy()
                                uri_to_prefix = uri_to_prefix.copy()
                                is_nsmap_scope_changed = True
                            prefix = _get_or_create_prefix(
                                uri_and_name[0],
                                nsmap_scope,
                                global_nsmap,
                                new_namespace_prefixes,
                                uri_to_prefix,
                                for_default_namespace_attr_prefix=True,
                            )
                            default_ns_attr_prefix = prefix
                k = f"{prefix}:{uri_and_name[1]}"
        except TypeError:
            ET._raise_serialization_error(k)

        if isinstance(v, ET.QName):
            if v.text[:1] != "{":
                v = v.text
            else:
                uri_and_name = v.text[1:].rsplit("}", 1)
                try:
                    prefix = uri_to_prefix[uri_and_name[0]]
                except KeyError:
                    if not is_nsmap_scope_changed:
                        # We're about to mutate the these dicts so
                        # let's copy them first. We don't have to
                        # recompute other mappings as we're looking up
                        # or creating a new prefix
                        nsmap_scope = nsmap_scope.copy()
                        uri_to_prefix = uri_to_prefix.copy()
                        is_nsmap_scope_changed = True
                    prefix = _get_or_create_prefix(
                        uri_and_name[0],
                        nsmap_scope,
                        global_nsmap,
                        new_namespace_prefixes,
                        uri_to_prefix,
                    )
                v = f"{prefix}:{uri_and_name[1]}"
        item_parts.append((k, v))
    return item_parts, default_ns_attr_prefix, nsmap_scope


def write_elem_start(
    write,
    elem,
    nsmap_scope,
    global_nsmap,
    short_empty_elements,
    is_html,
    is_root=False,
    uri_to_prefix=None,
    default_ns_attr_prefix=None,
    new_nsmap=None,
    **kwargs,
):
    """Write the opening tag (including self closing) and element text.

    Refer to _serialize_ns_xml for description of arguments.

    nsmap_scope should be an empty dictionary on first call. All nsmap prefixes
    must be strings with the default namespace prefix represented by "".

    eg.
    - <foo attr1="one">      (returns tag = 'foo')
    - <foo attr1="one">text  (returns tag = 'foo')
    - <foo attr1="one" />    (returns tag = None)

    Returns:
        tag:
            The tag name to be closed or None if no closing required.
        nsmap_scope:
            The current nsmap after any prefix to uri additions from this
            element. This is the input dict if unmodified or an updated copy.
        default_ns_attr_prefix:
            The prefix for the default namespace to use with attrs.
        uri_to_prefix:
            The current uri to prefix map after any uri to prefix additions
            from this element. This is the input dict if unmodified or an
            updated copy.
        next_remains_root:
            A bool indicating if the child element(s) should be treated as
            their own roots.
    """
    tag = elem.tag
    text = elem.text

    if tag is ET.Comment:
        write("<!--%s-->" % text)
        tag = None
        next_remains_root = False
    elif tag is ET.ProcessingInstruction:
        write("<?%s?>" % text)
        tag = None
        next_remains_root = False
    else:
        if new_nsmap:
            is_nsmap_scope_changed = True
            nsmap_scope = nsmap_scope.copy()
            nsmap_scope.update(new_nsmap)
            new_namespace_prefixes = set(new_nsmap.keys())
            new_namespace_prefixes.discard("xml")
            # We need to recompute the uri to prefixes
            uri_to_prefix = None
            default_ns_attr_prefix = None
        else:
            is_nsmap_scope_changed = False
            new_namespace_prefixes = set()

        if uri_to_prefix is None:
            if None in nsmap_scope:
                raise ValueError(
                    'Found None as a namespace prefix. Use "" as the default namespace prefix.'
                )
            uri_to_prefix = {uri: prefix for prefix, uri in nsmap_scope.items()}
            if "" in nsmap_scope:
                # There may be multiple prefixes for the default namespace but
                # we want to make sure we preferentially use "" (for elements)
                uri_to_prefix[nsmap_scope[""]] = ""

        if tag is None:
            # tag supression where tag is set to None
            # Don't change is_root so namespaces can be passed down
            next_remains_root = is_root
            if text:
                write(ET._escape_cdata(text))
        else:
            next_remains_root = False
            if isinstance(tag, ET.QName):
                tag = tag.text
            try:
                # These splits / fully qualified tag creationg are the
                # bottleneck in this implementation vs the python
                # implementation.
                # The following split takes ~42ns with no uri and ~85ns if a
                # prefix is present. If the uri was present, we then need to
                # look up a prefix (~14ns) and create the fully qualified
                # string (~41ns).  This gives a total of ~140ns where a uri is
                # present.
                # Python's implementation needs to preprocess the tree to
                # create a dict of qname -> tag by traversing the tree which
                # takes a bit of extra time but it quickly makes that back by
                # only having to do a dictionary look up (~14ns) for each tag /
                # attrname vs our splitting (~140ns).
                # So here we have the flexibility of being able to redefine the
                # uri a prefix points to midway through serialisation at the
                # expense of performance (~10% slower for a 1mb file on my
                # machine).
                if tag[:1] == "{":
                    uri_and_name = tag[1:].rsplit("}", 1)
                    try:
                        prefix = uri_to_prefix[uri_and_name[0]]
                    except KeyError:
                        if not is_nsmap_scope_changed:
                            # We're about to mutate the these dicts so let's
                            # copy them first. We don't have to recompute other
                            # mappings as we're looking up or creating a new
                            # prefix
                            nsmap_scope = nsmap_scope.copy()
                            uri_to_prefix = uri_to_prefix.copy()
                            is_nsmap_scope_changed = True
                        prefix = _get_or_create_prefix(
                            uri_and_name[0],
                            nsmap_scope,
                            global_nsmap,
                            new_namespace_prefixes,
                            uri_to_prefix,
                        )
                    if prefix:
                        tag = f"{prefix}:{uri_and_name[1]}"
                    else:
                        tag = uri_and_name[1]
                elif "" in nsmap_scope:
                    raise ValueError(
                        "cannot use non-qualified names with default_namespace option"
                    )
            except TypeError:
                ET._raise_serialization_error(tag)

            write("<" + tag)

            if elem.attrib:
                item_parts, default_ns_attr_prefix, nsmap_scope = process_attribs(
                    elem,
                    is_nsmap_scope_changed,
                    default_ns_attr_prefix,
                    nsmap_scope,
                    global_nsmap,
                    new_namespace_prefixes,
                    uri_to_prefix,
                )
            else:
                item_parts = []
            if new_namespace_prefixes:
                ns_attrs = []
                for k in sorted(new_namespace_prefixes):
                    v = nsmap_scope[k]
                    if k:
                        k = "xmlns:" + k
                    else:
                        k = "xmlns"
                    ns_attrs.append((k, v))
                if is_html:
                    write("".join([f' {k}="{ET._escape_attrib_html(v)}"' for k, v in ns_attrs]))
                else:
                    write("".join([f' {k}="{ET._escape_attrib(v)}"' for k, v in ns_attrs]))
            if item_parts:
                if is_html:
                    write("".join([f' {k}="{ET._escape_attrib_html(v)}"' for k, v in item_parts]))
                else:
                    write("".join([f' {k}="{ET._escape_attrib(v)}"' for k, v in item_parts]))
            if is_html:
                write(">")
                ltag = tag.lower()
                if text:
                    if ltag == "script" or ltag == "style":
                        write(text)
                    else:
                        write(ET._escape_cdata(text))
                if ltag in ET.HTML_EMPTY:
                    tag = None
            elif text or len(elem) or not short_empty_elements:
                write(">")
                if text:
                    write(ET._escape_cdata(text))
            else:
                tag = None
                write(" />")
    return (
        tag,
        nsmap_scope,
        default_ns_attr_prefix,
        uri_to_prefix,
        next_remains_root,
    )


def _serialize_ns_xml(
    write,
    elem,
    nsmap_scope,
    global_nsmap,
    short_empty_elements,
    is_html,
    is_root=False,
    uri_to_prefix=None,
    default_ns_attr_prefix=None,
    new_nsmap=None,
    **kwargs,
):
    """Serialize an element or tree using 'write' for output.

    Args:
        write:
            A function to write the xml to its destination.
        elem:
            The element to serialize.
        nsmap_scope:
            The current prefix to uri mapping for this element. This should be
            an empty dictionary for the root element. Additional namespaces are
            progressively added using the new_nsmap arg.
        global_nsmap:
            A dict copy of the globally registered _namespace_map in uri to
            prefix form
        short_empty_elements:
          Controls the formatting of elements that contain no content. If True
          (default) they are emitted as a single self-closed tag, otherwise
          they are emitted as a pair of start/end tags.
        is_html:
            Set to True to serialize as HTML otherwise XML.
        is_root:
            Boolean indicating if this is a root element.
        uri_to_prefix:
            Current state of the mapping of uri to prefix.
        default_ns_attr_prefix:
        new_nsmap:
            New prefix -> uri mapping to be applied to this element.
    """
    (
        tag,
        nsmap_scope,
        default_ns_attr_prefix,
        uri_to_prefix,
        next_remains_root,
    ) = write_elem_start(
        write,
        elem,
        nsmap_scope,
        global_nsmap,
        short_empty_elements,
        is_html,
        is_root,
        uri_to_prefix,
        default_ns_attr_prefix,
        new_nsmap=new_nsmap,
    )
    for e in elem:
        _serialize_ns_xml(
            write,
            e,
            nsmap_scope,
            global_nsmap,
            short_empty_elements,
            is_html,
            next_remains_root,
            uri_to_prefix,
            default_ns_attr_prefix,
            new_nsmap=None,
        )
    if tag:
        write(f"</{tag}>")
    if elem.tail:
        write(ET._escape_cdata(elem.tail))


def _qnames_iter(elem):
    """Iterate through all the qualified names in elem"""
    seen_el_qnames = set()
    seen_other_qnames = set()
    for this_elem in elem.iter():
        tag = this_elem.tag
        if isinstance(tag, str):
            if tag not in seen_el_qnames:
                seen_el_qnames.add(tag)
                yield tag, True
        elif isinstance(tag, ET.QName):
            tag = tag.text
            if tag not in seen_el_qnames:
                seen_el_qnames.add(tag)
                yield tag, True
        elif (
            tag is not None
            and tag is not ET.ProcessingInstruction
            and tag is not ET.Comment
        ):
            ET._raise_serialization_error(tag)

        for key, value in this_elem.items():
            if isinstance(key, ET.QName):
                key = key.text
            if key not in seen_other_qnames:
                seen_other_qnames.add(key)
                yield key, False

            if isinstance(value, ET.QName):
                if value.text not in seen_other_qnames:
                    seen_other_qnames.add(value.text)
                    yield value.text, False

        text = this_elem.text
        if isinstance(text, ET.QName):
            if text.text not in seen_other_qnames:
                seen_other_qnames.add(text.text)
                yield text.text, False


def _namespaces(
    elem,
    default_namespace=None,
    nsmap=None,
):
    """Find all namespaces used in the document and return a prefix to uri map"""
    if nsmap is None:
        nsmap = {}

    out_nsmap = {}

    seen_uri_to_prefix = {}
    # Multiple prefixes may be present for a single uri. This will select the
    # last prefix found in nsmap for a given uri.
    local_prefix_map = {uri: prefix for prefix, uri in nsmap.items()}
    if default_namespace is not None:
        local_prefix_map[default_namespace] = ""
    elif "" in nsmap:
        # but we make sure the default prefix always take precedence
        local_prefix_map[nsmap[""]] = ""

    global_prefixes = set(ET._namespace_map.values())
    has_unqual_el = False
    default_namespace_attr_prefix = None
    for qname, is_el in _qnames_iter(elem):
        try:
            if qname[:1] == "{":
                uri_and_name = qname[1:].rsplit("}", 1)

                prefix = seen_uri_to_prefix.get(uri_and_name[0])
                if prefix is None:
                    prefix = local_prefix_map.get(uri_and_name[0])
                    if prefix is None or prefix in out_nsmap:
                        prefix = ET._namespace_map.get(uri_and_name[0])
                        if prefix is None or prefix in out_nsmap:
                            prefix = _make_new_ns_prefix(
                                out_nsmap,
                                global_prefixes,
                                nsmap,
                                default_namespace,
                            )
                    if prefix or is_el:
                        out_nsmap[prefix] = uri_and_name[0]
                        seen_uri_to_prefix[uri_and_name[0]] = prefix

                if not is_el and not prefix and not default_namespace_attr_prefix:
                    # Find the alternative prefix to use with non-element
                    # names
                    default_namespace_attr_prefix = _find_default_namespace_attr_prefix(
                        uri_and_name[0],
                        out_nsmap,
                        nsmap,
                        global_prefixes,
                        default_namespace,
                    )
                    out_nsmap[default_namespace_attr_prefix] = uri_and_name[0]
                    # Don't add this uri to prefix mapping as it might override
                    # the uri -> "" default mapping. We'll fix this up at the
                    # end of the fn.
                    # local_prefix_map[uri_and_name[0]] = default_namespace_attr_prefix
            else:
                if is_el:
                    has_unqual_el = True
        except TypeError:
            ET._raise_serialization_error(qname)

    if "" in out_nsmap and has_unqual_el:
        # FIXME: can this be handled in XML 1.0?
        raise ValueError(
            "cannot use non-qualified names with default_namespace option"
        )

    # The xml prefix doesn't need to be declared but may have been used to
    # prefix names. Let's remove it if it has been used
    out_nsmap.pop("xml", None)
    return out_nsmap


def tostring(
    element,
    encoding=None,
    method=None,
    *,
    xml_declaration=None,
    default_namespace=None,
    short_empty_elements=True,
    nsmap=None,
    root_ns_only=False,
    minimal_ns_only=False,
    tree_cls=IncrementalTree,
):
    """Generate string representation of XML element.

    All subelements are included.  If encoding is "unicode", a string
    is returned. Otherwise a bytestring is returned.

    *element* is an Element instance, *encoding* is an optional output
    encoding defaulting to US-ASCII, *method* is an optional output which can
    be one of "xml" (default), "html", "text" or "c14n", *default_namespace*
    sets the default XML namespace (for "xmlns").

    Returns an (optionally) encoded string containing the XML data.

    """
    stream = io.StringIO() if encoding == "unicode" else io.BytesIO()
    tree_cls(element).write(
        stream,
        encoding,
        xml_declaration=xml_declaration,
        default_namespace=default_namespace,
        method=method,
        short_empty_elements=short_empty_elements,
        nsmap=nsmap,
        root_ns_only=root_ns_only,
        minimal_ns_only=minimal_ns_only,
    )
    return stream.getvalue()


def tostringlist(
    element,
    encoding=None,
    method=None,
    *,
    xml_declaration=None,
    default_namespace=None,
    short_empty_elements=True,
    nsmap=None,
    root_ns_only=False,
    minimal_ns_only=False,
    tree_cls=IncrementalTree,
):
    lst = []
    stream = ET._ListDataStream(lst)
    tree_cls(element).write(
        stream,
        encoding,
        xml_declaration=xml_declaration,
        default_namespace=default_namespace,
        method=method,
        short_empty_elements=short_empty_elements,
        nsmap=nsmap,
        root_ns_only=root_ns_only,
        minimal_ns_only=minimal_ns_only,
    )
    return lst


def compat_tostring(
    element,
    encoding=None,
    method=None,
    *,
    xml_declaration=None,
    default_namespace=None,
    short_empty_elements=True,
    nsmap=None,
    root_ns_only=True,
    minimal_ns_only=False,
    tree_cls=IncrementalTree,
):
    """tostring with options that produce the same results as xml.etree.ElementTree.tostring

    root_ns_only=True is a bit slower than False as it needs to traverse the
    tree one more time to collect all the namespaces.
    """
    return tostring(
        element,
        encoding=encoding,
        method=method,
        xml_declaration=xml_declaration,
        default_namespace=default_namespace,
        short_empty_elements=short_empty_elements,
        nsmap=nsmap,
        root_ns_only=root_ns_only,
        minimal_ns_only=minimal_ns_only,
        tree_cls=tree_cls,
    )


# --------------------------------------------------------------------
# serialization support

@contextlib.contextmanager
def _get_writer(file_or_filename, encoding):
    # Copied from Python 3.12
    # returns text write method and release all resources after using
    try:
        write = file_or_filename.write
    except AttributeError:
        # file_or_filename is a file name
        if encoding.lower() == "unicode":
            encoding = "utf-8"
        with open(file_or_filename, "w", encoding=encoding,
                  errors="xmlcharrefreplace") as file:
            yield file.write, encoding
    else:
        # file_or_filename is a file-like object
        # encoding determines if it is a text or binary writer
        if encoding.lower() == "unicode":
            # use a text writer as is
            yield write, getattr(file_or_filename, "encoding", None) or "utf-8"
        else:
            # wrap a binary writer with TextIOWrapper
            with contextlib.ExitStack() as stack:
                if isinstance(file_or_filename, io.BufferedIOBase):
                    file = file_or_filename
                elif isinstance(file_or_filename, io.RawIOBase):
                    file = io.BufferedWriter(file_or_filename)
                    # Keep the original file open when the BufferedWriter is
                    # destroyed
                    stack.callback(file.detach)
                else:
                    # This is to handle passed objects that aren't in the
                    # IOBase hierarchy, but just have a write method
                    file = io.BufferedIOBase()
                    file.writable = lambda: True
                    file.write = write
                    try:
                        # TextIOWrapper uses this methods to determine
                        # if BOM (for UTF-16, etc) should be added
                        file.seekable = file_or_filename.seekable
                        file.tell = file_or_filename.tell
                    except AttributeError:
                        pass
                file = io.TextIOWrapper(file,
                                        encoding=encoding,
                                        errors="xmlcharrefreplace",
                                        newline="\n")
                # Keep the original file open when the TextIOWrapper is
                # destroyed
                stack.callback(file.detach)
                yield file.write, encoding

''',
    },
    'et_xmlfile.xmlfile': {
        'is_package': False,
        'source': r'''
from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Implements the lxml.etree.xmlfile API using the standard library xml.etree"""


from contextlib import contextmanager

from xml.etree.ElementTree import (
    Element,
    _escape_cdata,
)

from . import incremental_tree


class LxmlSyntaxError(Exception):
    pass


class _IncrementalFileWriter(object):
    """Replacement for _IncrementalFileWriter of lxml"""
    def __init__(self, output_file):
        self._element_stack = []
        self._file = output_file
        self._have_root = False
        self.global_nsmap = incremental_tree.current_global_nsmap()
        self.is_html = False

    @contextmanager
    def element(self, tag, attrib=None, nsmap=None, **_extra):
        """Create a new xml element using a context manager."""
        if nsmap and None in nsmap:
            # Normalise None prefix (lxml's default namespace prefix) -> "", as
            # required for incremental_tree
            if "" in nsmap and nsmap[""] != nsmap[None]:
                raise ValueError(
                    'Found None and "" as default nsmap prefixes with different URIs'
                )
            nsmap = nsmap.copy()
            nsmap[""] = nsmap.pop(None)

        # __enter__ part
        self._have_root = True
        if attrib is None:
            attrib = {}
        elem = Element(tag, attrib=attrib, **_extra)
        elem.text = ''
        elem.tail = ''
        if self._element_stack:
            is_root = False
            (
                nsmap_scope,
                default_ns_attr_prefix,
                uri_to_prefix,
            ) = self._element_stack[-1]
        else:
            is_root = True
            nsmap_scope = {}
            default_ns_attr_prefix = None
            uri_to_prefix = {}
        (
            tag,
            nsmap_scope,
            default_ns_attr_prefix,
            uri_to_prefix,
            next_remains_root,
        ) = incremental_tree.write_elem_start(
            self._file,
            elem,
            nsmap_scope=nsmap_scope,
            global_nsmap=self.global_nsmap,
            short_empty_elements=False,
            is_html=self.is_html,
            is_root=is_root,
            uri_to_prefix=uri_to_prefix,
            default_ns_attr_prefix=default_ns_attr_prefix,
            new_nsmap=nsmap,
        )
        self._element_stack.append(
            (
                nsmap_scope,
                default_ns_attr_prefix,
                uri_to_prefix,
            )
        )
        yield

        # __exit__ part
        self._element_stack.pop()
        self._file(f"</{tag}>")
        if elem.tail:
            self._file(_escape_cdata(elem.tail))

    def write(self, arg):
        """Write a string or subelement."""

        if isinstance(arg, str):
            # it is not allowed to write a string outside of an element
            if not self._element_stack:
                raise LxmlSyntaxError()
            self._file(_escape_cdata(arg))

        else:
            if not self._element_stack and self._have_root:
                raise LxmlSyntaxError()

            if self._element_stack:
                is_root = False
                (
                    nsmap_scope,
                    default_ns_attr_prefix,
                    uri_to_prefix,
                ) = self._element_stack[-1]
            else:
                is_root = True
                nsmap_scope = {}
                default_ns_attr_prefix = None
                uri_to_prefix = {}
            incremental_tree._serialize_ns_xml(
                self._file,
                arg,
                nsmap_scope=nsmap_scope,
                global_nsmap=self.global_nsmap,
                short_empty_elements=True,
                is_html=self.is_html,
                is_root=is_root,
                uri_to_prefix=uri_to_prefix,
                default_ns_attr_prefix=default_ns_attr_prefix,
            )

    def __enter__(self):
        pass

    def __exit__(self, type, value, traceback):
        # without root the xml document is incomplete
        if not self._have_root:
            raise LxmlSyntaxError()


class xmlfile(object):
    """Context manager that can replace lxml.etree.xmlfile."""
    def __init__(self, output_file, buffered=False, encoding="utf-8", close=False):
        self._file = output_file
        self._close = close
        self.encoding = encoding
        self.writer_cm = None

    def __enter__(self):
        self.writer_cm = incremental_tree._get_writer(self._file, encoding=self.encoding)
        writer, declared_encoding = self.writer_cm.__enter__()
        return _IncrementalFileWriter(writer)

    def __exit__(self, type, value, traceback):
        if self.writer_cm:
            self.writer_cm.__exit__(type, value, traceback)
        if self._close:
            self._file.close()

''',
    },
    'openpyxl': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

DEBUG = False

from openpyxl.compat.numbers import NUMPY
from openpyxl.xml import DEFUSEDXML, LXML
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook as open
from openpyxl.reader.excel import load_workbook
import openpyxl._constants as constants

# Expose constants especially the version number

__author__ = constants.__author__
__author_email__ = constants.__author_email__
__license__ = constants.__license__
__maintainer_email__ = constants.__maintainer_email__
__url__ = constants.__url__
__version__ = constants.__version__

""",
    },
    'openpyxl._constants': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Package metadata
"""

__author__ = "See AUTHORS"
__author_email__ = "charlie.clark@clark-consulting.eu"
__license__ = "MIT"
__maintainer_email__ = "openpyxl-users@googlegroups.com"
__url__ = "https://openpyxl.readthedocs.io"
__version__ = "3.1.5"
__python__ = "3.8"

''',
    },
    'openpyxl.cell': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .cell import Cell, WriteOnlyCell, MergedCell
from .read_only import ReadOnlyCell

""",
    },
    'openpyxl.cell._writer': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement, whitespace, XML_NS
from openpyxl import LXML
from openpyxl.utils.datetime import to_excel, to_ISO8601
from datetime import timedelta

from openpyxl.worksheet.formula import DataTableFormula, ArrayFormula
from openpyxl.cell.rich_text import CellRichText

def _set_attributes(cell, styled=None):
    """
    Set coordinate and datatype
    """
    coordinate = cell.coordinate
    attrs = {'r': coordinate}
    if styled:
        attrs['s'] = f"{cell.style_id}"

    if cell.data_type == "s":
        attrs['t'] = "inlineStr"
    elif cell.data_type != 'f':
        attrs['t'] = cell.data_type

    value = cell._value

    if cell.data_type == "d":
        if hasattr(value, "tzinfo") and value.tzinfo is not None:
            raise TypeError("Excel does not support timezones in datetimes. "
                    "The tzinfo in the datetime/time object must be set to None.")

        if cell.parent.parent.iso_dates and not isinstance(value, timedelta):
            value = to_ISO8601(value)
        else:
            attrs['t'] = "n"
            value = to_excel(value, cell.parent.parent.epoch)

    if cell.hyperlink:
        cell.parent._hyperlinks.append(cell.hyperlink)

    return value, attrs


def etree_write_cell(xf, worksheet, cell, styled=None):

    value, attributes = _set_attributes(cell, styled)

    el = Element("c", attributes)
    if value is None or value == "":
        xf.write(el)
        return

    if cell.data_type == 'f':
        attrib = {}

        if isinstance(value, ArrayFormula):
            attrib = dict(value)
            value = value.text

        elif isinstance(value, DataTableFormula):
            attrib = dict(value)
            value = None

        formula = SubElement(el, 'f', attrib)
        if value is not None and not attrib.get('t') == "dataTable":
            formula.text = value[1:]
            value = None

    if cell.data_type == 's':
        if isinstance(value, CellRichText):
            el.append(value.to_tree())
        else:
            inline_string = Element("is")
            text = Element('t')
            text.text = value
            whitespace(text)
            inline_string.append(text)
            el.append(inline_string)

    else:
        cell_content = SubElement(el, 'v')
        if value is not None:
            cell_content.text = safe_string(value)

    xf.write(el)


def lxml_write_cell(xf, worksheet, cell, styled=False):
    value, attributes = _set_attributes(cell, styled)

    if value == '' or value is None:
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            attrib = {}

            if isinstance(value, ArrayFormula):
                attrib = dict(value)
                value = value.text

            elif isinstance(value, DataTableFormula):
                attrib = dict(value)
                value = None

            with xf.element('f', attrib):
                if value is not None and not attrib.get('t') == "dataTable":
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            if isinstance(value, CellRichText):
                el = value.to_tree()
                xf.write(el)
            else:
                with xf.element("is"):
                    if isinstance(value, str):
                        attrs = {}
                        if value != value.strip():
                            attrs["{%s}space" % XML_NS] = "preserve"
                        el = Element("t", attrs) # lxml can't handle xml-ns
                        el.text = value
                        xf.write(el)

        else:
            with xf.element("v"):
                if value is not None:
                    xf.write(safe_string(value))


if LXML:
    write_cell = lxml_write_cell
else:
    write_cell = etree_write_cell

''',
    },
    'openpyxl.cell.cell': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Manage individual cells in a spreadsheet.

The Cell class is required to know its value and type, display options,
and any other features of an Excel cell.  Utilities for referencing
cells using Excel's 'A1' column/row nomenclature are also provided.

"""

__docformat__ = "restructuredtext en"

# Python stdlib imports
from copy import copy
import datetime
import re


from openpyxl.compat import (
    NUMERIC_TYPES,
)

from openpyxl.utils.exceptions import IllegalCharacterError

from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers, is_date_format
from openpyxl.styles.styleable import StyleableObject
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.worksheet.formula import DataTableFormula, ArrayFormula
from openpyxl.cell.rich_text import CellRichText

# constants

TIME_TYPES = (datetime.datetime, datetime.date, datetime.time, datetime.timedelta)
TIME_FORMATS = {
    datetime.datetime:numbers.FORMAT_DATE_DATETIME,
    datetime.date:numbers.FORMAT_DATE_YYYYMMDD2,
    datetime.time:numbers.FORMAT_DATE_TIME6,
    datetime.timedelta:numbers.FORMAT_DATE_TIMEDELTA,
                }

STRING_TYPES = (str, bytes, CellRichText)
KNOWN_TYPES = NUMERIC_TYPES + TIME_TYPES + STRING_TYPES + (bool, type(None))

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')
ERROR_CODES = ('#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!',
               '#N/A')

TYPE_STRING = 's'
TYPE_FORMULA = 'f'
TYPE_NUMERIC = 'n'
TYPE_BOOL = 'b'
TYPE_NULL = 'n'
TYPE_INLINE = 'inlineStr'
TYPE_ERROR = 'e'
TYPE_FORMULA_CACHE_STRING = 'str'

VALID_TYPES = (TYPE_STRING, TYPE_FORMULA, TYPE_NUMERIC, TYPE_BOOL,
               TYPE_NULL, TYPE_INLINE, TYPE_ERROR, TYPE_FORMULA_CACHE_STRING)


_TYPES = {int:'n', float:'n', str:'s', bool:'b'}


def get_type(t, value):
    if isinstance(value, NUMERIC_TYPES):
        dt = 'n'
    elif isinstance(value, STRING_TYPES):
        dt = 's'
    elif isinstance(value, TIME_TYPES):
        dt = 'd'
    elif isinstance(value, (DataTableFormula, ArrayFormula)):
        dt = 'f'
    else:
        return
    _TYPES[t] = dt
    return dt


def get_time_format(t):
    value = TIME_FORMATS.get(t)
    if value:
        return value
    for base in t.mro()[1:]:
        value = TIME_FORMATS.get(base)
        if value:
            TIME_FORMATS[t] = value
            return value
    raise ValueError("Could not get time format for {0!r}".format(value))


class Cell(StyleableObject):
    """Describes cell associated properties.

    Properties of interest include style, type, value, and address.

    """
    __slots__ = (
        'row',
        'column',
        '_value',
        'data_type',
        'parent',
        '_hyperlink',
        '_comment',
                 )

    def __init__(self, worksheet, row=None, column=None, value=None, style_array=None):
        super().__init__(worksheet, style_array)
        self.row = row
        """Row number of this cell (1-based)"""
        self.column = column
        """Column number of this cell (1-based)"""
        # _value is the stored value, while value is the displayed value
        self._value = None
        self._hyperlink = None
        self.data_type = 'n'
        if value is not None:
            self.value = value
        self._comment = None


    @property
    def coordinate(self):
        """This cell's coordinate (ex. 'A5')"""
        col = get_column_letter(self.column)
        return f"{col}{self.row}"


    @property
    def col_idx(self):
        """The numerical index of the column"""
        return self.column


    @property
    def column_letter(self):
        return get_column_letter(self.column)


    @property
    def encoding(self):
        return self.parent.encoding

    @property
    def base_date(self):
        return self.parent.parent.epoch


    def __repr__(self):
        return "<Cell {0!r}.{1}>".format(self.parent.title, self.coordinate)

    def check_string(self, value):
        """Check string coding, length, and line break character"""
        if value is None:
            return
        # convert to str string
        if not isinstance(value, str):
            value = str(value, self.encoding)
        value = str(value)
        # string must never be longer than 32,767 characters
        # truncate if necessary
        value = value[:32767]
        if next(ILLEGAL_CHARACTERS_RE.finditer(value), None):
            raise IllegalCharacterError(f"{value} cannot be used in worksheets.")
        return value

    def check_error(self, value):
        """Tries to convert Error" else N/A"""
        try:
            return str(value)
        except UnicodeDecodeError:
            return u'#N/A'


    def _bind_value(self, value):
        """Given a value, infer the correct data type"""

        self.data_type = "n"
        t = type(value)
        try:
            dt = _TYPES[t]
        except KeyError:
            dt = get_type(t, value)

        if dt is None and value is not None:
            raise ValueError("Cannot convert {0!r} to Excel".format(value))

        if dt:
            self.data_type = dt

        if dt == 'd':
            if not is_date_format(self.number_format):
                self.number_format = get_time_format(t)

        elif dt == "s" and not isinstance(value, CellRichText):
            value = self.check_string(value)
            if len(value) > 1 and value.startswith("="):
                self.data_type = 'f'
            elif value in ERROR_CODES:
                self.data_type = 'e'

        self._value = value


    @property
    def value(self):
        """Get or set the value held in the cell.

        :type: depends on the value (string, float, int or
            :class:`datetime.datetime`)
        """
        return self._value

    @value.setter
    def value(self, value):
        """Set the value and infer type and display options."""
        self._bind_value(value)

    @property
    def internal_value(self):
        """Always returns the value for excel."""
        return self._value

    @property
    def hyperlink(self):
        """Return the hyperlink target or an empty string"""
        return self._hyperlink


    @hyperlink.setter
    def hyperlink(self, val):
        """Set value and display for hyperlinks in a cell.
        Automatically sets the `value` of the cell with link text,
        but you can modify it afterwards by setting the `value`
        property, and the hyperlink will remain.
        Hyperlink is removed if set to ``None``."""
        if val is None:
            self._hyperlink = None
        else:
            if not isinstance(val, Hyperlink):
                val = Hyperlink(ref="", target=val)
            val.ref = self.coordinate
            self._hyperlink = val
            if self._value is None:
                self.value = val.target or val.location


    @property
    def is_date(self):
        """True if the value is formatted as a date

        :type: bool
        """
        return self.data_type == 'd' or (
            self.data_type == 'n' and is_date_format(self.number_format)
            )


    def offset(self, row=0, column=0):
        """Returns a cell location relative to this cell.

        :param row: number of rows to offset
        :type row: int

        :param column: number of columns to offset
        :type column: int

        :rtype: :class:`openpyxl.cell.Cell`
        """
        offset_column = self.col_idx + column
        offset_row = self.row + row
        return self.parent.cell(column=offset_column, row=offset_row)


    @property
    def comment(self):
        """ Returns the comment associated with this cell

            :type: :class:`openpyxl.comments.Comment`
        """
        return self._comment


    @comment.setter
    def comment(self, value):
        """
        Assign a comment to a cell
        """

        if value is not None:
            if value.parent:
                value = copy(value)
            value.bind(self)
        elif value is None and self._comment:
            self._comment.unbind()
        self._comment = value


class MergedCell(StyleableObject):

    """
    Describes the properties of a cell in a merged cell and helps to
    display the borders of the merged cell.

    The value of a MergedCell is always None.
    """

    __slots__ = ('row', 'column')

    _value = None
    data_type = "n"
    comment = None
    hyperlink = None


    def __init__(self, worksheet, row=None, column=None):
        super().__init__(worksheet)
        self.row = row
        self.column = column


    def __repr__(self):
        return "<MergedCell {0!r}.{1}>".format(self.parent.title, self.coordinate)

    coordinate = Cell.coordinate
    _comment = comment
    value = _value


def WriteOnlyCell(ws=None, value=None):
    return Cell(worksheet=ws, column=1, row=1, value=value)

''',
    },
    'openpyxl.cell.read_only': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.styles import is_date_format
from openpyxl.styles.numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_MAX_SIZE


class ReadOnlyCell:

    __slots__ =  ('parent', 'row', 'column', '_value', 'data_type', '_style_id')

    def __init__(self, sheet, row, column, value, data_type='n', style_id=0):
        self.parent = sheet
        self._value = None
        self.row = row
        self.column = column
        self.data_type = data_type
        self.value = value
        self._style_id = style_id


    def __eq__(self, other):
        for a in self.__slots__:
            if getattr(self, a) != getattr(other, a):
                return
        return True

    def __ne__(self, other):
        return not self.__eq__(other)


    def __repr__(self):
        return "<ReadOnlyCell {0!r}.{1}>".format(self.parent.title, self.coordinate)


    @property
    def coordinate(self):
        column = get_column_letter(self.column)
        return "{1}{0}".format(self.row, column)


    @property
    def coordinate(self):
        return Cell.coordinate.__get__(self)


    @property
    def column_letter(self):
        return Cell.column_letter.__get__(self)


    @property
    def style_array(self):
        return self.parent.parent._cell_styles[self._style_id]


    @property
    def has_style(self):
        return self._style_id != 0


    @property
    def number_format(self):
        _id = self.style_array.numFmtId
        if _id < BUILTIN_FORMATS_MAX_SIZE:
            return BUILTIN_FORMATS.get(_id, "General")
        else:
            return self.parent.parent._number_formats[
                _id - BUILTIN_FORMATS_MAX_SIZE]

    @property
    def font(self):
        _id = self.style_array.fontId
        return self.parent.parent._fonts[_id]

    @property
    def fill(self):
        _id = self.style_array.fillId
        return self.parent.parent._fills[_id]

    @property
    def border(self):
        _id = self.style_array.borderId
        return self.parent.parent._borders[_id]

    @property
    def alignment(self):
        _id = self.style_array.alignmentId
        return self.parent.parent._alignments[_id]

    @property
    def protection(self):
        _id = self.style_array.protectionId
        return self.parent.parent._protections[_id]


    @property
    def is_date(self):
        return Cell.is_date.__get__(self)


    @property
    def internal_value(self):
        return self._value

    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        if self._value is not None:
            raise AttributeError("Cell is read only")
        self._value = value


class EmptyCell:

    __slots__ = ()

    value = None
    is_date = False
    font = None
    border = None
    fill = None
    number_format = None
    alignment = None
    data_type = 'n'


    def __repr__(self):
        return "<EmptyCell>"

EMPTY_CELL = EmptyCell()

""",
    },
    'openpyxl.cell.rich_text': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
RichText definition
"""
from copy import copy
from openpyxl.compat import NUMERIC_TYPES
from openpyxl.cell.text import InlineFont, Text
from openpyxl.descriptors import (
    Strict,
    String,
    Typed
)

from openpyxl.xml.functions import Element, whitespace

class TextBlock(Strict):
    """ Represents text string in a specific format

    This class is used as part of constructing a rich text strings.
    """
    font = Typed(expected_type=InlineFont)
    text = String()

    def __init__(self, font, text):
        self.font = font
        self.text = text


    def __eq__(self, other):
        return self.text == other.text and self.font == other.font


    def __str__(self):
        """Just retun the text"""
        return self.text


    def __repr__(self):
        font = self.font != InlineFont() and self.font or "default"
        return f"{self.__class__.__name__} text={self.text}, font={font}"


    def to_tree(self):
        el = Element("r")
        el.append(self.font.to_tree(tagname="rPr"))
        t = Element("t")
        t.text = self.text
        whitespace(t)
        el.append(t)
        return el

#
# Rich Text class.
# This class behaves just like a list whose members are either simple strings, or TextBlock() instances.
# In addition, it can be initialized in several ways:
# t = CellRFichText([...]) # initialize with a list.
# t = CellRFichText((...)) # initialize with a tuple.
# t = CellRichText(node) # where node is an Element() from either lxml or xml.etree (has a 'tag' element)
class CellRichText(list):
    """Represents a rich text string.

    Initialize with a list made of pure strings or :class:`TextBlock` elements
    Can index object to access or modify individual rich text elements
    it also supports the + and += operators between rich text strings
    There are no user methods for this class

    operations which modify the string will generally call an optimization pass afterwards,
    that merges text blocks with identical formats, consecutive pure text strings,
    and remove empty strings and empty text blocks
    """

    def __init__(self, *args):
        if len(args) == 1:
            args = args[0]
            if isinstance(args, (list, tuple)):
                CellRichText._check_rich_text(args)
            else:
                CellRichText._check_element(args)
                args = [args]
        else:
            CellRichText._check_rich_text(args)
        super().__init__(args)


    @classmethod
    def _check_element(cls, value):
        if not isinstance(value, (str, TextBlock, NUMERIC_TYPES)):
            raise TypeError(f"Illegal CellRichText element {value}")


    @classmethod
    def _check_rich_text(cls, rich_text):
        for t in rich_text:
            CellRichText._check_element(t)

    @classmethod
    def from_tree(cls, node):
        text = Text.from_tree(node)
        if text.t:
            return (text.t.replace('x005F_', ''),)
        s = []
        for r in text.r:
            t = ""
            if r.t:
                t = r.t.replace('x005F_', '')
            if r.rPr:
                s.append(TextBlock(r.rPr, t))
            else:
                s.append(t)
        return cls(s)

    # Merge TextBlocks with identical formatting
    # remove empty elements
    def _opt(self):
        last_t = None
        l = CellRichText(tuple())
        for t in self:
            if isinstance(t, str):
                if not t:
                    continue
            elif not t.text:
                continue
            if type(last_t) == type(t):
                if isinstance(t, str):
                    last_t += t
                    continue
                elif last_t.font == t.font:
                    last_t.text += t.text
                    continue
            if last_t:
                l.append(last_t)
            last_t = t
        if last_t:
            # Add remaining TextBlock at end of rich text
            l.append(last_t)
        super().__setitem__(slice(None), l)
        return self


    def __iadd__(self, arg):
        # copy used here to create new TextBlock() so we don't modify the right hand side in _opt()
        CellRichText._check_rich_text(arg)
        super().__iadd__([copy(e) for e in list(arg)])
        return self._opt()


    def __add__(self, arg):
        return CellRichText([copy(e) for e in list(self) + list(arg)])._opt()


    def __setitem__(self, indx, val):
        CellRichText._check_element(val)
        super().__setitem__(indx, val)
        self._opt()


    def append(self, arg):
        CellRichText._check_element(arg)
        super().append(arg)


    def extend(self, arg):
        CellRichText._check_rich_text(arg)
        super().extend(arg)


    def __repr__(self):
        return "CellRichText([{}])".format(', '.join((repr(s) for s in self)))


    def __str__(self):
        return ''.join([str(s) for s in self])


    def as_list(self):
        """
        Returns a list of the strings contained.
        The main reason for this is to make editing easier.
        """
        return [str(s) for s in self]


    def to_tree(self):
        """
        Return the full XML representation
        """
        container = Element("is")
        for obj in self:
            if isinstance(obj, TextBlock):
                container.append(obj.to_tree())

            else:
                el = Element("r")
                t = Element("t")
                t.text = obj
                whitespace(t)
                el.append(t)
                container.append(el)

        return container


''',
    },
    'openpyxl.cell.text': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Richtext definition
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    Integer,
    Set,
    NoneSet,
    Bool,
    String,
    Sequence,
)
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedString,
    NestedText,
)
from openpyxl.styles.fonts import Font


class PhoneticProperties(Serialisable):

    tagname = "phoneticPr"

    fontId = Integer()
    type = NoneSet(values=(['halfwidthKatakana', 'fullwidthKatakana',
                            'Hiragana', 'noConversion']))
    alignment = NoneSet(values=(['noControl', 'left', 'center', 'distributed']))

    def __init__(self,
                 fontId=None,
                 type=None,
                 alignment=None,
                ):
        self.fontId = fontId
        self.type = type
        self.alignment = alignment


class PhoneticText(Serialisable):

    tagname = "rPh"

    sb = Integer()
    eb = Integer()
    t = NestedText(expected_type=str)
    text = Alias('t')

    def __init__(self,
                 sb=None,
                 eb=None,
                 t=None,
                ):
        self.sb = sb
        self.eb = eb
        self.t = t


class InlineFont(Font):

    """
    Font for inline text because, yes what you need are different objects with the same elements but different constraints.
    """

    tagname = "RPrElt"

    rFont = NestedString(allow_none=True)
    charset = Font.charset
    family = Font.family
    b =Font.b
    i = Font.i
    strike = Font.strike
    outline = Font.outline
    shadow = Font.shadow
    condense = Font.condense
    extend = Font.extend
    color = Font.color
    sz = Font.sz
    u = Font.u
    vertAlign = Font.vertAlign
    scheme = Font.scheme

    __elements__ = ('rFont', 'charset', 'family', 'b', 'i', 'strike',
                    'outline', 'shadow', 'condense', 'extend', 'color', 'sz', 'u',
                    'vertAlign', 'scheme')

    def __init__(self,
                 rFont=None,
                 charset=None,
                 family=None,
                 b=None,
                 i=None,
                 strike=None,
                 outline=None,
                 shadow=None,
                 condense=None,
                 extend=None,
                 color=None,
                 sz=None,
                 u=None,
                 vertAlign=None,
                 scheme=None,
                ):
        self.rFont = rFont
        self.charset = charset
        self.family = family
        self.b = b
        self.i = i
        self.strike = strike
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.color = color
        self.sz = sz
        self.u = u
        self.vertAlign = vertAlign
        self.scheme = scheme


class RichText(Serialisable):

    tagname = "RElt"

    rPr = Typed(expected_type=InlineFont, allow_none=True)
    font = Alias("rPr")
    t = NestedText(expected_type=str, allow_none=True)
    text = Alias("t")

    __elements__ = ('rPr', 't')

    def __init__(self,
                 rPr=None,
                 t=None,
                ):
        self.rPr = rPr
        self.t = t


class Text(Serialisable):

    tagname = "text"

    t = NestedText(allow_none=True, expected_type=str)
    plain = Alias("t")
    r = Sequence(expected_type=RichText, allow_none=True)
    formatted = Alias("r")
    rPh = Sequence(expected_type=PhoneticText, allow_none=True)
    phonetic = Alias("rPh")
    phoneticPr = Typed(expected_type=PhoneticProperties, allow_none=True)
    PhoneticProperties = Alias("phoneticPr")

    __elements__ = ('t', 'r', 'rPh', 'phoneticPr')

    def __init__(self,
                 t=None,
                 r=(),
                 rPh=(),
                 phoneticPr=None,
                ):
        self.t = t
        self.r = r
        self.rPh = rPh
        self.phoneticPr = phoneticPr


    @property
    def content(self):
        """
        Text stripped of all formatting
        """
        snippets = []
        if self.plain is not None:
            snippets.append(self.plain)
        for block in self.formatted:
            if block.t is not None:
                snippets.append(block.t)
        return u"".join(snippets)

''',
    },
    'openpyxl.chart': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import (
    PieChart,
    PieChart3D,
    DoughnutChart,
    ProjectedPieChart
)
from .radar_chart import RadarChart
from .scatter_chart import ScatterChart
from .stock_chart import StockChart
from .surface_chart import SurfaceChart, SurfaceChart3D

from .series_factory import SeriesFactory as Series
from .reference import Reference

""",
    },
    'openpyxl.chart._3d': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import Typed, Alias
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedMinMax,
)
from openpyxl.descriptors.excel import ExtensionList
from .marker import PictureOptions
from .shapes import GraphicalProperties


class View3D(Serialisable):

    tagname = "view3D"

    rotX = NestedMinMax(min=-90, max=90, allow_none=True)
    x_rotation = Alias('rotX')
    hPercent = NestedMinMax(min=5, max=500, allow_none=True)
    height_percent = Alias('hPercent')
    rotY = NestedInteger(min=-90, max=90, allow_none=True)
    y_rotation = Alias('rotY')
    depthPercent = NestedInteger(allow_none=True)
    rAngAx = NestedBool(allow_none=True)
    right_angle_axes = Alias('rAngAx')
    perspective = NestedInteger(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('rotX', 'hPercent', 'rotY', 'depthPercent', 'rAngAx',
                    'perspective',)

    def __init__(self,
                 rotX=15,
                 hPercent=None,
                 rotY=20,
                 depthPercent=None,
                 rAngAx=True,
                 perspective=None,
                 extLst=None,
                ):
        self.rotX = rotX
        self.hPercent = hPercent
        self.rotY = rotY
        self.depthPercent = depthPercent
        self.rAngAx = rAngAx
        self.perspective = perspective


class Surface(Serialisable):

    tagname = "surface"

    thickness = NestedInteger(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('thickness', 'spPr', 'pictureOptions',)

    def __init__(self,
                 thickness=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.thickness = thickness
        self.spPr = spPr
        self.pictureOptions = pictureOptions


class _3DBase(Serialisable):

    """
    Base class for 3D charts
    """

    tagname = "ChartBase"

    view3D = Typed(expected_type=View3D, allow_none=True)
    floor = Typed(expected_type=Surface, allow_none=True)
    sideWall = Typed(expected_type=Surface, allow_none=True)
    backWall = Typed(expected_type=Surface, allow_none=True)

    def __init__(self,
                 view3D=None,
                 floor=None,
                 sideWall=None,
                 backWall=None,
                 ):
        if view3D is None:
            view3D = View3D()
        self.view3D = view3D
        if floor is None:
            floor = Surface()
        self.floor = floor
        if sideWall is None:
            sideWall = Surface()
        self.sideWall = sideWall
        if backWall is None:
            backWall = Surface()
        self.backWall = backWall
        super(_3DBase, self).__init__()

''',
    },
    'openpyxl.chart._chart': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from collections import OrderedDict
from operator import attrgetter

from openpyxl.descriptors import (
    Typed,
    Integer,
    Alias,
    MinMax,
    Bool,
    Set,
)
from openpyxl.descriptors.sequence import ValueSequence
from openpyxl.descriptors.serialisable import Serialisable

from ._3d import _3DBase
from .data_source import AxDataSource, NumRef
from .layout import Layout
from .legend import Legend
from .reference import Reference
from .series_factory import SeriesFactory
from .series import attribute_mapping
from .shapes import GraphicalProperties
from .title import TitleDescriptor

class AxId(Serialisable):

    val = Integer()

    def __init__(self, val):
        self.val = val


def PlotArea():
    from .chartspace import PlotArea
    return PlotArea()


class ChartBase(Serialisable):

    """
    Base class for all charts
    """

    legend = Typed(expected_type=Legend, allow_none=True)
    layout = Typed(expected_type=Layout, allow_none=True)
    roundedCorners = Bool(allow_none=True)
    axId = ValueSequence(expected_type=int)
    visible_cells_only = Bool(allow_none=True)
    display_blanks = Set(values=['span', 'gap', 'zero'])
    graphical_properties = Typed(expected_type=GraphicalProperties, allow_none=True)

    _series_type = ""
    ser = ()
    series = Alias('ser')
    title = TitleDescriptor()
    anchor = "E15" # default anchor position
    width = 15 # in cm, approx 5 rows
    height = 7.5 # in cm, approx 14 rows
    _id = 1
    _path = "/xl/charts/chart{0}.xml"
    style = MinMax(allow_none=True, min=1, max=48)
    mime_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    graphical_properties = Typed(expected_type=GraphicalProperties, allow_none=True) # mapped to chartspace

    __elements__ = ()


    def __init__(self, axId=(), **kw):
        self._charts = [self]
        self.title = None
        self.layout = None
        self.roundedCorners = None
        self.legend = Legend()
        self.graphical_properties = None
        self.style = None
        self.plot_area = PlotArea()
        self.axId = axId
        self.display_blanks = 'gap'
        self.pivotSource = None
        self.pivotFormats = ()
        self.visible_cells_only = True
        self.idx_base = 0
        self.graphical_properties = None
        super().__init__()


    def __hash__(self):
        """
        Just need to check for identity
        """
        return id(self)

    def __iadd__(self, other):
        """
        Combine the chart with another one
        """
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self


    def to_tree(self, namespace=None, tagname=None, idx=None):
        self.axId = [id for id in self._axes]
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        return super().to_tree(tagname, idx)


    def _reindex(self):
        """
        Normalise and rebase series: sort by order and then rebase order

        """
        # sort data series in order and rebase
        ds = sorted(self.series, key=attrgetter("order"))
        for idx, s in enumerate(ds):
            s.order = idx
        self.series = ds


    def _write(self):
        from .chartspace import ChartSpace, ChartContainer
        self.plot_area.layout = self.layout

        idx_base = self.idx_base
        for chart in self._charts:
            if chart not in self.plot_area._charts:
                chart.idx_base = idx_base
                idx_base += len(chart.series)
        self.plot_area._charts = self._charts

        container = ChartContainer(plotArea=self.plot_area, legend=self.legend, title=self.title)
        if isinstance(chart, _3DBase):
            container.view3D = chart.view3D
            container.floor = chart.floor
            container.sideWall = chart.sideWall
            container.backWall = chart.backWall
        container.plotVisOnly = self.visible_cells_only
        container.dispBlanksAs = self.display_blanks
        container.pivotFmts = self.pivotFormats
        cs = ChartSpace(chart=container)
        cs.style = self.style
        cs.roundedCorners = self.roundedCorners
        cs.pivotSource = self.pivotSource
        cs.spPr = self.graphical_properties
        return cs.to_tree()


    @property
    def _axes(self):
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        return OrderedDict([(axis.axId, axis) for axis in (x, y, z) if axis])


    def set_categories(self, labels):
        """
        Set the categories / x-axis values
        """
        if not isinstance(labels, Reference):
            labels = Reference(range_string=labels)
        for s in self.ser:
            s.cat = AxDataSource(numRef=NumRef(f=labels))


    def add_data(self, data, from_rows=False, titles_from_data=False):
        """
        Add a range of data in a single pass.
        The default is to treat each column as a data series.
        """
        if not isinstance(data, Reference):
            data = Reference(range_string=data)

        if from_rows:
            values = data.rows

        else:
            values = data.cols

        for ref in values:
            series = SeriesFactory(ref, title_from_data=titles_from_data)
            self.series.append(series)


    def append(self, value):
        """Append a data series to the chart"""
        l = self.series[:]
        l.append(value)
        self.series = l


    @property
    def path(self):
        return self._path.format(self._id)

''',
    },
    'openpyxl.chart.area_chart': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Set,
    Bool,
    Integer,
    Sequence,
    Alias,
)

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedMinMax,
    NestedSet,
    NestedBool,
)

from ._chart import ChartBase
from .descriptors import NestedGapAmount
from .axis import TextAxis, NumericAxis, SeriesAxis, ChartLines
from .label import DataLabelList
from .series import Series


class _AreaChartBase(ChartBase):

    grouping = NestedSet(values=(['percentStacked', 'standard', 'stacked']))
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    dropLines = Typed(expected_type=ChartLines, allow_none=True)

    _series_type = "area"

    __elements__ = ('grouping', 'varyColors', 'ser', 'dLbls', 'dropLines')

    def __init__(self,
                 grouping="standard",
                 varyColors=None,
                 ser=(),
                 dLbls=None,
                 dropLines=None,
                ):
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        super().__init__()


class AreaChart(_AreaChartBase):

    tagname = "areaChart"

    grouping = _AreaChartBase.grouping
    varyColors = _AreaChartBase.varyColors
    ser = _AreaChartBase.ser
    dLbls = _AreaChartBase.dLbls
    dropLines = _AreaChartBase.dropLines

    # chart properties actually used by containing classes
    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)

    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _AreaChartBase.__elements__ + ('axId',)

    def __init__(self,
                 axId=None,
                 extLst=None,
                 **kw
                ):
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        super().__init__(**kw)


class AreaChart3D(AreaChart):

    tagname = "area3DChart"

    grouping = _AreaChartBase.grouping
    varyColors = _AreaChartBase.varyColors
    ser = _AreaChartBase.ser
    dLbls = _AreaChartBase.dLbls
    dropLines = _AreaChartBase.dropLines

    gapDepth = NestedGapAmount()

    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)
    z_axis = Typed(expected_type=SeriesAxis, allow_none=True)

    __elements__ = AreaChart.__elements__ + ('gapDepth', )

    def __init__(self, gapDepth=None, **kw):
        self.gapDepth = gapDepth
        super(AreaChart3D, self).__init__(**kw)
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()

""",
    },
    'openpyxl.chart.axis': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    NoneSet,
    Bool,
    Integer,
    MinMax,
    NoneSet,
    Set,
    String,
    Alias,
)

from openpyxl.descriptors.excel import (
    ExtensionList,
    Percentage,
    _explicit_none,
)
from openpyxl.descriptors.nested import (
    NestedValue,
    NestedSet,
    NestedBool,
    NestedNoneSet,
    NestedFloat,
    NestedInteger,
    NestedMinMax,
)
from openpyxl.xml.constants import CHART_NS

from .descriptors import NumberFormatDescriptor
from .layout import Layout
from .text import Text, RichText
from .shapes import GraphicalProperties
from .title import Title, TitleDescriptor


class ChartLines(Serialisable):

    tagname = "chartLines"

    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')

    def __init__(self, spPr=None):
        self.spPr = spPr


class Scaling(Serialisable):

    tagname = "scaling"

    logBase = NestedFloat(allow_none=True)
    orientation = NestedSet(values=(['maxMin', 'minMax']))
    max = NestedFloat(allow_none=True)
    min = NestedFloat(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('logBase', 'orientation', 'max', 'min',)

    def __init__(self,
                 logBase=None,
                 orientation="minMax",
                 max=None,
                 min=None,
                 extLst=None,
                ):
        self.logBase = logBase
        self.orientation = orientation
        self.max = max
        self.min = min


class _BaseAxis(Serialisable):

    axId = NestedInteger(expected_type=int)
    scaling = Typed(expected_type=Scaling)
    delete = NestedBool(allow_none=True)
    axPos = NestedSet(values=(['b', 'l', 'r', 't']))
    majorGridlines = Typed(expected_type=ChartLines, allow_none=True)
    minorGridlines = Typed(expected_type=ChartLines, allow_none=True)
    title = TitleDescriptor()
    numFmt = NumberFormatDescriptor()
    number_format = Alias("numFmt")
    majorTickMark = NestedNoneSet(values=(['cross', 'in', 'out']), to_tree=_explicit_none)
    minorTickMark = NestedNoneSet(values=(['cross', 'in', 'out']), to_tree=_explicit_none)
    tickLblPos = NestedNoneSet(values=(['high', 'low', 'nextTo']))
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    txPr = Typed(expected_type=RichText, allow_none=True)
    textProperties = Alias('txPr')
    crossAx = NestedInteger(expected_type=int) # references other axis
    crosses = NestedNoneSet(values=(['autoZero', 'max', 'min']))
    crossesAt = NestedFloat(allow_none=True)

    # crosses & crossesAt are mutually exclusive

    __elements__ = ('axId', 'scaling', 'delete', 'axPos', 'majorGridlines',
                    'minorGridlines', 'title', 'numFmt', 'majorTickMark', 'minorTickMark',
                    'tickLblPos', 'spPr', 'txPr', 'crossAx', 'crosses', 'crossesAt')

    def __init__(self,
                 axId=None,
                 scaling=None,
                 delete=None,
                 axPos='l',
                 majorGridlines=None,
                 minorGridlines=None,
                 title=None,
                 numFmt=None,
                 majorTickMark=None,
                 minorTickMark=None,
                 tickLblPos=None,
                 spPr=None,
                 txPr= None,
                 crossAx=None,
                 crosses=None,
                 crossesAt=None,
                ):
        self.axId = axId
        if scaling is None:
            scaling = Scaling()
        self.scaling = scaling
        self.delete = delete
        self.axPos = axPos
        self.majorGridlines = majorGridlines
        self.minorGridlines = minorGridlines
        self.title = title
        self.numFmt = numFmt
        self.majorTickMark = majorTickMark
        self.minorTickMark = minorTickMark
        self.tickLblPos = tickLblPos
        self.spPr = spPr
        self.txPr = txPr
        self.crossAx = crossAx
        self.crosses = crosses
        self.crossesAt = crossesAt


class DisplayUnitsLabel(Serialisable):

    tagname = "dispUnitsLbl"

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Text, allow_none=True)
    text = Alias("tx")
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    txPr = Typed(expected_type=RichText, allow_none=True)
    textPropertes = Alias("txPr")

    __elements__ = ('layout', 'tx', 'spPr', 'txPr')

    def __init__(self,
                 layout=None,
                 tx=None,
                 spPr=None,
                 txPr=None,
                ):
        self.layout = layout
        self.tx = tx
        self.spPr = spPr
        self.txPr = txPr


class DisplayUnitsLabelList(Serialisable):

    tagname = "dispUnits"

    custUnit = NestedFloat(allow_none=True)
    builtInUnit = NestedNoneSet(values=(['hundreds', 'thousands',
                                         'tenThousands', 'hundredThousands', 'millions', 'tenMillions',
                                         'hundredMillions', 'billions', 'trillions']))
    dispUnitsLbl = Typed(expected_type=DisplayUnitsLabel, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('custUnit', 'builtInUnit', 'dispUnitsLbl',)

    def __init__(self,
                 custUnit=None,
                 builtInUnit=None,
                 dispUnitsLbl=None,
                 extLst=None,
                ):
        self.custUnit = custUnit
        self.builtInUnit = builtInUnit
        self.dispUnitsLbl = dispUnitsLbl


class NumericAxis(_BaseAxis):

    tagname = "valAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txPr = _BaseAxis.txPr
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    crossBetween = NestedNoneSet(values=(['between', 'midCat']))
    majorUnit = NestedFloat(allow_none=True)
    minorUnit = NestedFloat(allow_none=True)
    dispUnits = Typed(expected_type=DisplayUnitsLabelList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('crossBetween', 'majorUnit',
                                             'minorUnit', 'dispUnits',)


    def __init__(self,
                 crossBetween=None,
                 majorUnit=None,
                 minorUnit=None,
                 dispUnits=None,
                 extLst=None,
                 **kw
                ):
        self.crossBetween = crossBetween
        self.majorUnit = majorUnit
        self.minorUnit = minorUnit
        self.dispUnits = dispUnits
        kw.setdefault('majorGridlines', ChartLines())
        kw.setdefault('axId', 100)
        kw.setdefault('crossAx', 10)
        super().__init__(**kw)


    @classmethod
    def from_tree(cls, node):
        """
        Special case value axes with no gridlines
        """
        self = super().from_tree(node)
        gridlines = node.find("{%s}majorGridlines" % CHART_NS)
        if gridlines is None:
            self.majorGridlines = None
        return self



class TextAxis(_BaseAxis):

    tagname = "catAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txPr = _BaseAxis.txPr
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    auto = NestedBool(allow_none=True)
    lblAlgn = NestedNoneSet(values=(['ctr', 'l', 'r']))
    lblOffset = NestedMinMax(min=0, max=1000)
    tickLblSkip = NestedInteger(allow_none=True)
    tickMarkSkip = NestedInteger(allow_none=True)
    noMultiLvlLbl = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('auto', 'lblAlgn', 'lblOffset',
                                             'tickLblSkip', 'tickMarkSkip', 'noMultiLvlLbl')

    def __init__(self,
                 auto=None,
                 lblAlgn=None,
                 lblOffset=100,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 noMultiLvlLbl=None,
                 extLst=None,
                 **kw
                ):
        self.auto = auto
        self.lblAlgn = lblAlgn
        self.lblOffset = lblOffset
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.noMultiLvlLbl = noMultiLvlLbl
        kw.setdefault('axId', 10)
        kw.setdefault('crossAx', 100)
        super().__init__(**kw)


class DateAxis(TextAxis):

    tagname = "dateAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txPr = _BaseAxis.txPr
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    auto = NestedBool(allow_none=True)
    lblOffset = NestedInteger(allow_none=True)
    baseTimeUnit = NestedNoneSet(values=(['days', 'months', 'years']))
    majorUnit = NestedFloat(allow_none=True)
    majorTimeUnit = NestedNoneSet(values=(['days', 'months', 'years']))
    minorUnit = NestedFloat(allow_none=True)
    minorTimeUnit = NestedNoneSet(values=(['days', 'months', 'years']))
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('auto', 'lblOffset',
                                             'baseTimeUnit', 'majorUnit', 'majorTimeUnit', 'minorUnit',
                                             'minorTimeUnit')

    def __init__(self,
                 auto=None,
                 lblOffset=None,
                 baseTimeUnit=None,
                 majorUnit=None,
                 majorTimeUnit=None,
                 minorUnit=None,
                 minorTimeUnit=None,
                 extLst=None,
                 **kw
                ):
        self.auto = auto
        self.lblOffset = lblOffset
        self.baseTimeUnit = baseTimeUnit
        self.majorUnit = majorUnit
        self.majorTimeUnit = majorTimeUnit
        self.minorUnit = minorUnit
        self.minorTimeUnit = minorTimeUnit
        kw.setdefault('axId', 500)
        kw.setdefault('lblOffset', lblOffset)
        super().__init__(**kw)


class SeriesAxis(_BaseAxis):

    tagname = "serAx"

    axId = _BaseAxis.axId
    scaling = _BaseAxis.scaling
    delete = _BaseAxis.delete
    axPos = _BaseAxis.axPos
    majorGridlines = _BaseAxis.majorGridlines
    minorGridlines = _BaseAxis.minorGridlines
    title = _BaseAxis.title
    numFmt = _BaseAxis.numFmt
    majorTickMark = _BaseAxis.majorTickMark
    minorTickMark = _BaseAxis.minorTickMark
    tickLblPos = _BaseAxis.tickLblPos
    spPr = _BaseAxis.spPr
    txPr = _BaseAxis.txPr
    crossAx = _BaseAxis.crossAx
    crosses = _BaseAxis.crosses
    crossesAt = _BaseAxis.crossesAt

    tickLblSkip = NestedInteger(allow_none=True)
    tickMarkSkip = NestedInteger(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _BaseAxis.__elements__ + ('tickLblSkip', 'tickMarkSkip')

    def __init__(self,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 extLst=None,
                 **kw
                ):
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        kw.setdefault('axId', 1000)
        kw.setdefault('crossAx', 10)
        super().__init__(**kw)

''',
    },
    'openpyxl.chart.bar_chart': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    Sequence,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedSet,
    NestedBool,
    NestedInteger,
    NestedMinMax,
)

from .descriptors import (
    NestedGapAmount,
    NestedOverlap,
)
from ._chart import ChartBase
from ._3d import _3DBase
from .axis import TextAxis, NumericAxis, SeriesAxis, ChartLines
from .shapes import GraphicalProperties
from .series import Series
from .legend import Legend
from .label import DataLabelList


class _BarChartBase(ChartBase):

    barDir = NestedSet(values=(['bar', 'col']))
    type = Alias("barDir")
    grouping = NestedSet(values=(['percentStacked', 'clustered', 'standard',
                                  'stacked']))
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")

    __elements__ = ('barDir', 'grouping', 'varyColors', 'ser', 'dLbls')

    _series_type = "bar"

    def __init__(self,
                 barDir="col",
                 grouping="clustered",
                 varyColors=None,
                 ser=(),
                 dLbls=None,
                 **kw
                ):
        self.barDir = barDir
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        super().__init__(**kw)


class BarChart(_BarChartBase):

    tagname = "barChart"

    barDir = _BarChartBase.barDir
    grouping = _BarChartBase.grouping
    varyColors = _BarChartBase.varyColors
    ser = _BarChartBase.ser
    dLbls = _BarChartBase.dLbls

    gapWidth = NestedGapAmount()
    overlap = NestedOverlap()
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    # chart properties actually used by containing classes
    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)

    __elements__ = _BarChartBase.__elements__ + ('gapWidth', 'overlap', 'serLines', 'axId')

    def __init__(self,
                 gapWidth=150,
                 overlap=None,
                 serLines=None,
                 extLst=None,
                 **kw
                ):
        self.gapWidth = gapWidth
        self.overlap = overlap
        self.serLines = serLines
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.legend = Legend()
        super().__init__(**kw)


class BarChart3D(_BarChartBase, _3DBase):

    tagname = "bar3DChart"

    barDir = _BarChartBase.barDir
    grouping = _BarChartBase.grouping
    varyColors = _BarChartBase.varyColors
    ser = _BarChartBase.ser
    dLbls = _BarChartBase.dLbls

    view3D = _3DBase.view3D
    floor = _3DBase.floor
    sideWall = _3DBase.sideWall
    backWall = _3DBase.backWall

    gapWidth = NestedGapAmount()
    gapDepth = NestedGapAmount()
    shape = NestedNoneSet(values=(['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax']))
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)
    z_axis = Typed(expected_type=SeriesAxis, allow_none=True)

    __elements__ = _BarChartBase.__elements__ + ('gapWidth', 'gapDepth', 'shape', 'serLines', 'axId')

    def __init__(self,
                 gapWidth=150,
                 gapDepth=150,
                 shape=None,
                 serLines=None,
                 extLst=None,
                 **kw
                ):
        self.gapWidth = gapWidth
        self.gapDepth = gapDepth
        self.shape = shape
        self.serLines = serLines
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()

        super(BarChart3D, self).__init__(**kw)

""",
    },
    'openpyxl.chart.bubble_chart': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Set,
    MinMax,
    Bool,
    Integer,
    Alias,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedMinMax,
    NestedBool,
)

from ._chart import ChartBase
from .axis import TextAxis, NumericAxis
from .series import XYSeries
from .label import DataLabelList


class BubbleChart(ChartBase):

    tagname = "bubbleChart"

    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=XYSeries, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    bubble3D = NestedBool(allow_none=True)
    bubbleScale = NestedMinMax(min=0, max=300, allow_none=True)
    showNegBubbles = NestedBool(allow_none=True)
    sizeRepresents = NestedNoneSet(values=(['area', 'w']))
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=NumericAxis)
    y_axis = Typed(expected_type=NumericAxis)

    _series_type = "bubble"

    __elements__ = ('varyColors', 'ser', 'dLbls', 'bubble3D', 'bubbleScale',
                    'showNegBubbles', 'sizeRepresents', 'axId')

    def __init__(self,
                 varyColors=None,
                 ser=(),
                 dLbls=None,
                 bubble3D=None,
                 bubbleScale=None,
                 showNegBubbles=None,
                 sizeRepresents=None,
                 extLst=None,
                 **kw
                ):
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.bubble3D = bubble3D
        self.bubbleScale = bubbleScale
        self.showNegBubbles = showNegBubbles
        self.sizeRepresents = sizeRepresents
        self.x_axis = NumericAxis(axId=10, crossAx=20)
        self.y_axis = NumericAxis(axId=20, crossAx=10)
        super().__init__(**kw)

""",
    },
    'openpyxl.chart.chartspace': {
        'is_package': False,
        'source': r'''

# Copyright (c) 2010-2024 openpyxl

"""
Enclosing chart object. The various chart types are actually child objects.
Will probably need to call this indirectly
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Alias,
)
from openpyxl.descriptors.excel import (
    ExtensionList,
    Relation
)
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedNoneSet,
    NestedString,
    NestedMinMax,
)
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.xml.constants import CHART_NS

from openpyxl.drawing.colors import ColorMapping
from .text import RichText
from .shapes import GraphicalProperties
from .legend import Legend
from ._3d import _3DBase
from .plotarea import PlotArea
from .title import Title
from .pivot import (
    PivotFormat,
    PivotSource,
)
from .print_settings import PrintSettings


class ChartContainer(Serialisable):

    tagname = "chart"

    title = Typed(expected_type=Title, allow_none=True)
    autoTitleDeleted = NestedBool(allow_none=True)
    pivotFmts = NestedSequence(expected_type=PivotFormat)
    view3D = _3DBase.view3D
    floor = _3DBase.floor
    sideWall = _3DBase.sideWall
    backWall = _3DBase.backWall
    plotArea = Typed(expected_type=PlotArea, )
    legend = Typed(expected_type=Legend, allow_none=True)
    plotVisOnly = NestedBool()
    dispBlanksAs = NestedNoneSet(values=(['span', 'gap', 'zero']))
    showDLblsOverMax = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('title', 'autoTitleDeleted', 'pivotFmts', 'view3D',
                    'floor', 'sideWall', 'backWall', 'plotArea', 'legend', 'plotVisOnly',
                    'dispBlanksAs', 'showDLblsOverMax')

    def __init__(self,
                 title=None,
                 autoTitleDeleted=None,
                 pivotFmts=(),
                 view3D=None,
                 floor=None,
                 sideWall=None,
                 backWall=None,
                 plotArea=None,
                 legend=None,
                 plotVisOnly=True,
                 dispBlanksAs="gap",
                 showDLblsOverMax=None,
                 extLst=None,
                ):
        self.title = title
        self.autoTitleDeleted = autoTitleDeleted
        self.pivotFmts = pivotFmts
        self.view3D = view3D
        self.floor = floor
        self.sideWall = sideWall
        self.backWall = backWall
        if plotArea is None:
            plotArea = PlotArea()
        self.plotArea = plotArea
        self.legend = legend
        self.plotVisOnly = plotVisOnly
        self.dispBlanksAs = dispBlanksAs
        self.showDLblsOverMax = showDLblsOverMax


class Protection(Serialisable):

    tagname = "protection"

    chartObject = NestedBool(allow_none=True)
    data = NestedBool(allow_none=True)
    formatting = NestedBool(allow_none=True)
    selection = NestedBool(allow_none=True)
    userInterface = NestedBool(allow_none=True)

    __elements__ = ("chartObject", "data", "formatting", "selection", "userInterface")

    def __init__(self,
                 chartObject=None,
                 data=None,
                 formatting=None,
                 selection=None,
                 userInterface=None,
                ):
        self.chartObject = chartObject
        self.data = data
        self.formatting = formatting
        self.selection = selection
        self.userInterface = userInterface


class ExternalData(Serialisable):

    tagname = "externalData"

    autoUpdate = NestedBool(allow_none=True)
    id = String() # Needs namespace

    def __init__(self,
                 autoUpdate=None,
                 id=None
                ):
        self.autoUpdate = autoUpdate
        self.id = id


class ChartSpace(Serialisable):

    tagname = "chartSpace"

    date1904 = NestedBool(allow_none=True)
    lang = NestedString(allow_none=True)
    roundedCorners = NestedBool(allow_none=True)
    style = NestedMinMax(allow_none=True, min=1, max=48)
    clrMapOvr = Typed(expected_type=ColorMapping, allow_none=True)
    pivotSource = Typed(expected_type=PivotSource, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)
    chart = Typed(expected_type=ChartContainer)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphical_properties = Alias("spPr")
    txPr = Typed(expected_type=RichText, allow_none=True)
    textProperties = Alias("txPr")
    externalData = Typed(expected_type=ExternalData, allow_none=True)
    printSettings = Typed(expected_type=PrintSettings, allow_none=True)
    userShapes = Relation()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('date1904', 'lang', 'roundedCorners', 'style',
                    'clrMapOvr', 'pivotSource', 'protection', 'chart', 'spPr', 'txPr',
                    'externalData', 'printSettings', 'userShapes')

    def __init__(self,
                 date1904=None,
                 lang=None,
                 roundedCorners=None,
                 style=None,
                 clrMapOvr=None,
                 pivotSource=None,
                 protection=None,
                 chart=None,
                 spPr=None,
                 txPr=None,
                 externalData=None,
                 printSettings=None,
                 userShapes=None,
                 extLst=None,
                ):
        self.date1904 = date1904
        self.lang = lang
        self.roundedCorners = roundedCorners
        self.style = style
        self.clrMapOvr = clrMapOvr
        self.pivotSource = pivotSource
        self.protection = protection
        self.chart = chart
        self.spPr = spPr
        self.txPr = txPr
        self.externalData = externalData
        self.printSettings = printSettings
        self.userShapes = userShapes


    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree()
        tree.set("xmlns", CHART_NS)
        return tree

''',
    },
    'openpyxl.chart.data_source': {
        'is_package': False,
        'source': r'''
"""
Collection of utility primitives for charts.
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Typed,
    Alias,
    String,
    Integer,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedString,
    NestedText,
    NestedInteger,
)


class NumFmt(Serialisable):

    formatCode = String()
    sourceLinked = Bool()

    def __init__(self,
                 formatCode=None,
                 sourceLinked=False
                ):
        self.formatCode = formatCode
        self.sourceLinked = sourceLinked


class NumberValueDescriptor(NestedText):
    """
    Data should be numerical but isn't always :-/
    """

    allow_none = True

    def __set__(self, instance, value):
        if value == "#N/A":
            self.expected_type = str
        else:
            self.expected_type = float
        super().__set__(instance, value)


class NumVal(Serialisable):

    idx = Integer()
    formatCode = NestedText(allow_none=True, expected_type=str)
    v = NumberValueDescriptor()

    def __init__(self,
                 idx=None,
                 formatCode=None,
                 v=None,
                ):
        self.idx = idx
        self.formatCode = formatCode
        self.v = v


class NumData(Serialisable):

    formatCode = NestedText(expected_type=str, allow_none=True)
    ptCount = NestedInteger(allow_none=True)
    pt = Sequence(expected_type=NumVal)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('formatCode', 'ptCount', 'pt')

    def __init__(self,
                 formatCode=None,
                 ptCount=None,
                 pt=(),
                 extLst=None,
                ):
        self.formatCode = formatCode
        self.ptCount = ptCount
        self.pt = pt


class NumRef(Serialisable):

    f = NestedText(expected_type=str)
    ref = Alias('f')
    numCache = Typed(expected_type=NumData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f', 'numCache')

    def __init__(self,
                 f=None,
                 numCache=None,
                 extLst=None,
                ):
        self.f = f
        self.numCache = numCache


class StrVal(Serialisable):

    tagname = "strVal"

    idx = Integer()
    v = NestedText(expected_type=str)

    def __init__(self,
                 idx=0,
                 v=None,
                ):
        self.idx = idx
        self.v = v


class StrData(Serialisable):

    tagname = "strData"

    ptCount = NestedInteger(allow_none=True)
    pt = Sequence(expected_type=StrVal)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ptCount', 'pt')

    def __init__(self,
                 ptCount=None,
                 pt=(),
                 extLst=None,
                ):
        self.ptCount = ptCount
        self.pt = pt


class StrRef(Serialisable):

    tagname = "strRef"

    f = NestedText(expected_type=str, allow_none=True)
    strCache = Typed(expected_type=StrData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f', 'strCache')

    def __init__(self,
                 f=None,
                 strCache=None,
                 extLst=None,
                ):
        self.f = f
        self.strCache = strCache


class NumDataSource(Serialisable):

    numRef = Typed(expected_type=NumRef, allow_none=True)
    numLit = Typed(expected_type=NumData, allow_none=True)


    def __init__(self,
                 numRef=None,
                 numLit=None,
                 ):
        self.numRef = numRef
        self.numLit = numLit


class Level(Serialisable):

    tagname = "lvl"

    pt = Sequence(expected_type=StrVal)

    __elements__ = ('pt',)

    def __init__(self,
                 pt=(),
                ):
        self.pt = pt


class MultiLevelStrData(Serialisable):

    tagname = "multiLvlStrData"

    ptCount = Integer(allow_none=True)
    lvl = Sequence(expected_type=Level)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ptCount', 'lvl',)

    def __init__(self,
                 ptCount=None,
                 lvl=(),
                 extLst=None,
                ):
        self.ptCount = ptCount
        self.lvl = lvl


class MultiLevelStrRef(Serialisable):

    tagname = "multiLvlStrRef"

    f = NestedText(expected_type=str)
    multiLvlStrCache = Typed(expected_type=MultiLevelStrData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('multiLvlStrCache', 'f')

    def __init__(self,
                 f=None,
                 multiLvlStrCache=None,
                 extLst=None,
                ):
        self.f = f
        self.multiLvlStrCache = multiLvlStrCache


class AxDataSource(Serialisable):

    tagname = "cat"

    numRef = Typed(expected_type=NumRef, allow_none=True)
    numLit = Typed(expected_type=NumData, allow_none=True)
    strRef = Typed(expected_type=StrRef, allow_none=True)
    strLit = Typed(expected_type=StrData, allow_none=True)
    multiLvlStrRef = Typed(expected_type=MultiLevelStrRef, allow_none=True)

    def __init__(self,
                 numRef=None,
                 numLit=None,
                 strRef=None,
                 strLit=None,
                 multiLvlStrRef=None,
                 ):
        if not any([numLit, numRef, strRef, strLit, multiLvlStrRef]):
            raise TypeError("A data source must be provided")
        self.numRef = numRef
        self.numLit = numLit
        self.strRef = strRef
        self.strLit = strLit
        self.multiLvlStrRef = multiLvlStrRef

''',
    },
    'openpyxl.chart.descriptors': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl



from openpyxl.descriptors.nested import (
    NestedMinMax
    )

from openpyxl.descriptors import Typed

from .data_source import NumFmt

"""
Utility descriptors for the chart module.
For convenience but also clarity.
"""

class NestedGapAmount(NestedMinMax):

    allow_none = True
    min = 0
    max = 500


class NestedOverlap(NestedMinMax):

    allow_none = True
    min = -100
    max = 100


class NumberFormatDescriptor(Typed):
    """
    Allow direct assignment of format code
    """

    expected_type = NumFmt
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, str):
            value = NumFmt(value)
        super().__set__(instance, value)

''',
    },
    'openpyxl.chart.error_bar': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Set,
    Alias
)

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedSet,
    NestedBool,
    NestedFloat,
)

from .data_source import NumDataSource
from .shapes import GraphicalProperties


class ErrorBars(Serialisable):

    tagname = "errBars"

    errDir = NestedNoneSet(values=(['x', 'y']))
    direction = Alias("errDir")
    errBarType = NestedSet(values=(['both', 'minus', 'plus']))
    style = Alias("errBarType")
    errValType = NestedSet(values=(['cust', 'fixedVal', 'percentage', 'stdDev', 'stdErr']))
    size = Alias("errValType")
    noEndCap = NestedBool(nested=True, allow_none=True)
    plus = Typed(expected_type=NumDataSource, allow_none=True)
    minus = Typed(expected_type=NumDataSource, allow_none=True)
    val = NestedFloat(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('errDir','errBarType', 'errValType', 'noEndCap','minus', 'plus', 'val', 'spPr')


    def __init__(self,
                 errDir=None,
                 errBarType="both",
                 errValType="fixedVal",
                 noEndCap=None,
                 plus=None,
                 minus=None,
                 val=None,
                 spPr=None,
                 extLst=None,
                ):
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr

""",
    },
    'openpyxl.chart.label': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    Alias,
    Typed
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedBool,
    NestedString,
    NestedInteger,
    )

from .shapes import GraphicalProperties
from .text import RichText


class _DataLabelBase(Serialisable):

    numFmt = NestedString(allow_none=True, attribute="formatCode")
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    txPr = Typed(expected_type=RichText, allow_none=True)
    textProperties = Alias('txPr')
    dLblPos = NestedNoneSet(values=['bestFit', 'b', 'ctr', 'inBase', 'inEnd',
                                    'l', 'outEnd', 'r', 't'])
    position = Alias('dLblPos')
    showLegendKey = NestedBool(allow_none=True)
    showVal = NestedBool(allow_none=True)
    showCatName = NestedBool(allow_none=True)
    showSerName = NestedBool(allow_none=True)
    showPercent = NestedBool(allow_none=True)
    showBubbleSize = NestedBool(allow_none=True)
    showLeaderLines = NestedBool(allow_none=True)
    separator = NestedString(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ("numFmt", "spPr", "txPr", "dLblPos", "showLegendKey",
                    "showVal", "showCatName", "showSerName", "showPercent", "showBubbleSize",
                    "showLeaderLines", "separator")

    def __init__(self,
                 numFmt=None,
                 spPr=None,
                 txPr=None,
                 dLblPos=None,
                 showLegendKey=None,
                 showVal=None,
                 showCatName=None,
                 showSerName=None,
                 showPercent=None,
                 showBubbleSize=None,
                 showLeaderLines=None,
                 separator=None,
                 extLst=None,
                 ):
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr
        self.dLblPos = dLblPos
        self.showLegendKey = showLegendKey
        self.showVal = showVal
        self.showCatName = showCatName
        self.showSerName = showSerName
        self.showPercent = showPercent
        self.showBubbleSize = showBubbleSize
        self.showLeaderLines = showLeaderLines
        self.separator = separator


class DataLabel(_DataLabelBase):

    tagname = "dLbl"

    idx = NestedInteger()

    numFmt = _DataLabelBase.numFmt
    spPr = _DataLabelBase.spPr
    txPr = _DataLabelBase.txPr
    dLblPos = _DataLabelBase.dLblPos
    showLegendKey = _DataLabelBase.showLegendKey
    showVal = _DataLabelBase.showVal
    showCatName = _DataLabelBase.showCatName
    showSerName = _DataLabelBase.showSerName
    showPercent = _DataLabelBase.showPercent
    showBubbleSize = _DataLabelBase.showBubbleSize
    showLeaderLines = _DataLabelBase.showLeaderLines
    separator = _DataLabelBase.separator
    extLst = _DataLabelBase.extLst

    __elements__ = ("idx",)  + _DataLabelBase.__elements__

    def __init__(self, idx=0, **kw ):
        self.idx = idx
        super().__init__(**kw)


class DataLabelList(_DataLabelBase):

    tagname = "dLbls"

    dLbl = Sequence(expected_type=DataLabel, allow_none=True)

    delete = NestedBool(allow_none=True)
    numFmt = _DataLabelBase.numFmt
    spPr = _DataLabelBase.spPr
    txPr = _DataLabelBase.txPr
    dLblPos = _DataLabelBase.dLblPos
    showLegendKey = _DataLabelBase.showLegendKey
    showVal = _DataLabelBase.showVal
    showCatName = _DataLabelBase.showCatName
    showSerName = _DataLabelBase.showSerName
    showPercent = _DataLabelBase.showPercent
    showBubbleSize = _DataLabelBase.showBubbleSize
    showLeaderLines = _DataLabelBase.showLeaderLines
    separator = _DataLabelBase.separator
    extLst = _DataLabelBase.extLst

    __elements__ = ("delete", "dLbl",) + _DataLabelBase.__elements__

    def __init__(self, dLbl=(), delete=None,  **kw):
        self.dLbl = dLbl
        self.delete = delete
        super().__init__(**kw)

""",
    },
    'openpyxl.chart.layout': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    NoneSet,
    Float,
    Typed,
    Alias,
)

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedSet,
    NestedMinMax,
)

class ManualLayout(Serialisable):

    tagname = "manualLayout"

    layoutTarget = NestedNoneSet(values=(['inner', 'outer']))
    xMode = NestedNoneSet(values=(['edge', 'factor']))
    yMode = NestedNoneSet(values=(['edge', 'factor']))
    wMode = NestedSet(values=(['edge', 'factor']))
    hMode = NestedSet(values=(['edge', 'factor']))
    x = NestedMinMax(min=-1, max=1, allow_none=True)
    y = NestedMinMax(min=-1, max=1, allow_none=True)
    w = NestedMinMax(min=0, max=1, allow_none=True)
    width = Alias('w')
    h = NestedMinMax(min=0, max=1,  allow_none=True)
    height = Alias('h')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('layoutTarget', 'xMode', 'yMode', 'wMode', 'hMode', 'x',
                    'y', 'w', 'h')

    def __init__(self,
                 layoutTarget=None,
                 xMode=None,
                 yMode=None,
                 wMode="factor",
                 hMode="factor",
                 x=None,
                 y=None,
                 w=None,
                 h=None,
                 extLst=None,
                ):
        self.layoutTarget = layoutTarget
        self.xMode = xMode
        self.yMode = yMode
        self.wMode = wMode
        self.hMode = hMode
        self.x = x
        self.y = y
        self.w = w
        self.h = h


class Layout(Serialisable):

    tagname = "layout"

    manualLayout = Typed(expected_type=ManualLayout, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('manualLayout',)

    def __init__(self,
                 manualLayout=None,
                 extLst=None,
                ):
        self.manualLayout = manualLayout

""",
    },
    'openpyxl.chart.legend': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    Alias,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedSet,
    NestedInteger
)

from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText


class LegendEntry(Serialisable):

    tagname = "legendEntry"

    idx = NestedInteger()
    delete = NestedBool()
    txPr = Typed(expected_type=RichText, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'delete', 'txPr')

    def __init__(self,
                 idx=0,
                 delete=False,
                 txPr=None,
                 extLst=None,
                ):
        self.idx = idx
        self.delete = delete
        self.txPr = txPr


class Legend(Serialisable):

    tagname = "legend"

    legendPos = NestedSet(values=(['b', 'tr', 'l', 'r', 't']))
    position = Alias('legendPos')
    legendEntry = Sequence(expected_type=LegendEntry)
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = NestedBool(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    txPr = Typed(expected_type=RichText, allow_none=True)
    textProperties = Alias('txPr')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr',)

    def __init__(self,
                 legendPos="r",
                 legendEntry=(),
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.legendPos = legendPos
        self.legendEntry = legendEntry
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr

""",
    },
    'openpyxl.chart.line_chart': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors import (
    Typed,
    Sequence,
    Alias,
    )
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedSet,
    NestedBool,
)

from ._chart import ChartBase
from .updown_bars import UpDownBars
from .descriptors import NestedGapAmount
from .axis import TextAxis, NumericAxis, SeriesAxis, ChartLines, _BaseAxis
from .label import DataLabelList
from .series import Series


class _LineChartBase(ChartBase):

    grouping = NestedSet(values=(['percentStacked', 'standard', 'stacked']))
    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    dropLines = Typed(expected_type=ChartLines, allow_none=True)

    _series_type = "line"

    __elements__ = ('grouping', 'varyColors', 'ser', 'dLbls', 'dropLines')

    def __init__(self,
                 grouping="standard",
                 varyColors=None,
                 ser=(),
                 dLbls=None,
                 dropLines=None,
                 **kw
                ):
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        super().__init__(**kw)


class LineChart(_LineChartBase):

    tagname = "lineChart"

    grouping = _LineChartBase.grouping
    varyColors = _LineChartBase.varyColors
    ser = _LineChartBase.ser
    dLbls = _LineChartBase.dLbls
    dropLines =_LineChartBase.dropLines

    hiLowLines = Typed(expected_type=ChartLines, allow_none=True)
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    marker = NestedBool(allow_none=True)
    smooth = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=_BaseAxis)
    y_axis = Typed(expected_type=NumericAxis)

    __elements__ = _LineChartBase.__elements__ + ('hiLowLines', 'upDownBars', 'marker', 'smooth', 'axId')

    def __init__(self,
                 hiLowLines=None,
                 upDownBars=None,
                 marker=None,
                 smooth=None,
                 extLst=None,
                 **kw
                ):
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.marker = marker
        self.smooth = smooth
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()

        super().__init__(**kw)


class LineChart3D(_LineChartBase):

    tagname = "line3DChart"

    grouping = _LineChartBase.grouping
    varyColors = _LineChartBase.varyColors
    ser = _LineChartBase.ser
    dLbls = _LineChartBase.dLbls
    dropLines =_LineChartBase.dropLines

    gapDepth = NestedGapAmount()
    hiLowLines = Typed(expected_type=ChartLines, allow_none=True)
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    marker = NestedBool(allow_none=True)
    smooth = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)
    z_axis = Typed(expected_type=SeriesAxis)

    __elements__ = _LineChartBase.__elements__ + ('gapDepth', 'hiLowLines',
                                                  'upDownBars', 'marker', 'smooth', 'axId')

    def __init__(self,
                 gapDepth=None,
                 hiLowLines=None,
                 upDownBars=None,
                 marker=None,
                 smooth=None,
                 **kw
                ):
        self.gapDepth = gapDepth
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.marker = marker
        self.smooth = smooth
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()
        super(LineChart3D, self).__init__(**kw)

""",
    },
    'openpyxl.chart.marker': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Alias,
)

from openpyxl.descriptors.excel import(
    ExtensionList,
    _explicit_none,
)

from openpyxl.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedMinMax,
    NestedNoneSet,
)

from .layout import Layout
from .picture import PictureOptions
from .shapes import *
from .text import *
from .error_bar import *


class Marker(Serialisable):

    tagname = "marker"

    symbol = NestedNoneSet(values=(['circle', 'dash', 'diamond', 'dot', 'picture',
                              'plus', 'square', 'star', 'triangle', 'x', 'auto']),
                           to_tree=_explicit_none)
    size = NestedMinMax(min=2, max=72, allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('symbol', 'size', 'spPr')

    def __init__(self,
                 symbol=None,
                 size=None,
                 spPr=None,
                 extLst=None,
                ):
        self.symbol = symbol
        self.size = size
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr


class DataPoint(Serialisable):

    tagname = "dPt"

    idx = NestedInteger()
    invertIfNegative = NestedBool(allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    bubble3D = NestedBool(allow_none=True)
    explosion = NestedInteger(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'invertIfNegative', 'marker', 'bubble3D',
                    'explosion', 'spPr', 'pictureOptions')

    def __init__(self,
                 idx=None,
                 invertIfNegative=None,
                 marker=None,
                 bubble3D=None,
                 explosion=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.idx = idx
        self.invertIfNegative = invertIfNegative
        self.marker = marker
        self.bubble3D = bubble3D
        self.explosion = explosion
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.pictureOptions = pictureOptions

""",
    },
    'openpyxl.chart.picture': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.descriptors.nested import (
    NestedBool,
    NestedFloat,
    NestedMinMax,
    NestedNoneSet,
)

class PictureOptions(Serialisable):

    tagname = "pictureOptions"

    applyToFront = NestedBool(allow_none=True, nested=True)
    applyToSides = NestedBool(allow_none=True, nested=True)
    applyToEnd = NestedBool(allow_none=True, nested=True)
    pictureFormat = NestedNoneSet(values=(['stretch', 'stack', 'stackScale']), nested=True)
    pictureStackUnit = NestedFloat(allow_none=True, nested=True)

    __elements__ = ('applyToFront', 'applyToSides', 'applyToEnd', 'pictureFormat', 'pictureStackUnit')

    def __init__(self,
                 applyToFront=None,
                 applyToSides=None,
                 applyToEnd=None,
                 pictureFormat=None,
                 pictureStackUnit=None,
                ):
        self.applyToFront = applyToFront
        self.applyToSides = applyToSides
        self.applyToEnd = applyToEnd
        self.pictureFormat = pictureFormat
        self.pictureStackUnit = pictureStackUnit

""",
    },
    'openpyxl.chart.pie_chart': {
        'is_package': False,
        'source': r'''
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    MinMax,
    Integer,
    NoneSet,
    Float,
    Alias,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList, Percentage
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedMinMax,
    NestedInteger,
    NestedFloat,
    NestedNoneSet,
    NestedSet,
)
from openpyxl.descriptors.sequence import ValueSequence

from ._chart import ChartBase
from .axis import ChartLines
from .descriptors import NestedGapAmount
from .series import Series
from .label import DataLabelList


class _PieChartBase(ChartBase):

    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")

    _series_type = "pie"

    __elements__ = ('varyColors', 'ser', 'dLbls')

    def __init__(self,
                 varyColors=True,
                 ser=(),
                 dLbls=None,
                ):
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        super().__init__()



class PieChart(_PieChartBase):

    tagname = "pieChart"

    varyColors = _PieChartBase.varyColors
    ser = _PieChartBase.ser
    dLbls = _PieChartBase.dLbls

    firstSliceAng = NestedMinMax(min=0, max=360)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _PieChartBase.__elements__ + ('firstSliceAng', )

    def __init__(self,
                 firstSliceAng=0,
                 extLst=None,
                 **kw
                ):
        self.firstSliceAng = firstSliceAng
        super().__init__(**kw)


class PieChart3D(_PieChartBase):

    tagname = "pie3DChart"

    varyColors = _PieChartBase.varyColors
    ser = _PieChartBase.ser
    dLbls = _PieChartBase.dLbls

    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _PieChartBase.__elements__


class DoughnutChart(_PieChartBase):

    tagname = "doughnutChart"

    varyColors = _PieChartBase.varyColors
    ser = _PieChartBase.ser
    dLbls = _PieChartBase.dLbls

    firstSliceAng = NestedMinMax(min=0, max=360)
    holeSize = NestedMinMax(min=1, max=90, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _PieChartBase.__elements__ + ('firstSliceAng', 'holeSize')

    def __init__(self,
                 firstSliceAng=0,
                 holeSize=10,
                 extLst=None,
                 **kw
                ):
        self.firstSliceAng = firstSliceAng
        self.holeSize = holeSize
        super().__init__(**kw)


class CustomSplit(Serialisable):

    tagname = "custSplit"

    secondPiePt = ValueSequence(expected_type=int)

    __elements__ = ('secondPiePt',)

    def __init__(self,
                 secondPiePt=(),
                ):
        self.secondPiePt = secondPiePt


class ProjectedPieChart(_PieChartBase):

    """
    From the spec 21.2.2.126

    This element contains the pie of pie or bar of pie series on this
    chart. Only the first series shall be displayed. The splitType element
    shall determine whether the splitPos and custSplit elements apply.
    """

    tagname = "ofPieChart"

    varyColors = _PieChartBase.varyColors
    ser = _PieChartBase.ser
    dLbls = _PieChartBase.dLbls

    ofPieType = NestedSet(values=(['pie', 'bar']))
    type = Alias('ofPieType')
    gapWidth = NestedGapAmount()
    splitType = NestedNoneSet(values=(['auto', 'cust', 'percent', 'pos', 'val']))
    splitPos = NestedFloat(allow_none=True)
    custSplit = Typed(expected_type=CustomSplit, allow_none=True)
    secondPieSize = NestedMinMax(min=5, max=200, allow_none=True)
    serLines = Typed(expected_type=ChartLines, allow_none=True)
    join_lines = Alias('serLines')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _PieChartBase.__elements__ + ('ofPieType', 'gapWidth',
                                                 'splitType', 'splitPos', 'custSplit', 'secondPieSize', 'serLines')

    def __init__(self,
                 ofPieType="pie",
                 gapWidth=None,
                 splitType="auto",
                 splitPos=None,
                 custSplit=None,
                 secondPieSize=75,
                 serLines=None,
                 extLst=None,
                 **kw
                ):
        self.ofPieType = ofPieType
        self.gapWidth = gapWidth
        self.splitType = splitType
        self.splitPos = splitPos
        self.custSplit = custSplit
        self.secondPieSize = secondPieSize
        if serLines is None:
            self.serLines = ChartLines()
        super().__init__(**kw)

''',
    },
    'openpyxl.chart.pivot': {
        'is_package': False,
        'source': r"""

# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
)
from openpyxl.descriptors.nested import NestedInteger, NestedText
from openpyxl.descriptors.excel import ExtensionList

from .label import DataLabel
from .marker import Marker
from .shapes import GraphicalProperties
from .text import RichText


class PivotSource(Serialisable):

    tagname = "pivotSource"

    name = NestedText(expected_type=str)
    fmtId = NestedInteger(expected_type=int)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('name', 'fmtId')

    def __init__(self,
                 name=None,
                 fmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fmtId = fmtId


class PivotFormat(Serialisable):

    tagname = "pivotFmt"

    idx = NestedInteger(nested=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    txPr = Typed(expected_type=RichText, allow_none=True)
    TextBody = Alias("txPr")
    marker = Typed(expected_type=Marker, allow_none=True)
    dLbl = Typed(expected_type=DataLabel, allow_none=True)
    DataLabel = Alias("dLbl")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'spPr', 'txPr', 'marker', 'dLbl')

    def __init__(self,
                 idx=0,
                 spPr=None,
                 txPr=None,
                 marker=None,
                 dLbl=None,
                 extLst=None,
                ):
        self.idx = idx
        self.spPr = spPr
        self.txPr = txPr
        self.marker = marker
        self.dLbl = dLbl

""",
    },
    'openpyxl.chart.plotarea': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Alias,
)
from openpyxl.descriptors.excel import (
    ExtensionList,
)
from openpyxl.descriptors.sequence import (
    MultiSequence,
    MultiSequencePart,
)
from openpyxl.descriptors.nested import (
    NestedBool,
)

from ._3d import _3DBase
from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import PieChart, PieChart3D, ProjectedPieChart, DoughnutChart
from .radar_chart import RadarChart
from .scatter_chart import ScatterChart
from .stock_chart import StockChart
from .surface_chart import SurfaceChart, SurfaceChart3D
from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText

from .axis import (
    NumericAxis,
    TextAxis,
    SeriesAxis,
    DateAxis,
)


class DataTable(Serialisable):

    tagname = "dTable"

    showHorzBorder = NestedBool(allow_none=True)
    showVertBorder = NestedBool(allow_none=True)
    showOutline = NestedBool(allow_none=True)
    showKeys = NestedBool(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    txPr = Typed(expected_type=RichText, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('showHorzBorder', 'showVertBorder', 'showOutline',
                    'showKeys', 'spPr', 'txPr')

    def __init__(self,
                 showHorzBorder=None,
                 showVertBorder=None,
                 showOutline=None,
                 showKeys=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.showHorzBorder = showHorzBorder
        self.showVertBorder = showVertBorder
        self.showOutline = showOutline
        self.showKeys = showKeys
        self.spPr = spPr
        self.txPr = txPr


class PlotArea(Serialisable):

    tagname = "plotArea"

    layout = Typed(expected_type=Layout, allow_none=True)
    dTable = Typed(expected_type=DataTable, allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    # at least one chart
    _charts = MultiSequence()
    areaChart = MultiSequencePart(expected_type=AreaChart, store="_charts")
    area3DChart = MultiSequencePart(expected_type=AreaChart3D, store="_charts")
    lineChart = MultiSequencePart(expected_type=LineChart, store="_charts")
    line3DChart = MultiSequencePart(expected_type=LineChart3D, store="_charts")
    stockChart = MultiSequencePart(expected_type=StockChart, store="_charts")
    radarChart = MultiSequencePart(expected_type=RadarChart, store="_charts")
    scatterChart = MultiSequencePart(expected_type=ScatterChart, store="_charts")
    pieChart = MultiSequencePart(expected_type=PieChart, store="_charts")
    pie3DChart = MultiSequencePart(expected_type=PieChart3D, store="_charts")
    doughnutChart = MultiSequencePart(expected_type=DoughnutChart, store="_charts")
    barChart = MultiSequencePart(expected_type=BarChart, store="_charts")
    bar3DChart = MultiSequencePart(expected_type=BarChart3D, store="_charts")
    ofPieChart = MultiSequencePart(expected_type=ProjectedPieChart, store="_charts")
    surfaceChart = MultiSequencePart(expected_type=SurfaceChart, store="_charts")
    surface3DChart = MultiSequencePart(expected_type=SurfaceChart3D, store="_charts")
    bubbleChart = MultiSequencePart(expected_type=BubbleChart, store="_charts")

    # axes
    _axes = MultiSequence()
    valAx = MultiSequencePart(expected_type=NumericAxis, store="_axes")
    catAx = MultiSequencePart(expected_type=TextAxis, store="_axes")
    dateAx = MultiSequencePart(expected_type=DateAxis, store="_axes")
    serAx = MultiSequencePart(expected_type=SeriesAxis, store="_axes")

    __elements__ = ('layout', '_charts', '_axes', 'dTable', 'spPr')

    def __init__(self,
                 layout=None,
                 dTable=None,
                 spPr=None,
                 _charts=(),
                 _axes=(),
                 extLst=None,
                ):
        self.layout = layout
        self.dTable = dTable
        self.spPr = spPr
        self._charts = _charts
        self._axes = _axes


    def to_tree(self, tagname=None, idx=None, namespace=None):
        axIds = {ax.axId for ax in self._axes}
        for chart in self._charts:
            for id, axis in chart._axes.items():
                if id not in axIds:
                    setattr(self, axis.tagname, axis)
                    axIds.add(id)

        return super().to_tree(tagname)


    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        axes = dict((axis.axId, axis) for axis in self._axes)
        for chart in self._charts:
            if isinstance(chart, (ScatterChart, BubbleChart)):
                x, y = (axes[axId] for axId in chart.axId)
                chart.x_axis = x
                chart.y_axis = y
                continue

            for axId in chart.axId:
                axis = axes.get(axId)
                if axis is None and isinstance(chart, _3DBase):
                    # Series Axis can be optional
                    chart.z_axis = None
                    continue
                if axis.tagname in ("catAx", "dateAx"):
                    chart.x_axis = axis
                elif axis.tagname == "valAx":
                    chart.y_axis = axis
                elif axis.tagname == "serAx":
                    chart.z_axis = axis

        return self

""",
    },
    'openpyxl.chart.print_settings': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Float,
    Typed,
    Alias,
)

from openpyxl.worksheet.page import PrintPageSetup
from openpyxl.worksheet.header_footer import HeaderFooter


class PageMargins(Serialisable):
    """
    Identical to openpyxl.worksheet.page.Pagemargins but element names are different :-/
    """
    tagname = "pageMargins"

    l = Float()
    left = Alias('l')
    r = Float()
    right = Alias('r')
    t = Float()
    top = Alias('t')
    b = Float()
    bottom = Alias('b')
    header = Float()
    footer = Float()

    def __init__(self, l=0.75, r=0.75, t=1, b=1, header=0.5, footer=0.5):
        self.l = l
        self.r = r
        self.t = t
        self.b = b
        self.header = header
        self.footer = footer


class PrintSettings(Serialisable):

    tagname = "printSettings"

    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)

    __elements__ = ("headerFooter", "pageMargins", "pageMargins")

    def __init__(self,
                 headerFooter=None,
                 pageMargins=None,
                 pageSetup=None,
                ):
        self.headerFooter = headerFooter
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup

''',
    },
    'openpyxl.chart.radar_chart': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    Typed,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedSet
)

from ._chart import ChartBase
from .axis import TextAxis, NumericAxis
from .series import Series
from .label import DataLabelList


class RadarChart(ChartBase):

    tagname = "radarChart"

    radarStyle = NestedSet(values=(['standard', 'marker', 'filled']))
    type = Alias("radarStyle")
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    _series_type = "radar"

    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)

    __elements__ = ('radarStyle', 'varyColors', 'ser', 'dLbls', 'axId')

    def __init__(self,
                 radarStyle="standard",
                 varyColors=None,
                 ser=(),
                 dLbls=None,
                 extLst=None,
                 **kw
                ):
        self.radarStyle = radarStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        super().__init__(**kw)

""",
    },
    'openpyxl.chart.reader': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Read a chart
"""

def read_chart(chartspace):
    cs = chartspace
    plot = cs.chart.plotArea

    chart = plot._charts[0]
    chart._charts = plot._charts

    chart.title = cs.chart.title
    chart.display_blanks = cs.chart.dispBlanksAs
    chart.visible_cells_only = cs.chart.plotVisOnly
    chart.layout = plot.layout
    chart.legend = cs.chart.legend

    # 3d attributes
    chart.floor = cs.chart.floor
    chart.sideWall = cs.chart.sideWall
    chart.backWall = cs.chart.backWall
    chart.pivotSource = cs.pivotSource
    chart.pivotFormats = cs.chart.pivotFmts
    chart.idx_base = min((s.idx for s in chart.series), default=0)
    chart._reindex()

    # Border, fill, etc.
    chart.graphical_properties = cs.graphical_properties

    return chart

''',
    },
    'openpyxl.chart.reference': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from itertools import chain

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    MinMax,
    Typed,
    String,
    Strict,
)
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import (
    get_column_letter,
    range_to_tuple,
    quote_sheetname
)


class DummyWorksheet:


    def __init__(self, title):
        self.title = title


class Reference(Strict):

    """
    Normalise cell range references
    """

    min_row = MinMax(min=1, max=1000000, expected_type=int)
    max_row = MinMax(min=1, max=1000000, expected_type=int)
    min_col = MinMax(min=1, max=16384, expected_type=int)
    max_col = MinMax(min=1, max=16384, expected_type=int)
    range_string = String(allow_none=True)

    def __init__(self,
                 worksheet=None,
                 min_col=None,
                 min_row=None,
                 max_col=None,
                 max_row=None,
                 range_string=None
                 ):
        if range_string is not None:
            sheetname, boundaries = range_to_tuple(range_string)
            min_col, min_row, max_col, max_row = boundaries
            worksheet = DummyWorksheet(sheetname)

        self.worksheet = worksheet
        self.min_col = min_col
        self.min_row = min_row
        if max_col is None:
            max_col = min_col
        self.max_col = max_col
        if max_row is None:
            max_row = min_row
        self.max_row = max_row


    def __repr__(self):
        return str(self)


    def __str__(self):
        fmt = u"{0}!${1}${2}:${3}${4}"
        if (self.min_col == self.max_col
            and self.min_row == self.max_row):
            fmt = u"{0}!${1}${2}"
        return fmt.format(self.sheetname,
                          get_column_letter(self.min_col), self.min_row,
                          get_column_letter(self.max_col), self.max_row
                          )


    __str__ = __str__



    def __len__(self):
        if self.min_row == self.max_row:
            return 1 + self.max_col - self.min_col
        return 1 + self.max_row - self.min_row


    def __eq__(self, other):
        return str(self) == str(other)


    @property
    def rows(self):
        """
        Return all rows in the range
        """
        for row in range(self.min_row, self.max_row+1):
            yield Reference(self.worksheet, self.min_col, row, self.max_col, row)


    @property
    def cols(self):
        """
        Return all columns in the range
        """
        for col in range(self.min_col, self.max_col+1):
            yield Reference(self.worksheet, col, self.min_row, col, self.max_row)


    def pop(self):
        """
        Return and remove the first cell
        """
        cell = "{0}{1}".format(get_column_letter(self.min_col), self.min_row)
        if self.min_row == self.max_row:
            self.min_col += 1
        else:
            self.min_row += 1
        return cell


    @property
    def sheetname(self):
        return quote_sheetname(self.worksheet.title)

''',
    },
    'openpyxl.chart.scatter_chart': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Sequence,
    Alias
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedBool,
)

from ._chart import ChartBase
from .axis import NumericAxis, TextAxis
from .series import XYSeries
from .label import DataLabelList


class ScatterChart(ChartBase):

    tagname = "scatterChart"

    scatterStyle = NestedNoneSet(values=(['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker']))
    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=XYSeries, allow_none=True)
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias("dLbls")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=(NumericAxis, TextAxis))
    y_axis = Typed(expected_type=NumericAxis)

    _series_type = "scatter"

    __elements__ = ('scatterStyle', 'varyColors', 'ser', 'dLbls', 'axId',)

    def __init__(self,
                 scatterStyle=None,
                 varyColors=None,
                 ser=(),
                 dLbls=None,
                 extLst=None,
                 **kw
                ):
        self.scatterStyle = scatterStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.x_axis = NumericAxis(axId=10, crossAx=20)
        self.y_axis = NumericAxis(axId=20, crossAx=10)
        super().__init__(**kw)

""",
    },
    'openpyxl.chart.series': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
    Alias,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedBool,
    NestedNoneSet,
    NestedText,
)

from .shapes import GraphicalProperties
from .data_source import (
    AxDataSource,
    NumDataSource,
    NumRef,
    StrRef,
)
from .error_bar import ErrorBars
from .label import DataLabelList
from .marker import DataPoint, PictureOptions, Marker
from .trendline import Trendline

attribute_mapping = {
    'area': ('idx', 'order', 'tx', 'spPr', 'pictureOptions', 'dPt', 'dLbls', 'errBars',
             'trendline', 'cat', 'val',),
    'bar':('idx', 'order','tx', 'spPr', 'invertIfNegative', 'pictureOptions', 'dPt',
           'dLbls', 'trendline', 'errBars', 'cat', 'val', 'shape'),
    'bubble':('idx','order', 'tx', 'spPr', 'invertIfNegative', 'dPt', 'dLbls',
              'trendline', 'errBars', 'xVal', 'yVal', 'bubbleSize', 'bubble3D'),
    'line':('idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline',
            'errBars', 'cat', 'val', 'smooth'),
    'pie':('idx', 'order', 'tx', 'spPr', 'explosion', 'dPt', 'dLbls', 'cat', 'val'),
    'radar':('idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'cat', 'val'),
    'scatter':('idx', 'order', 'tx', 'spPr', 'marker', 'dPt', 'dLbls', 'trendline',
               'errBars', 'xVal', 'yVal', 'smooth'),
    'surface':('idx', 'order', 'tx', 'spPr', 'cat', 'val'),
                     }


class SeriesLabel(Serialisable):

    tagname = "tx"

    strRef = Typed(expected_type=StrRef, allow_none=True)
    v = NestedText(expected_type=str, allow_none=True)
    value = Alias('v')

    __elements__ = ('strRef', 'v')

    def __init__(self,
                 strRef=None,
                 v=None):
        self.strRef = strRef
        self.v = v


class Series(Serialisable):

    """
    Generic series object. Should not be instantiated directly.
    User the chart.Series factory instead.
    """

    tagname = "ser"

    idx = NestedInteger()
    order = NestedInteger()
    tx = Typed(expected_type=SeriesLabel, allow_none=True)
    title = Alias('tx')
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')

    # area chart
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    dPt = Sequence(expected_type=DataPoint, allow_none=True)
    data_points = Alias("dPt")
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    labels = Alias("dLbls")
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    identifiers = Alias("cat")
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    #bar chart
    invertIfNegative = NestedBool(allow_none=True)
    shape = NestedNoneSet(values=(['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax']))

    #bubble chart
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    bubbleSize = Typed(expected_type=NumDataSource, allow_none=True)
    zVal = Alias("bubbleSize")
    bubble3D = NestedBool(allow_none=True)

    #line chart
    marker = Typed(expected_type=Marker, allow_none=True)
    smooth = NestedBool(allow_none=True)

    #pie chart
    explosion = NestedInteger(allow_none=True)

    __elements__ = ()


    def __init__(self,
                 idx=0,
                 order=0,
                 tx=None,
                 spPr=None,
                 pictureOptions=None,
                 dPt=(),
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 invertIfNegative=None,
                 shape=None,
                 xVal=None,
                 yVal=None,
                 bubbleSize=None,
                 bubble3D=None,
                 marker=None,
                 smooth=None,
                 explosion=None,
                 extLst=None,
                ):
        self.idx = idx
        self.order = order
        self.tx = tx
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.invertIfNegative = invertIfNegative
        self.shape = shape
        self.xVal = xVal
        self.yVal = yVal
        self.bubbleSize = bubbleSize
        self.bubble3D = bubble3D
        if marker is None:
            marker = Marker()
        self.marker = marker
        self.smooth = smooth
        self.explosion = explosion


    def to_tree(self, tagname=None, idx=None):
        """The index can need rebasing"""
        if idx is not None:
            if self.order == self.idx:
                self.order = idx # rebase the order if the index has been rebased
            self.idx = idx
        return super().to_tree(tagname)


class XYSeries(Series):

    """Dedicated series for charts that have x and y series"""

    idx = Series.idx
    order = Series.order
    tx = Series.tx
    spPr = Series.spPr

    dPt = Series.dPt
    dLbls = Series.dLbls
    trendline = Series.trendline
    errBars = Series.errBars
    xVal = Series.xVal
    yVal = Series.yVal

    invertIfNegative = Series.invertIfNegative

    bubbleSize = Series.bubbleSize
    bubble3D = Series.bubble3D

    marker = Series.marker
    smooth = Series.smooth

''',
    },
    'openpyxl.chart.series_factory': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from .data_source import NumDataSource, NumRef, AxDataSource
from .reference import Reference
from .series import Series, XYSeries, SeriesLabel, StrRef
from  openpyxl.utils import rows_from_range, quote_sheetname


def SeriesFactory(values, xvalues=None, zvalues=None, title=None, title_from_data=False):
    """
    Convenience Factory for creating chart data series.
    """

    if not isinstance(values, Reference):
        values = Reference(range_string=values)

    if title_from_data:
        cell = values.pop()
        title = u"{0}!{1}".format(values.sheetname, cell)
        title = SeriesLabel(strRef=StrRef(title))
    elif title is not None:
        title = SeriesLabel(v=title)

    source = NumDataSource(numRef=NumRef(f=values))
    if xvalues is not None:
        if not isinstance(xvalues, Reference):
            xvalues = Reference(range_string=xvalues)
        series = XYSeries()
        series.yVal = source
        series.xVal = AxDataSource(numRef=NumRef(f=xvalues))
        if zvalues is not None:
            if not isinstance(zvalues, Reference):
                zvalues = Reference(range_string=zvalues)
            series.zVal = NumDataSource(NumRef(f=zvalues))
    else:
        series = Series()
        series.val = source

    if title is not None:
        series.title = title
    return series

''',
    },
    'openpyxl.chart.shapes': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Alias
)
from openpyxl.descriptors.nested import (
    EmptyTag
)
from openpyxl.drawing.colors import ColorChoiceDescriptor
from openpyxl.drawing.fill import *
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.geometry import (
    Shape3D,
    Scene3D,
    Transform2D,
    CustomGeometry2D,
    PresetGeometry2D,
)


class GraphicalProperties(Serialisable):

    """
    Somewhat vaguely 21.2.2.197 says this:

    This element specifies the formatting for the parent chart element. The
    custGeom, prstGeom, scene3d, and xfrm elements are not supported. The
    bwMode attribute is not supported.

    This doesn't leave much. And the element is used in different places.
    """

    tagname = "spPr"

    bwMode = NoneSet(values=(['clr', 'auto', 'gray', 'ltGray', 'invGray',
                          'grayWhite', 'blackGray', 'blackWhite', 'black', 'white', 'hidden']
                         )
                 )

    xfrm = Typed(expected_type=Transform2D, allow_none=True)
    transform = Alias('xfrm')
    custGeom = Typed(expected_type=CustomGeometry2D, allow_none=True) # either or
    prstGeom = Typed(expected_type=PresetGeometry2D, allow_none=True)

    # fills one of
    noFill = EmptyTag(namespace=DRAWING_NS)
    solidFill = ColorChoiceDescriptor()
    gradFill = Typed(expected_type=GradientFillProperties, allow_none=True)
    pattFill = Typed(expected_type=PatternFillProperties, allow_none=True)

    ln = Typed(expected_type=LineProperties, allow_none=True)
    line = Alias('ln')
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    sp3d = Typed(expected_type=Shape3D, allow_none=True)
    shape3D = Alias('sp3d')
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ('xfrm', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'pattFill',
                    'ln', 'scene3d', 'sp3d')

    def __init__(self,
                 bwMode=None,
                 xfrm=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 pattFill=None,
                 ln=None,
                 scene3d=None,
                 custGeom=None,
                 prstGeom=None,
                 sp3d=None,
                 extLst=None,
                ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.noFill = noFill
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.pattFill = pattFill
        if ln is None:
            ln = LineProperties()
        self.ln = ln
        self.custGeom = custGeom
        self.prstGeom = prstGeom
        self.scene3d = scene3d
        self.sp3d = sp3d

''',
    },
    'openpyxl.chart.stock_chart': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Sequence,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList

from ._chart import ChartBase
from .axis import TextAxis, NumericAxis, ChartLines
from .updown_bars import UpDownBars
from .label import DataLabelList
from .series import Series


class StockChart(ChartBase):

    tagname = "stockChart"

    ser = Sequence(expected_type=Series) #min 3, max4
    dLbls = Typed(expected_type=DataLabelList, allow_none=True)
    dataLabels = Alias('dLbls')
    dropLines = Typed(expected_type=ChartLines, allow_none=True)
    hiLowLines = Typed(expected_type=ChartLines, allow_none=True)
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)

    _series_type = "line"

    __elements__ = ('ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars',
                    'axId')

    def __init__(self,
                 ser=(),
                 dLbls=None,
                 dropLines=None,
                 hiLowLines=None,
                 upDownBars=None,
                 extLst=None,
                 **kw
                ):
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        super().__init__(**kw)

""",
    },
    'openpyxl.chart.surface_chart': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    Bool,
    Alias,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedBool,
)

from ._chart import ChartBase
from ._3d import _3DBase
from .axis import TextAxis, NumericAxis, SeriesAxis
from .shapes import GraphicalProperties
from .series import Series


class BandFormat(Serialisable):

    tagname = "bandFmt"

    idx = NestedInteger()
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")

    __elements__ = ('idx', 'spPr')

    def __init__(self,
                 idx=0,
                 spPr=None,
                ):
        self.idx = idx
        self.spPr = spPr


class BandFormatList(Serialisable):

    tagname = "bandFmts"

    bandFmt = Sequence(expected_type=BandFormat, allow_none=True)

    __elements__ = ('bandFmt',)

    def __init__(self,
                 bandFmt=(),
                ):
        self.bandFmt = bandFmt


class _SurfaceChartBase(ChartBase):

    wireframe = NestedBool(allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    bandFmts = Typed(expected_type=BandFormatList, allow_none=True)

    _series_type = "surface"

    __elements__ = ('wireframe', 'ser', 'bandFmts')

    def __init__(self,
                 wireframe=None,
                 ser=(),
                 bandFmts=None,
                 **kw
                ):
        self.wireframe = wireframe
        self.ser = ser
        self.bandFmts = bandFmts
        super().__init__(**kw)


class SurfaceChart3D(_SurfaceChartBase, _3DBase):

    tagname = "surface3DChart"

    wireframe = _SurfaceChartBase.wireframe
    ser = _SurfaceChartBase.ser
    bandFmts = _SurfaceChartBase.bandFmts

    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=TextAxis)
    y_axis = Typed(expected_type=NumericAxis)
    z_axis = Typed(expected_type=SeriesAxis)

    __elements__ = _SurfaceChartBase.__elements__ + ('axId',)

    def __init__(self, **kw):
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()
        super(SurfaceChart3D, self).__init__(**kw)


class SurfaceChart(SurfaceChart3D):

    tagname = "surfaceChart"

    wireframe = _SurfaceChartBase.wireframe
    ser = _SurfaceChartBase.ser
    bandFmts = _SurfaceChartBase.bandFmts

    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = SurfaceChart3D.__elements__

    def __init__(self, **kw):
        super().__init__(**kw)
        self.y_axis.delete = True
        self.view3D.x_rotation = 90
        self.view3D.y_rotation = 0
        self.view3D.perspective = False
        self.view3D.right_angle_axes = False

""",
    },
    'openpyxl.chart.text': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Alias,
    Sequence,
)


from openpyxl.drawing.text import (
    RichTextProperties,
    ListStyle,
    Paragraph,
)

from .data_source import StrRef


class RichText(Serialisable):

    """
    From the specification: 21.2.2.216

    This element specifies text formatting. The lstStyle element is not supported.
    """

    tagname = "rich"

    bodyPr = Typed(expected_type=RichTextProperties)
    properties = Alias("bodyPr")
    lstStyle = Typed(expected_type=ListStyle, allow_none=True)
    p = Sequence(expected_type=Paragraph)
    paragraphs = Alias('p')

    __elements__ = ("bodyPr", "lstStyle", "p")

    def __init__(self,
                 bodyPr=None,
                 lstStyle=None,
                 p=None,
                ):
        if bodyPr is None:
            bodyPr = RichTextProperties()
        self.bodyPr = bodyPr
        self.lstStyle = lstStyle
        if p is None:
            p = [Paragraph()]
        self.p = p


class Text(Serialisable):

    """
    The value can be either a cell reference or a text element
    If both are present then the reference will be used.
    """

    tagname = "tx"

    strRef = Typed(expected_type=StrRef, allow_none=True)
    rich = Typed(expected_type=RichText, allow_none=True)

    __elements__ = ("strRef", "rich")

    def __init__(self,
                 strRef=None,
                 rich=None
                 ):
        self.strRef = strRef
        if rich is None:
            rich = RichText()
        self.rich = rich


    def to_tree(self, tagname=None, idx=None, namespace=None):
        if self.strRef and self.rich:
            self.rich = None # can only have one
        return super().to_tree(tagname, idx, namespace)

''',
    },
    'openpyxl.chart.title': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Alias,
)

from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import NestedBool

from .text import Text, RichText
from .layout import Layout
from .shapes import GraphicalProperties

from openpyxl.drawing.text import (
    Paragraph,
    RegularTextRun,
    LineBreak,
    ParagraphProperties,
    CharacterProperties,
)


class Title(Serialisable):
    tagname = "title"

    tx = Typed(expected_type=Text, allow_none=True)
    text = Alias('tx')
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = NestedBool(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    txPr = Typed(expected_type=RichText, allow_none=True)
    body = Alias('txPr')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('tx', 'layout', 'overlay', 'spPr', 'txPr')

    def __init__(self,
                 tx=None,
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        if tx is None:
            tx = Text()
        self.tx = tx
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr



def title_maker(text):
    title = Title()
    paraprops = ParagraphProperties()
    paraprops.defRPr = CharacterProperties()
    paras = [Paragraph(r=[RegularTextRun(t=s)], pPr=paraprops) for s in text.split("\n")]

    title.tx.rich.paragraphs = paras
    return title


class TitleDescriptor(Typed):

    expected_type = Title
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, str):
            value = title_maker(value)
        super().__set__(instance, value)

""",
    },
    'openpyxl.chart.trendline': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Alias
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedFloat,
    NestedSet
)

from .data_source import NumFmt
from .shapes import GraphicalProperties
from .text import RichText, Text
from .layout import Layout


class TrendlineLabel(Serialisable):

    tagname = "trendlineLbl"

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Text, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias("spPr")
    txPr = Typed(expected_type=RichText, allow_none=True)
    textProperties = Alias("txPr")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('layout', 'tx', 'numFmt', 'spPr', 'txPr')

    def __init__(self,
                 layout=None,
                 tx=None,
                 numFmt=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.layout = layout
        self.tx = tx
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr


class Trendline(Serialisable):

    tagname = "trendline"

    name = String(allow_none=True)
    spPr = Typed(expected_type=GraphicalProperties, allow_none=True)
    graphicalProperties = Alias('spPr')
    trendlineType = NestedSet(values=(['exp', 'linear', 'log', 'movingAvg', 'poly', 'power']))
    order = NestedInteger(allow_none=True)
    period = NestedInteger(allow_none=True)
    forward = NestedFloat(allow_none=True)
    backward = NestedFloat(allow_none=True)
    intercept = NestedFloat(allow_none=True)
    dispRSqr = NestedBool(allow_none=True)
    dispEq = NestedBool(allow_none=True)
    trendlineLbl = Typed(expected_type=TrendlineLabel, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('spPr', 'trendlineType', 'order', 'period', 'forward',
                    'backward', 'intercept', 'dispRSqr', 'dispEq', 'trendlineLbl')

    def __init__(self,
                 name=None,
                 spPr=None,
                 trendlineType='linear',
                 order=None,
                 period=None,
                 forward=None,
                 backward=None,
                 intercept=None,
                 dispRSqr=None,
                 dispEq=None,
                 trendlineLbl=None,
                 extLst=None,
                ):
        self.name = name
        self.spPr = spPr
        self.trendlineType = trendlineType
        self.order = order
        self.period = period
        self.forward = forward
        self.backward = backward
        self.intercept = intercept
        self.dispRSqr = dispRSqr
        self.dispEq = dispEq
        self.trendlineLbl = trendlineLbl

""",
    },
    'openpyxl.chart.updown_bars': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import Typed
from openpyxl.descriptors.excel import ExtensionList

from .shapes import GraphicalProperties
from .axis import ChartLines
from .descriptors import NestedGapAmount


class UpDownBars(Serialisable):

    tagname = "upbars"

    gapWidth = NestedGapAmount()
    upBars = Typed(expected_type=ChartLines, allow_none=True)
    downBars = Typed(expected_type=ChartLines, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'upBars', 'downBars')

    def __init__(self,
                 gapWidth=150,
                 upBars=None,
                 downBars=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars

""",
    },
    'openpyxl.chartsheet': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .chartsheet import Chartsheet

""",
    },
    'openpyxl.chartsheet.chartsheet': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors import Typed, Set, Alias
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.drawing.spreadsheet_drawing import (
    AbsoluteAnchor,
    SpreadsheetDrawing,
)
from openpyxl.worksheet.page import (
    PageMargins,
    PrintPageSetup
)
from openpyxl.worksheet.drawing import Drawing
from openpyxl.worksheet.header_footer import HeaderFooter
from openpyxl.workbook.child import _WorkbookChild
from openpyxl.xml.constants import SHEET_MAIN_NS, REL_NS

from .relation import DrawingHF, SheetBackgroundPicture
from .properties import ChartsheetProperties
from .protection import ChartsheetProtection
from .views import ChartsheetViewList
from .custom import CustomChartsheetViews
from .publish import WebPublishItems


class Chartsheet(_WorkbookChild, Serialisable):

    tagname = "chartsheet"
    _default_title = "Chart"
    _rel_type = "chartsheet"
    _path = "/xl/chartsheets/sheet{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"

    sheetPr = Typed(expected_type=ChartsheetProperties, allow_none=True)
    sheetViews = Typed(expected_type=ChartsheetViewList)
    sheetProtection = Typed(expected_type=ChartsheetProtection, allow_none=True)
    customSheetViews = Typed(expected_type=CustomChartsheetViews, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)
    drawing = Typed(expected_type=Drawing, allow_none=True)
    drawingHF = Typed(expected_type=DrawingHF, allow_none=True)
    picture = Typed(expected_type=SheetBackgroundPicture, allow_none=True)
    webPublishItems = Typed(expected_type=WebPublishItems, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    sheet_state = Set(values=('visible', 'hidden', 'veryHidden'))
    headerFooter = Typed(expected_type=HeaderFooter)
    HeaderFooter = Alias('headerFooter')

    __elements__ = (
        'sheetPr', 'sheetViews', 'sheetProtection', 'customSheetViews',
        'pageMargins', 'pageSetup', 'headerFooter', 'drawing', 'drawingHF',
        'picture', 'webPublishItems')

    __attrs__ = ()

    def __init__(self,
                 sheetPr=None,
                 sheetViews=None,
                 sheetProtection=None,
                 customSheetViews=None,
                 pageMargins=None,
                 pageSetup=None,
                 headerFooter=None,
                 drawing=None,
                 drawingHF=None,
                 picture=None,
                 webPublishItems=None,
                 extLst=None,
                 parent=None,
                 title="",
                 sheet_state='visible',
                 ):
        super().__init__(parent, title)
        self._charts = []
        self.sheetPr = sheetPr
        if sheetViews is None:
            sheetViews = ChartsheetViewList()
        self.sheetViews = sheetViews
        self.sheetProtection = sheetProtection
        self.customSheetViews = customSheetViews
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        if headerFooter is not None:
            self.headerFooter = headerFooter
        self.drawing = Drawing("rId1")
        self.drawingHF = drawingHF
        self.picture = picture
        self.webPublishItems = webPublishItems
        self.sheet_state = sheet_state


    def add_chart(self, chart):
        chart.anchor = AbsoluteAnchor()
        self._charts.append(chart)


    def to_tree(self):
        self._drawing = SpreadsheetDrawing()
        self._drawing.charts = self._charts
        tree = super().to_tree()
        if not self.headerFooter:
            el = tree.find('headerFooter')
            tree.remove(el)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree

""",
    },
    'openpyxl.chartsheet.custom': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.worksheet.header_footer import HeaderFooter

from openpyxl.descriptors import (
    Bool,
    Integer,
    Set,
    Typed,
    Sequence
)
from openpyxl.descriptors.excel import Guid
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.worksheet.page import (
    PageMargins,
    PrintPageSetup
)


class CustomChartsheetView(Serialisable):
    tagname = "customSheetView"

    guid = Guid()
    scale = Integer()
    state = Set(values=(['visible', 'hidden', 'veryHidden']))
    zoomToFit = Bool(allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)
    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)

    __elements__ = ('pageMargins', 'pageSetup', 'headerFooter')

    def __init__(self,
                 guid=None,
                 scale=None,
                 state='visible',
                 zoomToFit=None,
                 pageMargins=None,
                 pageSetup=None,
                 headerFooter=None,
                 ):
        self.guid = guid
        self.scale = scale
        self.state = state
        self.zoomToFit = zoomToFit
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.headerFooter = headerFooter


class CustomChartsheetViews(Serialisable):
    tagname = "customSheetViews"

    customSheetView = Sequence(expected_type=CustomChartsheetView, allow_none=True)

    __elements__ = ('customSheetView',)

    def __init__(self,
                 customSheetView=None,
                 ):
        self.customSheetView = customSheetView

""",
    },
    'openpyxl.chartsheet.properties': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Bool,
    String,
    Typed
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles import Color


class ChartsheetProperties(Serialisable):
    tagname = "sheetPr"

    published = Bool(allow_none=True)
    codeName = String(allow_none=True)
    tabColor = Typed(expected_type=Color, allow_none=True)

    __elements__ = ('tabColor',)

    def __init__(self,
                 published=None,
                 codeName=None,
                 tabColor=None,
                 ):
        self.published = published
        self.codeName = codeName
        self.tabColor = tabColor

""",
    },
    'openpyxl.chartsheet.protection': {
        'is_package': False,
        'source': r"""
import hashlib

from openpyxl.descriptors import (Bool, Integer, String)
from openpyxl.descriptors.excel import Base64Binary
from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.worksheet.protection import (
    hash_password,
    _Protected
)


class ChartsheetProtection(Serialisable, _Protected):
    tagname = "sheetProtection"

    algorithmName = String(allow_none=True)
    hashValue = Base64Binary(allow_none=True)
    saltValue = Base64Binary(allow_none=True)
    spinCount = Integer(allow_none=True)
    content = Bool(allow_none=True)
    objects = Bool(allow_none=True)

    __attrs__ = ("content", "objects", "password", "hashValue", "spinCount", "saltValue", "algorithmName")

    def __init__(self,
                 content=None,
                 objects=None,
                 hashValue=None,
                 spinCount=None,
                 saltValue=None,
                 algorithmName=None,
                 password=None,
                 ):
        self.content = content
        self.objects = objects
        self.hashValue = hashValue
        self.spinCount = spinCount
        self.saltValue = saltValue
        self.algorithmName = algorithmName
        if password is not None:
            self.password = password

""",
    },
    'openpyxl.chartsheet.publish': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Bool,
    Integer,
    String,
    Set,
    Sequence
)
from openpyxl.descriptors.serialisable import Serialisable


class WebPublishItem(Serialisable):
    tagname = "webPublishItem"

    id = Integer()
    divId = String()
    sourceType = Set(values=(['sheet', 'printArea', 'autoFilter', 'range', 'chart', 'pivotTable', 'query', 'label']))
    sourceRef = String()
    sourceObject = String(allow_none=True)
    destinationFile = String()
    title = String(allow_none=True)
    autoRepublish = Bool(allow_none=True)

    def __init__(self,
                 id=None,
                 divId=None,
                 sourceType=None,
                 sourceRef=None,
                 sourceObject=None,
                 destinationFile=None,
                 title=None,
                 autoRepublish=None,
                 ):
        self.id = id
        self.divId = divId
        self.sourceType = sourceType
        self.sourceRef = sourceRef
        self.sourceObject = sourceObject
        self.destinationFile = destinationFile
        self.title = title
        self.autoRepublish = autoRepublish


class WebPublishItems(Serialisable):
    tagname = "WebPublishItems"

    count = Integer(allow_none=True)
    webPublishItem = Sequence(expected_type=WebPublishItem, )

    __elements__ = ('webPublishItem',)

    def __init__(self,
                 count=None,
                 webPublishItem=None,
                 ):
        self.count = len(webPublishItem)
        self.webPublishItem = webPublishItem

""",
    },
    'openpyxl.chartsheet.relation': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Integer,
    Alias
)
from openpyxl.descriptors.excel import Relation
from openpyxl.descriptors.serialisable import Serialisable


class SheetBackgroundPicture(Serialisable):
    tagname = "picture"
    id = Relation()

    def __init__(self, id):
        self.id = id


class DrawingHF(Serialisable):
    id = Relation()
    lho = Integer(allow_none=True)
    leftHeaderOddPages = Alias('lho')
    lhe = Integer(allow_none=True)
    leftHeaderEvenPages = Alias('lhe')
    lhf = Integer(allow_none=True)
    leftHeaderFirstPage = Alias('lhf')
    cho = Integer(allow_none=True)
    centerHeaderOddPages = Alias('cho')
    che = Integer(allow_none=True)
    centerHeaderEvenPages = Alias('che')
    chf = Integer(allow_none=True)
    centerHeaderFirstPage = Alias('chf')
    rho = Integer(allow_none=True)
    rightHeaderOddPages = Alias('rho')
    rhe = Integer(allow_none=True)
    rightHeaderEvenPages = Alias('rhe')
    rhf = Integer(allow_none=True)
    rightHeaderFirstPage = Alias('rhf')
    lfo = Integer(allow_none=True)
    leftFooterOddPages = Alias('lfo')
    lfe = Integer(allow_none=True)
    leftFooterEvenPages = Alias('lfe')
    lff = Integer(allow_none=True)
    leftFooterFirstPage = Alias('lff')
    cfo = Integer(allow_none=True)
    centerFooterOddPages = Alias('cfo')
    cfe = Integer(allow_none=True)
    centerFooterEvenPages = Alias('cfe')
    cff = Integer(allow_none=True)
    centerFooterFirstPage = Alias('cff')
    rfo = Integer(allow_none=True)
    rightFooterOddPages = Alias('rfo')
    rfe = Integer(allow_none=True)
    rightFooterEvenPages = Alias('rfe')
    rff = Integer(allow_none=True)
    rightFooterFirstPage = Alias('rff')

    def __init__(self,
                 id=None,
                 lho=None,
                 lhe=None,
                 lhf=None,
                 cho=None,
                 che=None,
                 chf=None,
                 rho=None,
                 rhe=None,
                 rhf=None,
                 lfo=None,
                 lfe=None,
                 lff=None,
                 cfo=None,
                 cfe=None,
                 cff=None,
                 rfo=None,
                 rfe=None,
                 rff=None,
                 ):
        self.id = id
        self.lho = lho
        self.lhe = lhe
        self.lhf = lhf
        self.cho = cho
        self.che = che
        self.chf = chf
        self.rho = rho
        self.rhe = rhe
        self.rhf = rhf
        self.lfo = lfo
        self.lfe = lfe
        self.lff = lff
        self.cfo = cfo
        self.cfe = cfe
        self.cff = cff
        self.rfo = rfo
        self.rfe = rfe
        self.rff = rff

""",
    },
    'openpyxl.chartsheet.views': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Bool,
    Integer,
    Typed,
    Sequence
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable


class ChartsheetView(Serialisable):
    tagname = "sheetView"

    tabSelected = Bool(allow_none=True)
    zoomScale = Integer(allow_none=True)
    workbookViewId = Integer()
    zoomToFit = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 tabSelected=None,
                 zoomScale=None,
                 workbookViewId=0,
                 zoomToFit=True,
                 extLst=None,
                 ):
        self.tabSelected = tabSelected
        self.zoomScale = zoomScale
        self.workbookViewId = workbookViewId
        self.zoomToFit = zoomToFit


class ChartsheetViewList(Serialisable):
    tagname = "sheetViews"

    sheetView = Sequence(expected_type=ChartsheetView, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('sheetView',)

    def __init__(self,
                 sheetView=None,
                 extLst=None,
                 ):
        if sheetView is None:
            sheetView = [ChartsheetView()]
        self.sheetView = sheetView

""",
    },
    'openpyxl.comments': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from .comments import Comment

""",
    },
    'openpyxl.comments.author': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    Alias
)


class AuthorList(Serialisable):

    tagname = "authors"

    author = Sequence(expected_type=str)
    authors = Alias("author")

    def __init__(self,
                 author=(),
                ):
        self.author = author

""",
    },
    'openpyxl.comments.comment_sheet': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

## Incomplete!
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    Set,
    String,
    Bool,
)
from openpyxl.descriptors.excel import Guid, ExtensionList
from openpyxl.descriptors.sequence import NestedSequence

from openpyxl.utils.indexed_list import IndexedList
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.cell.text import Text
from .author import AuthorList
from .comments import Comment
from .shape_writer import ShapeWriter


class Properties(Serialisable):

    locked = Bool(allow_none=True)
    defaultSize = Bool(allow_none=True)
    _print = Bool(allow_none=True)
    disabled = Bool(allow_none=True)
    uiObject = Bool(allow_none=True)
    autoFill = Bool(allow_none=True)
    autoLine = Bool(allow_none=True)
    altText = String(allow_none=True)
    textHAlign = Set(values=(['left', 'center', 'right', 'justify', 'distributed']))
    textVAlign = Set(values=(['top', 'center', 'bottom', 'justify', 'distributed']))
    lockText = Bool(allow_none=True)
    justLastX = Bool(allow_none=True)
    autoScale = Bool(allow_none=True)
    rowHidden = Bool(allow_none=True)
    colHidden = Bool(allow_none=True)
    # anchor = Typed(expected_type=ObjectAnchor, )

    __elements__ = ('anchor',)

    def __init__(self,
                 locked=None,
                 defaultSize=None,
                 _print=None,
                 disabled=None,
                 uiObject=None,
                 autoFill=None,
                 autoLine=None,
                 altText=None,
                 textHAlign=None,
                 textVAlign=None,
                 lockText=None,
                 justLastX=None,
                 autoScale=None,
                 rowHidden=None,
                 colHidden=None,
                 anchor=None,
                ):
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.altText = altText
        self.textHAlign = textHAlign
        self.textVAlign = textVAlign
        self.lockText = lockText
        self.justLastX = justLastX
        self.autoScale = autoScale
        self.rowHidden = rowHidden
        self.colHidden = colHidden
        self.anchor = anchor


class CommentRecord(Serialisable):

    tagname = "comment"

    ref = String()
    authorId = Integer()
    guid = Guid(allow_none=True)
    shapeId = Integer(allow_none=True)
    text = Typed(expected_type=Text)
    commentPr = Typed(expected_type=Properties, allow_none=True)
    author = String(allow_none=True)

    __elements__ = ('text', 'commentPr')
    __attrs__ = ('ref', 'authorId', 'guid', 'shapeId')

    def __init__(self,
                 ref="",
                 authorId=0,
                 guid=None,
                 shapeId=0,
                 text=None,
                 commentPr=None,
                 author=None,
                 height=79,
                 width=144
                ):
        self.ref = ref
        self.authorId = authorId
        self.guid = guid
        self.shapeId = shapeId
        if text is None:
            text = Text()
        self.text = text
        self.commentPr = commentPr
        self.author = author
        self.height = height
        self.width = width


    @classmethod
    def from_cell(cls, cell):
        """
        Class method to convert cell comment
        """
        comment = cell._comment
        ref = cell.coordinate
        self = cls(ref=ref, author=comment.author)
        self.text.t = comment.content
        self.height = comment.height
        self.width = comment.width
        return self


    @property
    def content(self):
        """
        Remove all inline formatting and stuff
        """
        return self.text.content


class CommentSheet(Serialisable):

    tagname = "comments"

    authors = Typed(expected_type=AuthorList)
    commentList = NestedSequence(expected_type=CommentRecord, count=0)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    _id = None
    _path = "/xl/comments/comment{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
    _rel_type = "comments"
    _rel_id = None

    __elements__ = ('authors', 'commentList')

    def __init__(self,
                 authors=None,
                 commentList=None,
                 extLst=None,
                ):
        self.authors = authors
        self.commentList = commentList


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def comments(self):
        """
        Return a dictionary of comments keyed by coord
        """
        authors = self.authors.author

        for c in self.commentList:
            yield c.ref, Comment(c.content, authors[c.authorId], c.height, c.width)


    @classmethod
    def from_comments(cls, comments):
        """
        Create a comment sheet from a list of comments for a particular worksheet
        """
        authors = IndexedList()

        # dedupe authors and get indexes
        for comment in comments:
            comment.authorId = authors.add(comment.author)

        return cls(authors=AuthorList(authors), commentList=comments)


    def write_shapes(self, vml=None):
        """
        Create the VML for comments
        """
        sw = ShapeWriter(self.comments)
        return sw.write(vml)


    @property
    def path(self):
        """
        Return path within the archive
        """
        return self._path.format(self._id)

''',
    },
    'openpyxl.comments.comments': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


class Comment:

    _parent = None

    def __init__(self, text, author, height=79, width=144):
        self.content = text
        self.author = author
        self.height = height
        self.width = width


    @property
    def parent(self):
        return self._parent


    def __eq__(self, other):
        return (
            self.content == other.content
            and self.author == other.author
        )

    def __repr__(self):
        return "Comment: {0} by {1}".format(self.content, self.author)


    def __copy__(self):
        """Create a detached copy of this comment."""
        clone = self.__class__(self.content, self.author, self.height, self.width)
        return clone


    def bind(self, cell):
        """
        Bind comment to a particular cell
        """
        if cell is not None and self._parent is not None and self._parent != cell:
            fmt = "Comment already assigned to {0} in worksheet {1}. Cannot assign a comment to more than one cell"
            raise AttributeError(fmt.format(cell.coordinate, cell.parent.title))
        self._parent = cell


    def unbind(self):
        """
        Unbind a comment from a cell
        """
        self._parent = None


    @property
    def text(self):
        """
        Any comment text stripped of all formatting.
        """
        return self.content

    @text.setter
    def text(self, value):
        self.content = value

''',
    },
    'openpyxl.comments.shape_writer': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.xml.functions import (
    Element,
    SubElement,
    tostring,
)

from openpyxl.utils import coordinate_to_tuple

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"


class ShapeWriter:
    """
    Create VML for comments
    """

    vml = None
    vml_path = None


    def __init__(self, comments):
        self.comments = comments


    def add_comment_shapetype(self, root):
        shape_layout = SubElement(root, "{%s}shapelayout" % officens,
                                  {"{%s}ext" % vmlns: "edit"})
        SubElement(shape_layout,
                   "{%s}idmap" % officens,
                   {"{%s}ext" % vmlns: "edit", "data": "1"})
        shape_type = SubElement(root,
                                "{%s}shapetype" % vmlns,
                                {"id": "_x0000_t202",
                                 "coordsize": "21600,21600",
                                 "{%s}spt" % officens: "202",
                                 "path": "m,l,21600r21600,l21600,xe"})
        SubElement(shape_type, "{%s}stroke" % vmlns, {"joinstyle": "miter"})
        SubElement(shape_type,
                   "{%s}path" % vmlns,
                   {"gradientshapeok": "t",
                    "{%s}connecttype" % officens: "rect"})


    def add_comment_shape(self, root, idx, coord, height, width):
        row, col = coordinate_to_tuple(coord)
        row -= 1
        col -= 1
        shape = _shape_factory(row, col, height, width)

        shape.set('id', "_x0000_s%04d" % idx)
        root.append(shape)


    def write(self, root):

        if not hasattr(root, "findall"):
            root = Element("xml")

        # Remove any existing comment shapes
        comments = root.findall("{%s}shape[@type='#_x0000_t202']" % vmlns)
        for c in comments:
            root.remove(c)

        # check whether comments shape type already exists
        shape_types = root.find("{%s}shapetype[@id='_x0000_t202']" % vmlns)
        if shape_types is None:
            self.add_comment_shapetype(root)

        for idx, (coord, comment) in enumerate(self.comments, 1026):
            self.add_comment_shape(root, idx, coord, comment.height, comment.width)

        return tostring(root)


def _shape_factory(row, column, height, width):
    style = ("position:absolute; "
             "margin-left:59.25pt;"
             "margin-top:1.5pt;"
             "width:{width}px;"
             "height:{height}px;"
             "z-index:1;"
             "visibility:hidden").format(height=height,
                                         width=width)
    attrs = {
        "type": "#_x0000_t202",
        "style": style,
        "fillcolor": "#ffffe1",
        "{%s}insetmode" % officens: "auto"
    }
    shape = Element("{%s}shape" % vmlns, attrs)

    SubElement(shape, "{%s}fill" % vmlns,
               {"color2": "#ffffe1"})
    SubElement(shape, "{%s}shadow" % vmlns,
               {"color": "black", "obscured": "t"})
    SubElement(shape, "{%s}path" % vmlns,
               {"{%s}connecttype" % officens: "none"})
    textbox = SubElement(shape, "{%s}textbox" % vmlns,
                         {"style": "mso-direction-alt:auto"})
    SubElement(textbox, "div", {"style": "text-align:left"})
    client_data = SubElement(shape, "{%s}ClientData" % excelns,
                             {"ObjectType": "Note"})
    SubElement(client_data, "{%s}MoveWithCells" % excelns)
    SubElement(client_data, "{%s}SizeWithCells" % excelns)
    SubElement(client_data, "{%s}AutoFill" % excelns).text = "False"
    SubElement(client_data, "{%s}Row" % excelns).text = str(row)
    SubElement(client_data, "{%s}Column" % excelns).text = str(column)
    return shape

''',
    },
    'openpyxl.compat': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .numbers import NUMERIC_TYPES
from .strings import safe_string

import warnings
from functools import wraps
import inspect


class DummyCode:

    pass


# from https://github.com/tantale/deprecated/blob/master/deprecated/__init__.py
# with an enhancement to update docstrings of deprecated functions
string_types = (type(b''), type(u''))
def deprecated(reason):

    if isinstance(reason, string_types):

        def decorator(func1):

            if inspect.isclass(func1):
                fmt1 = "Call to deprecated class {name} ({reason})."
            else:
                fmt1 = "Call to deprecated function {name} ({reason})."

            @wraps(func1)
            def new_func1(*args, **kwargs):
                #warnings.simplefilter('default', DeprecationWarning)
                warnings.warn(
                    fmt1.format(name=func1.__name__, reason=reason),
                    category=DeprecationWarning,
                    stacklevel=2
                )
                return func1(*args, **kwargs)

            # Enhance docstring with a deprecation note
            deprecationNote = "\n\n.. note::\n    Deprecated: " + reason
            if new_func1.__doc__:
                new_func1.__doc__ += deprecationNote
            else:
                new_func1.__doc__ = deprecationNote
            return new_func1

        return decorator

    elif inspect.isclass(reason) or inspect.isfunction(reason):
        raise TypeError("Reason for deprecation must be supplied")

    else:
        raise TypeError(repr(type(reason)))

""",
    },
    'openpyxl.compat.abc': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


try:
    from abc import ABC
except ImportError:
    from abc import ABCMeta
    ABC = ABCMeta('ABC', (object, ), {})

""",
    },
    'openpyxl.compat.numbers': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from decimal import Decimal

NUMERIC_TYPES = (int, float, Decimal)


try:
    import numpy
    NUMPY = True
except ImportError:
    NUMPY = False


if NUMPY:
    NUMERIC_TYPES = NUMERIC_TYPES + (numpy.short,
                                     numpy.ushort,
                                     numpy.intc,
                                     numpy.uintc,
                                     numpy.int_,
                                     numpy.uint,
                                     numpy.longlong,
                                     numpy.ulonglong,
                                     numpy.half,
                                     numpy.float16,
                                     numpy.single,
                                     numpy.double,
                                     numpy.longdouble,
                                     numpy.int8,
                                     numpy.int16,
                                     numpy.int32,
                                     numpy.int64,
                                     numpy.uint8,
                                     numpy.uint16,
                                     numpy.uint32,
                                     numpy.uint64,
                                     numpy.intp,
                                     numpy.uintp,
                                     numpy.float32,
                                     numpy.float64,
                                     numpy.bool_,
                                     numpy.floating,
                                     numpy.integer)

""",
    },
    'openpyxl.compat.product': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
math.prod equivalent for < Python 3.8
"""

import functools
import operator

def product(sequence):
    return functools.reduce(operator.mul, sequence)


try:
    from math import prod
except ImportError:
    prod = product

''',
    },
    'openpyxl.compat.singleton': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import weakref


class Singleton(type):
    """
    Singleton metaclass
    Based on Python Cookbook 3rd Edition Recipe 9.13
    Only one instance of a class can exist. Does not work with __slots__
    """

    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)
        self.__instance = None

    def __call__(self, *args, **kw):
        if self.__instance is None:
            self.__instance = super().__call__(*args, **kw)
        return self.__instance


class Cached(type):
    """
    Caching metaclass
    Child classes will only create new instances of themselves if
    one doesn't already exist. Does not work with __slots__
    """

    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)
        self.__cache = weakref.WeakValueDictionary()

    def __call__(self, *args):
        if args in self.__cache:
            return self.__cache[args]

        obj = super().__call__(*args)
        self.__cache[args] = obj
        return obj

''',
    },
    'openpyxl.compat.strings': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from datetime import datetime
from math import isnan, isinf
import sys

VER = sys.version_info

from .numbers import NUMERIC_TYPES


def safe_string(value):
    """Safely and consistently format numeric values"""
    if isinstance(value, NUMERIC_TYPES):
        if isnan(value) or isinf(value):
            value = ""
        else:
            value = "%.16g" % value
    elif value is None:
        value = "none"
    elif isinstance(value, datetime):
        value = value.isoformat()
    elif not isinstance(value, str):
        value = str(value)
    return value

''',
    },
    'openpyxl.descriptors': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .base import *
from .sequence import Sequence


class MetaStrict(type):

    def __new__(cls, clsname, bases, methods):
        for k, v in methods.items():
            if isinstance(v, Descriptor):
                v.name = k
        return type.__new__(cls, clsname, bases, methods)


class Strict(metaclass=MetaStrict):

    pass


class MetaSerialisable(type):

    def __new__(cls, clsname, bases, methods):
        attrs = []
        nested = []
        elements = []
        namespaced = []
        for k, v in methods.items():
            if isinstance(v, Descriptor):
                ns= getattr(v, 'namespace', None)
                if ns:
                    namespaced.append((k, "{%s}%s" % (ns, k)))
                if getattr(v, 'nested', False):
                    nested.append(k)
                    elements.append(k)
                elif isinstance(v, Sequence):
                    elements.append(k)
                elif isinstance(v, Typed):
                    if hasattr(v.expected_type, 'to_tree'):
                        elements.append(k)
                    elif isinstance(v.expected_type, tuple):
                        if any((hasattr(el, "to_tree") for el in v.expected_type)):
                            # don't bind elements as attrs
                            continue
                    else:
                        attrs.append(k)
                else:
                    if not isinstance(v, Alias):
                        attrs.append(k)

        if methods.get('__attrs__') is None:
            methods['__attrs__'] = tuple(attrs)
        methods['__namespaced__'] = tuple(namespaced)
        if methods.get('__nested__') is None:
            methods['__nested__'] = tuple(sorted(nested))
        if methods.get('__elements__') is None:
            methods['__elements__'] = tuple(sorted(elements))
        return MetaStrict.__new__(cls, clsname, bases, methods)

""",
    },
    'openpyxl.descriptors.base': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


"""
Based on Python Cookbook 3rd Edition, 8.13
http://chimera.labs.oreilly.com/books/1230000000393/ch08.html#_discussiuncion_130
"""

import datetime
import re

from openpyxl import DEBUG
from openpyxl.utils.datetime import from_ISO8601

from .namespace import namespaced

class Descriptor:

    def __init__(self, name=None, **kw):
        self.name = name
        for k, v in kw.items():
            setattr(self, k, v)

    def __set__(self, instance, value):
        instance.__dict__[self.name] = value


class Typed(Descriptor):
    """Values must of a particular type"""

    expected_type = type(None)
    allow_none = False
    nested = False

    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)
        self.__doc__ = f"Values must be of type {self.expected_type}"

    def __set__(self, instance, value):
        if not isinstance(value, self.expected_type):
            if (not self.allow_none
                or (self.allow_none and value is not None)):
                msg = f"{instance.__class__}.{self.name} should be {self.expected_type} but value is {type(value)}"
                if DEBUG:
                    msg = f"{instance.__class__}.{self.name} should be {self.expected_type} but {value} is {type(value)}"
                raise TypeError(msg)
        super().__set__(instance, value)

    def __repr__(self):
        return  self.__doc__


def _convert(expected_type, value):
    """
    Check value is of or can be converted to expected type.
    """
    if not isinstance(value, expected_type):
        try:
            value = expected_type(value)
        except:
            raise TypeError('expected ' + str(expected_type))
    return value


class Convertible(Typed):
    """Values must be convertible to a particular type"""

    def __set__(self, instance, value):
        if ((self.allow_none and value is not None)
            or not self.allow_none):
            value = _convert(self.expected_type, value)
        super().__set__(instance, value)


class Max(Convertible):
    """Values must be less than a `max` value"""

    expected_type = float
    allow_none = False

    def __init__(self, **kw):
        if 'max' not in kw and not hasattr(self, 'max'):
            raise TypeError('missing max value')
        super().__init__(**kw)

    def __set__(self, instance, value):
        if ((self.allow_none and value is not None)
            or not self.allow_none):
            value = _convert(self.expected_type, value)
            if value > self.max:
                raise ValueError('Max value is {0}'.format(self.max))
        super().__set__(instance, value)


class Min(Convertible):
    """Values must be greater than a `min` value"""

    expected_type = float
    allow_none = False

    def __init__(self, **kw):
        if 'min' not in kw and not hasattr(self, 'min'):
            raise TypeError('missing min value')
        super().__init__(**kw)

    def __set__(self, instance, value):
        if ((self.allow_none and value is not None)
            or not self.allow_none):
            value = _convert(self.expected_type, value)
            if value < self.min:
                raise ValueError('Min value is {0}'.format(self.min))
        super().__set__(instance, value)


class MinMax(Min, Max):
    """Values must be greater than `min` value and less than a `max` one"""
    pass


class Set(Descriptor):
    """Value can only be from a set of know values"""

    def __init__(self, name=None, **kw):
        if not 'values' in kw:
            raise TypeError("missing set of values")
        kw['values'] = set(kw['values'])
        super().__init__(name, **kw)
        self.__doc__ = "Value must be one of {0}".format(self.values)

    def __set__(self, instance, value):
        if value not in self.values:
            raise ValueError(self.__doc__)
        super().__set__(instance, value)


class NoneSet(Set):

    """'none' will be treated as None"""

    def __init__(self, name=None, **kw):
        super().__init__(name, **kw)
        self.values.add(None)

    def __set__(self, instance, value):
        if value == 'none':
            value = None
        super().__set__(instance, value)


class Integer(Convertible):

    expected_type = int


class Float(Convertible):

    expected_type = float


class Bool(Convertible):

    expected_type = bool

    def __set__(self, instance, value):
        if isinstance(value, str):
            if value in ('false', 'f', '0'):
                value = False
        super().__set__(instance, value)


class String(Typed):

    expected_type = str


class Text(String, Convertible):

    pass


class ASCII(Typed):

    expected_type = bytes


class Tuple(Typed):

    expected_type = tuple


class Length(Descriptor):

    def __init__(self, name=None, **kw):
        if "length" not in kw:
            raise TypeError("value length must be supplied")
        super().__init__(**kw)


    def __set__(self, instance, value):
        if len(value) != self.length:
            raise ValueError("Value must be length {0}".format(self.length))
        super().__set__(instance, value)


class Default(Typed):
    """
    When called returns an instance of the expected type.
    Additional default values can be passed in to the descriptor
    """

    def __init__(self, name=None, **kw):
        if "defaults" not in kw:
            kw['defaults'] = {}
        super().__init__(**kw)

    def __call__(self):
        return self.expected_type()


class Alias(Descriptor):
    """
    Aliases can be used when either the desired attribute name is not allowed
    or confusing in Python (eg. "type") or a more descriptive name is desired
    (eg. "underline" for "u")
    """

    def __init__(self, alias):
        self.alias = alias

    def __set__(self, instance, value):
        setattr(instance, self.alias, value)

    def __get__(self, instance, cls):
        return getattr(instance, self.alias)


class MatchPattern(Descriptor):
    """Values must match a regex pattern """
    allow_none = False

    def __init__(self, name=None, **kw):
        if 'pattern' not in kw and not hasattr(self, 'pattern'):
            raise TypeError('missing pattern value')

        super().__init__(name, **kw)
        self.test_pattern = re.compile(self.pattern, re.VERBOSE)


    def __set__(self, instance, value):

        if value is None and not self.allow_none:
            raise ValueError("Value must not be none")

        if ((self.allow_none and value is not None)
            or not self.allow_none):
            if not self.test_pattern.match(value):
                raise ValueError('Value does not match pattern {0}'.format(self.pattern))

        super().__set__(instance, value)


class DateTime(Typed):

    expected_type = datetime.datetime

    def __set__(self, instance, value):
        if value is not None and isinstance(value, str):
            try:
                value = from_ISO8601(value)
            except ValueError:
                raise ValueError("Value must be ISO datetime format")
        super().__set__(instance, value)

''',
    },
    'openpyxl.descriptors.container': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Utility list for top level containers that contain one type of element

Provides the necessary API to read and write XML
"""

from openpyxl.xml.functions import Element


class ElementList(list):


    @property
    def tagname(self):
        raise NotImplementedError


    @property
    def expected_type(self):
        raise NotImplementedError


    @classmethod
    def from_tree(cls, tree):
        l = [cls.expected_type.from_tree(el) for el in tree]
        return cls(l)


    def to_tree(self):
        container = Element(self.tagname)
        for el in self:
            container.append(el.to_tree())
        return container


    def append(self, value):
        if not isinstance(value, self.expected_type):
            raise TypeError(f"Value must of type {self.expected_type} {type(value)} provided")
        super().append(value)

''',
    },
    'openpyxl.descriptors.excel': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Excel specific descriptors
"""

from openpyxl.xml.constants import REL_NS
from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element

from . import (
    MatchPattern,
    MinMax,
    Integer,
    String,
    Sequence,
)
from .serialisable import Serialisable


class HexBinary(MatchPattern):

    pattern = "[0-9a-fA-F]+$"


class UniversalMeasure(MatchPattern):

    pattern = r"[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"


class TextPoint(MinMax):
    """
    Size in hundredths of points.
    In theory other units of measurement can be used but these are unbounded
    """
    expected_type = int

    min = -400000
    max = 400000


Coordinate = Integer


class Percentage(MinMax):

    pattern = r"((100)|([0-9][0-9]?))(\.[0-9][0-9]?)?%" # strict
    min = -1000000
    max = 1000000

    def __set__(self, instance, value):
        if isinstance(value, str) and "%" in value:
            value = value.replace("%", "")
            value = int(float(value) * 1000)
        super().__set__(instance, value)


class Extension(Serialisable):

    uri = String()

    def __init__(self,
                 uri=None,
                ):
        self.uri = uri


class ExtensionList(Serialisable):

    ext = Sequence(expected_type=Extension)

    def __init__(self,
                 ext=(),
                ):
        self.ext = ext


class Relation(String):

    namespace = REL_NS
    allow_none = True


class Base64Binary(MatchPattern):
    # http://www.w3.org/TR/xmlschema11-2/#nt-Base64Binary
    pattern = "^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=|[A-Za-z0-9+/]{4})$"


class Guid(MatchPattern):
    # https://msdn.microsoft.com/en-us/library/dd946381(v=office.12).aspx
    pattern = r"{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}"


class CellRange(MatchPattern):

    pattern = r"^[$]?([A-Za-z]{1,3})[$]?(\d+)(:[$]?([A-Za-z]{1,3})[$]?(\d+)?)?$|^[A-Za-z]{1,3}:[A-Za-z]{1,3}$"
    allow_none = True

    def __set__(self, instance, value):

        if value is not None:
            value = value.upper()
        super().__set__(instance, value)


def _explicit_none(tagname, value, namespace=None):
    """
    Override serialisation because explicit none required
    """
    if namespace is not None:
        tagname = "{%s}%s" % (namespace, tagname)
    return Element(tagname, val=safe_string(value))

''',
    },
    'openpyxl.descriptors.namespace': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


def namespaced(obj, tagname, namespace=None):
    """
    Utility to create a namespaced tag for an object
    """

    namespace = getattr(obj, "namespace", None) or namespace
    if namespace is not None:
        tagname = "{%s}%s" % (namespace, tagname)
    return tagname

''',
    },
    'openpyxl.descriptors.nested': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Generic serialisable classes
"""
from .base import (
    Convertible,
    Bool,
    Descriptor,
    NoneSet,
    MinMax,
    Set,
    Float,
    Integer,
    String,
    )
from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, localname, whitespace


class Nested(Descriptor):

    nested = True
    attribute = "val"

    def __set__(self, instance, value):
        if hasattr(value, "tag"):
            tag = localname(value)
            if tag != self.name:
                raise ValueError("Tag does not match attribute")

            value = self.from_tree(value)
        super().__set__(instance, value)


    def from_tree(self, node):
        return node.get(self.attribute)


    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if value is not None:
            if namespace is not None:
                tagname = "{%s}%s" % (namespace, tagname)
            value = safe_string(value)
            return Element(tagname, {self.attribute:value})


class NestedValue(Nested, Convertible):
    """
    Nested tag storing the value on the 'val' attribute
    """
    pass


class NestedText(NestedValue):
    """
    Represents any nested tag with the value as the contents of the tag
    """


    def from_tree(self, node):
        return node.text


    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if value is not None:
            if namespace is not None:
                tagname = "{%s}%s" % (namespace, tagname)
            el = Element(tagname)
            el.text = safe_string(value)
            whitespace(el)
            return el


class NestedFloat(NestedValue, Float):

    pass


class NestedInteger(NestedValue, Integer):

    pass


class NestedString(NestedValue, String):

    pass


class NestedBool(NestedValue, Bool):


    def from_tree(self, node):
        return node.get("val", True)


class NestedNoneSet(Nested, NoneSet):

    pass


class NestedSet(Nested, Set):

    pass


class NestedMinMax(Nested, MinMax):

    pass


class EmptyTag(Nested, Bool):

    """
    Boolean if a tag exists or not.
    """

    def from_tree(self, node):
        return True


    def to_tree(self, tagname=None, value=None, namespace=None):
        if value:
            namespace = getattr(self, "namespace", namespace)
            if namespace is not None:
                tagname = "{%s}%s" % (namespace, tagname)
            return Element(tagname)

''',
    },
    'openpyxl.descriptors.sequence': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element
from openpyxl.utils.indexed_list import IndexedList

from .base import Descriptor, Alias, _convert
from .namespace import namespaced


class Sequence(Descriptor):
    """
    A sequence (list or tuple) that may only contain objects of the declared
    type
    """

    expected_type = type(None)
    seq_types = (list, tuple)
    idx_base = 0
    unique = False
    container = list


    def __set__(self, instance, seq):
        if not isinstance(seq, self.seq_types):
            raise TypeError("Value must be a sequence")
        seq = self.container(_convert(self.expected_type, value) for value in seq)
        if self.unique:
            seq = IndexedList(seq)

        super().__set__(instance, seq)


    def to_tree(self, tagname, obj, namespace=None):
        """
        Convert the sequence represented by the descriptor to an XML element
        """
        for idx, v in enumerate(obj, self.idx_base):
            if hasattr(v, "to_tree"):
                el = v.to_tree(tagname, idx)
            else:
                tagname = namespaced(obj, tagname, namespace)
                el = Element(tagname)
                el.text = safe_string(v)
            yield el


class UniqueSequence(Sequence):
    """
    Use a set to keep values unique
    """
    seq_types = (list, tuple, set)
    container = set


class ValueSequence(Sequence):
    """
    A sequence of primitive types that are stored as a single attribute.
    "val" is the default attribute
    """

    attribute = "val"


    def to_tree(self, tagname, obj, namespace=None):
        tagname = namespaced(self, tagname, namespace)
        for v in obj:
            yield Element(tagname, {self.attribute:safe_string(v)})


    def from_tree(self, node):

        return node.get(self.attribute)


class NestedSequence(Sequence):
    """
    Wrap a sequence in an containing object
    """

    count = False

    def to_tree(self, tagname, obj, namespace=None):
        tagname = namespaced(self, tagname, namespace)
        container = Element(tagname)
        if self.count:
            container.set('count', str(len(obj)))
        for v in obj:
            container.append(v.to_tree())
        return container


    def from_tree(self, node):
        return [self.expected_type.from_tree(el) for el in node]


class MultiSequence(Sequence):
    """
    Sequences can contain objects with different tags
    """

    def __set__(self, instance, seq):
        if not isinstance(seq, (tuple, list)):
            raise ValueError("Value must be a sequence")
        seq = list(seq)
        Descriptor.__set__(self, instance, seq)


    def to_tree(self, tagname, obj, namespace=None):
        """
        Convert the sequence represented by the descriptor to an XML element
        """
        for v in obj:
            el = v.to_tree(namespace=namespace)
            yield el


class MultiSequencePart(Alias):
    """
    Allow a multisequence to be built up from parts

    Excluded from the instance __elements__ or __attrs__ as is effectively an Alias
    """

    def __init__(self, expected_type, store):
        self.expected_type = expected_type
        self.store = store


    def __set__(self, instance, value):
        value = _convert(self.expected_type, value)
        instance.__dict__[self.store].append(value)


    def __get__(self, instance, cls):
        return self

''',
    },
    'openpyxl.descriptors.serialisable': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from copy import copy
from keyword import kwlist
KEYWORDS = frozenset(kwlist)

from . import Descriptor
from . import MetaSerialisable
from .sequence import (
    Sequence,
    NestedSequence,
    MultiSequencePart,
)
from .namespace import namespaced

from openpyxl.compat import safe_string
from openpyxl.xml.functions import (
    Element,
    localname,
)

seq_types = (list, tuple)

class Serialisable(metaclass=MetaSerialisable):
    """
    Objects can serialise to XML their attributes and child objects.
    The following class attributes are created by the metaclass at runtime:
    __attrs__ = attributes
    __nested__ = single-valued child treated as an attribute
    __elements__ = child elements
    """

    __attrs__ = None
    __nested__ = None
    __elements__ = None
    __namespaced__ = None

    idx_base = 0

    @property
    def tagname(self):
        raise(NotImplementedError)

    namespace = None

    @classmethod
    def from_tree(cls, node):
        """
        Create object from XML
        """
        # strip known namespaces from attributes
        attrib = dict(node.attrib)
        for key, ns in cls.__namespaced__:
            if ns in attrib:
                attrib[key] = attrib[ns]
                del attrib[ns]

        # strip attributes with unknown namespaces
        for key in list(attrib):
            if key.startswith('{'):
                del attrib[key]
            elif key in KEYWORDS:
                attrib["_" + key] = attrib[key]
                del attrib[key]
            elif "-" in key:
                n = key.replace("-", "_")
                attrib[n] = attrib[key]
                del attrib[key]

        if node.text and "attr_text" in cls.__attrs__:
            attrib["attr_text"] = node.text

        for el in node:
            tag = localname(el)
            if tag in KEYWORDS:
                tag = "_" + tag
            desc = getattr(cls, tag, None)
            if desc is None or isinstance(desc, property):
                continue

            if hasattr(desc, 'from_tree'):
                #descriptor manages conversion
                obj = desc.from_tree(el)
            else:
                if hasattr(desc.expected_type, "from_tree"):
                    #complex type
                    obj = desc.expected_type.from_tree(el)
                else:
                    #primitive
                    obj = el.text

            if isinstance(desc, NestedSequence):
                attrib[tag] = obj
            elif isinstance(desc, Sequence):
                attrib.setdefault(tag, [])
                attrib[tag].append(obj)
            elif isinstance(desc, MultiSequencePart):
                attrib.setdefault(desc.store, [])
                attrib[desc.store].append(obj)
            else:
                attrib[tag] = obj

        return cls(**attrib)


    def to_tree(self, tagname=None, idx=None, namespace=None):

        if tagname is None:
            tagname = self.tagname

        # keywords have to be masked
        if tagname.startswith("_"):
            tagname = tagname[1:]

        tagname = namespaced(self, tagname, namespace)
        namespace = getattr(self, "namespace", namespace)

        attrs = dict(self)
        for key, ns in self.__namespaced__:
            if key in attrs:
                attrs[ns] = attrs[key]
                del attrs[key]

        el = Element(tagname, attrs)
        if "attr_text" in self.__attrs__:
            el.text = safe_string(getattr(self, "attr_text"))

        for child_tag in self.__elements__:
            desc = getattr(self.__class__, child_tag, None)
            obj = getattr(self, child_tag)
            if hasattr(desc, "namespace") and hasattr(obj, 'namespace'):
                obj.namespace = desc.namespace

            if isinstance(obj, seq_types):
                if isinstance(desc, NestedSequence):
                    # wrap sequence in container
                    if not obj:
                        continue
                    nodes = [desc.to_tree(child_tag, obj, namespace)]
                elif isinstance(desc, Sequence):
                    # sequence
                    desc.idx_base = self.idx_base
                    nodes = (desc.to_tree(child_tag, obj, namespace))
                else: # property
                    nodes = (v.to_tree(child_tag, namespace) for v in obj)
                for node in nodes:
                    el.append(node)
            else:
                if child_tag in self.__nested__:
                    node = desc.to_tree(child_tag, obj, namespace)
                elif obj is None:
                    continue
                else:
                    node = obj.to_tree(child_tag)
                if node is not None:
                    el.append(node)
        return el


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if attr.startswith("_"):
                attr = attr[1:]
            elif attr != "attr_text" and "_" in attr:
                desc = getattr(self.__class__, attr)
                if getattr(desc, "hyphenated", False):
                    attr = attr.replace("_", "-")
            if attr != "attr_text" and value is not None:
                yield attr, safe_string(value)


    def __eq__(self, other):
        if not self.__class__ == other.__class__:
            return False
        elif not dict(self) == dict(other):
            return False
        for el in self.__elements__:
            if getattr(self, el) != getattr(other, el):
                return False
        return True


    def __ne__(self, other):
        return not self == other


    def __repr__(self):
        s = u"<{0}.{1} object>\nParameters:".format(
            self.__module__,
            self.__class__.__name__
        )
        args = []
        for k in self.__attrs__ + self.__elements__:
            v = getattr(self, k)
            if isinstance(v, Descriptor):
                v = None
            args.append(u"{0}={1}".format(k, repr(v)))
        args = u", ".join(args)

        return u"\n".join([s, args])


    def __hash__(self):
        fields = []
        for attr in self.__attrs__ + self.__elements__:
            val = getattr(self, attr)
            if isinstance(val, list):
                val = tuple(val)
            fields.append(val)

        return hash(tuple(fields))


    def __add__(self, other):
        if type(self) != type(other):
            raise TypeError("Cannot combine instances of different types")
        vals = {}
        for attr in self.__attrs__:
            vals[attr] = getattr(self, attr) or getattr(other, attr)
        for el in self.__elements__:
            a = getattr(self, el)
            b = getattr(other, el)
            if a and b:
                vals[el] = a + b
            else:
                vals[el] = a or b
        return self.__class__(**vals)


    def __copy__(self):
        # serialise to xml and back to avoid shallow copies
        xml = self.to_tree(tagname="dummy")
        cp = self.__class__.from_tree(xml)
        # copy any non-persisted attributed
        for k in self.__dict__:
            if k not in self.__attrs__ + self.__elements__:
                v = copy(getattr(self, k))
                setattr(cp, k, v)
        return cp

''',
    },
    'openpyxl.descriptors.slots': {
        'is_package': False,
        'source': r"""
# Metaclass for mixing slots and descriptors
# From "Programming in Python 3" by Mark Summerfield Ch.8 p. 383

class AutoSlotProperties(type):

    def __new__(mcl, classname, bases, dictionary):
        slots = list(dictionary.get("__slots__", []))
        for getter_name in [key for key in dictionary if key.startswith("get_")]:
            name = getter_name
            slots.append("__" + name)
            getter = dictionary.pop(getter_name)
            setter = dictionary.get(setter_name, None)
            if (setter is not None
                and isinstance(setter, collections.Callable)):
                del dictionary[setter_name]
            dictionary[name] = property(getter. setter)
            dictionary["__slots__"] = tuple(slots)
            return super().__new__(mcl, classname, bases, dictionary)

""",
    },
    'openpyxl.drawing': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from .drawing import Drawing

""",
    },
    'openpyxl.drawing.colors': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    Integer,
    Set,
    MinMax,
)
from openpyxl.descriptors.excel import Percentage
from openpyxl.descriptors.nested import (
    NestedNoneSet,
    NestedValue,
    NestedInteger,
    EmptyTag,
)

from openpyxl.styles.colors import RGB
from openpyxl.xml.constants import DRAWING_NS

from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

PRESET_COLORS = [
        'aliceBlue', 'antiqueWhite', 'aqua', 'aquamarine',
        'azure', 'beige', 'bisque', 'black', 'blanchedAlmond', 'blue',
        'blueViolet', 'brown', 'burlyWood', 'cadetBlue', 'chartreuse',
        'chocolate', 'coral', 'cornflowerBlue', 'cornsilk', 'crimson', 'cyan',
        'darkBlue', 'darkCyan', 'darkGoldenrod', 'darkGray', 'darkGrey',
        'darkGreen', 'darkKhaki', 'darkMagenta', 'darkOliveGreen', 'darkOrange',
        'darkOrchid', 'darkRed', 'darkSalmon', 'darkSeaGreen', 'darkSlateBlue',
        'darkSlateGray', 'darkSlateGrey', 'darkTurquoise', 'darkViolet',
        'dkBlue', 'dkCyan', 'dkGoldenrod', 'dkGray', 'dkGrey', 'dkGreen',
        'dkKhaki', 'dkMagenta', 'dkOliveGreen', 'dkOrange', 'dkOrchid', 'dkRed',
        'dkSalmon', 'dkSeaGreen', 'dkSlateBlue', 'dkSlateGray', 'dkSlateGrey',
        'dkTurquoise', 'dkViolet', 'deepPink', 'deepSkyBlue', 'dimGray',
        'dimGrey', 'dodgerBlue', 'firebrick', 'floralWhite', 'forestGreen',
        'fuchsia', 'gainsboro', 'ghostWhite', 'gold', 'goldenrod', 'gray',
        'grey', 'green', 'greenYellow', 'honeydew', 'hotPink', 'indianRed',
        'indigo', 'ivory', 'khaki', 'lavender', 'lavenderBlush', 'lawnGreen',
        'lemonChiffon', 'lightBlue', 'lightCoral', 'lightCyan',
        'lightGoldenrodYellow', 'lightGray', 'lightGrey', 'lightGreen',
        'lightPink', 'lightSalmon', 'lightSeaGreen', 'lightSkyBlue',
        'lightSlateGray', 'lightSlateGrey', 'lightSteelBlue', 'lightYellow',
        'ltBlue', 'ltCoral', 'ltCyan', 'ltGoldenrodYellow', 'ltGray', 'ltGrey',
        'ltGreen', 'ltPink', 'ltSalmon', 'ltSeaGreen', 'ltSkyBlue',
        'ltSlateGray', 'ltSlateGrey', 'ltSteelBlue', 'ltYellow', 'lime',
        'limeGreen', 'linen', 'magenta', 'maroon', 'medAquamarine', 'medBlue',
        'medOrchid', 'medPurple', 'medSeaGreen', 'medSlateBlue',
        'medSpringGreen', 'medTurquoise', 'medVioletRed', 'mediumAquamarine',
        'mediumBlue', 'mediumOrchid', 'mediumPurple', 'mediumSeaGreen',
        'mediumSlateBlue', 'mediumSpringGreen', 'mediumTurquoise',
        'mediumVioletRed', 'midnightBlue', 'mintCream', 'mistyRose', 'moccasin',
        'navajoWhite', 'navy', 'oldLace', 'olive', 'oliveDrab', 'orange',
        'orangeRed', 'orchid', 'paleGoldenrod', 'paleGreen', 'paleTurquoise',
        'paleVioletRed', 'papayaWhip', 'peachPuff', 'peru', 'pink', 'plum',
        'powderBlue', 'purple', 'red', 'rosyBrown', 'royalBlue', 'saddleBrown',
        'salmon', 'sandyBrown', 'seaGreen', 'seaShell', 'sienna', 'silver',
        'skyBlue', 'slateBlue', 'slateGray', 'slateGrey', 'snow', 'springGreen',
        'steelBlue', 'tan', 'teal', 'thistle', 'tomato', 'turquoise', 'violet',
        'wheat', 'white', 'whiteSmoke', 'yellow', 'yellowGreen'
    ]


SCHEME_COLORS= ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3',
                'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'phClr', 'dk1', 'lt1',
                'dk2', 'lt2'
                ]


class Transform(Serialisable):

    pass


class SystemColor(Serialisable):

    tagname = "sysClr"
    namespace = DRAWING_NS

    # color transform options
    tint = NestedInteger(allow_none=True)
    shade = NestedInteger(allow_none=True)
    comp = Typed(expected_type=Transform, allow_none=True)
    inv = Typed(expected_type=Transform, allow_none=True)
    gray = Typed(expected_type=Transform, allow_none=True)
    alpha = NestedInteger(allow_none=True)
    alphaOff = NestedInteger(allow_none=True)
    alphaMod = NestedInteger(allow_none=True)
    hue = NestedInteger(allow_none=True)
    hueOff = NestedInteger(allow_none=True)
    hueMod = NestedInteger(allow_none=True)
    sat = NestedInteger(allow_none=True)
    satOff = NestedInteger(allow_none=True)
    satMod = NestedInteger(allow_none=True)
    lum = NestedInteger(allow_none=True)
    lumOff = NestedInteger(allow_none=True)
    lumMod = NestedInteger(allow_none=True)
    red = NestedInteger(allow_none=True)
    redOff = NestedInteger(allow_none=True)
    redMod = NestedInteger(allow_none=True)
    green = NestedInteger(allow_none=True)
    greenOff = NestedInteger(allow_none=True)
    greenMod = NestedInteger(allow_none=True)
    blue = NestedInteger(allow_none=True)
    blueOff = NestedInteger(allow_none=True)
    blueMod = NestedInteger(allow_none=True)
    gamma = Typed(expected_type=Transform, allow_none=True)
    invGamma = Typed(expected_type=Transform, allow_none=True)

    val = Set(values=( ['scrollBar', 'background', 'activeCaption',
                        'inactiveCaption', 'menu', 'window', 'windowFrame', 'menuText',
                        'windowText', 'captionText', 'activeBorder', 'inactiveBorder',
                        'appWorkspace', 'highlight', 'highlightText', 'btnFace', 'btnShadow',
                        'grayText', 'btnText', 'inactiveCaptionText', 'btnHighlight',
                        '3dDkShadow', '3dLight', 'infoText', 'infoBk', 'hotLight',
                        'gradientActiveCaption', 'gradientInactiveCaption', 'menuHighlight',
                        'menuBar'] )
              )
    lastClr = RGB(allow_none=True)

    __elements__ = ('tint', 'shade', 'comp', 'inv', 'gray', "alpha",
                    "alphaOff", "alphaMod", "hue", "hueOff", "hueMod", "hueOff", "sat",
                    "satOff", "satMod", "lum", "lumOff", "lumMod", "red", "redOff", "redMod",
                    "green", "greenOff", "greenMod", "blue", "blueOff", "blueMod", "gamma",
                    "invGamma")

    def __init__(self,
                 val="windowText",
                 lastClr=None,
                 tint=None,
                 shade=None,
                 comp=None,
                 inv=None,
                 gray=None,
                 alpha=None,
                 alphaOff=None,
                 alphaMod=None,
                 hue=None,
                 hueOff=None,
                 hueMod=None,
                 sat=None,
                 satOff=None,
                 satMod=None,
                 lum=None,
                 lumOff=None,
                 lumMod=None,
                 red=None,
                 redOff=None,
                 redMod=None,
                 green=None,
                 greenOff=None,
                 greenMod=None,
                 blue=None,
                 blueOff=None,
                 blueMod=None,
                 gamma=None,
                 invGamma=None
                ):
        self.val = val
        self.lastClr = lastClr
        self.tint = tint
        self.shade = shade
        self.comp = comp
        self.inv = inv
        self.gray = gray
        self.alpha = alpha
        self.alphaOff = alphaOff
        self.alphaMod = alphaMod
        self.hue = hue
        self.hueOff = hueOff
        self.hueMod = hueMod
        self.sat = sat
        self.satOff = satOff
        self.satMod = satMod
        self.lum = lum
        self.lumOff = lumOff
        self.lumMod = lumMod
        self.red = red
        self.redOff = redOff
        self.redMod = redMod
        self.green = green
        self.greenOff = greenOff
        self.greenMod = greenMod
        self.blue = blue
        self.blueOff = blueOff
        self.blueMod = blueMod
        self.gamma = gamma
        self.invGamma = invGamma


class HSLColor(Serialisable):

    tagname = "hslClr"

    hue = Integer()
    sat = MinMax(min=0, max=100)
    lum = MinMax(min=0, max=100)

    #TODO add color transform options

    def __init__(self,
                 hue=None,
                 sat=None,
                 lum=None,
                ):
        self.hue = hue
        self.sat = sat
        self.lum = lum



class RGBPercent(Serialisable):

    tagname = "rgbClr"

    r = MinMax(min=0, max=100)
    g = MinMax(min=0, max=100)
    b = MinMax(min=0, max=100)

    #TODO add color transform options

    def __init__(self,
                 r=None,
                 g=None,
                 b=None,
                ):
        self.r = r
        self.g = g
        self.b = b


class SchemeColor(Serialisable):

    tagname = "schemeClr"
    namespace = DRAWING_NS

    tint = NestedInteger(allow_none=True)
    shade = NestedInteger(allow_none=True)
    comp = EmptyTag(allow_none=True)
    inv = NestedInteger(allow_none=True)
    gray = NestedInteger(allow_none=True)
    alpha = NestedInteger(allow_none=True)
    alphaOff = NestedInteger(allow_none=True)
    alphaMod = NestedInteger(allow_none=True)
    hue = NestedInteger(allow_none=True)
    hueOff = NestedInteger(allow_none=True)
    hueMod = NestedInteger(allow_none=True)
    sat = NestedInteger(allow_none=True)
    satOff = NestedInteger(allow_none=True)
    satMod = NestedInteger(allow_none=True)
    lum = NestedInteger(allow_none=True)
    lumOff = NestedInteger(allow_none=True)
    lumMod = NestedInteger(allow_none=True)
    red = NestedInteger(allow_none=True)
    redOff = NestedInteger(allow_none=True)
    redMod = NestedInteger(allow_none=True)
    green = NestedInteger(allow_none=True)
    greenOff = NestedInteger(allow_none=True)
    greenMod = NestedInteger(allow_none=True)
    blue = NestedInteger(allow_none=True)
    blueOff = NestedInteger(allow_none=True)
    blueMod = NestedInteger(allow_none=True)
    gamma = EmptyTag(allow_none=True)
    invGamma = EmptyTag(allow_none=True)
    val = Set(values=(['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2',
                       'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'phClr',
                       'dk1', 'lt1', 'dk2', 'lt2']))

    __elements__ = ('tint', 'shade', 'comp', 'inv', 'gray', 'alpha',
                    'alphaOff', 'alphaMod', 'hue', 'hueOff', 'hueMod', 'sat', 'satOff',
                    'satMod', 'lum', 'lumMod', 'lumOff', 'red', 'redOff', 'redMod', 'green',
                    'greenOff', 'greenMod', 'blue', 'blueOff', 'blueMod', 'gamma',
                    'invGamma')

    def __init__(self,
                 tint=None,
                 shade=None,
                 comp=None,
                 inv=None,
                 gray=None,
                 alpha=None,
                 alphaOff=None,
                 alphaMod=None,
                 hue=None,
                 hueOff=None,
                 hueMod=None,
                 sat=None,
                 satOff=None,
                 satMod=None,
                 lum=None,
                 lumOff=None,
                 lumMod=None,
                 red=None,
                 redOff=None,
                 redMod=None,
                 green=None,
                 greenOff=None,
                 greenMod=None,
                 blue=None,
                 blueOff=None,
                 blueMod=None,
                 gamma=None,
                 invGamma=None,
                 val=None,
                ):
        self.tint = tint
        self.shade = shade
        self.comp = comp
        self.inv = inv
        self.gray = gray
        self.alpha = alpha
        self.alphaOff = alphaOff
        self.alphaMod = alphaMod
        self.hue = hue
        self.hueOff = hueOff
        self.hueMod = hueMod
        self.sat = sat
        self.satOff = satOff
        self.satMod = satMod
        self.lum = lum
        self.lumOff = lumOff
        self.lumMod = lumMod
        self.red = red
        self.redOff = redOff
        self.redMod = redMod
        self.green = green
        self.greenOff = greenOff
        self.greenMod = greenMod
        self.blue = blue
        self.blueOff = blueOff
        self.blueMod = blueMod
        self.gamma = gamma
        self.invGamma = invGamma
        self.val = val

class ColorChoice(Serialisable):

    tagname = "colorChoice"
    namespace = DRAWING_NS

    scrgbClr = Typed(expected_type=RGBPercent, allow_none=True)
    RGBPercent = Alias('scrgbClr')
    srgbClr = NestedValue(expected_type=str, allow_none=True) # needs pattern and can have transform
    RGB = Alias('srgbClr')
    hslClr = Typed(expected_type=HSLColor, allow_none=True)
    sysClr = Typed(expected_type=SystemColor, allow_none=True)
    schemeClr = Typed(expected_type=SchemeColor, allow_none=True)
    prstClr = NestedNoneSet(values=PRESET_COLORS)

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 scrgbClr=None,
                 srgbClr=None,
                 hslClr=None,
                 sysClr=None,
                 schemeClr=None,
                 prstClr=None,
                ):
        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr

_COLOR_SET = ('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
               'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')


class ColorMapping(Serialisable):

    tagname = "clrMapOvr"

    bg1 = Set(values=_COLOR_SET)
    tx1 = Set(values=_COLOR_SET)
    bg2 = Set(values=_COLOR_SET)
    tx2 = Set(values=_COLOR_SET)
    accent1 = Set(values=_COLOR_SET)
    accent2 = Set(values=_COLOR_SET)
    accent3 = Set(values=_COLOR_SET)
    accent4 = Set(values=_COLOR_SET)
    accent5 = Set(values=_COLOR_SET)
    accent6 = Set(values=_COLOR_SET)
    hlink = Set(values=_COLOR_SET)
    folHlink = Set(values=_COLOR_SET)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bg1="lt1",
                 tx1="dk1",
                 bg2="lt2",
                 tx2="dk2",
                 accent1="accent1",
                 accent2="accent2",
                 accent3="accent3",
                 accent4="accent4",
                 accent5="accent5",
                 accent6="accent6",
                 hlink="hlink",
                 folHlink="folHlink",
                 extLst=None,
                ):
        self.bg1 = bg1
        self.tx1 = tx1
        self.bg2 = bg2
        self.tx2 = tx2
        self.accent1 = accent1
        self.accent2 = accent2
        self.accent3 = accent3
        self.accent4 = accent4
        self.accent5 = accent5
        self.accent6 = accent6
        self.hlink = hlink
        self.folHlink = folHlink
        self.extLst = extLst


class ColorChoiceDescriptor(Typed):
    """
    Objects can choose from 7 different kinds of color system.
    Assume RGBHex if a string is passed in.
    """

    expected_type = ColorChoice
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, str):
            value = ColorChoice(srgbClr=value)
        else:
            if hasattr(self, "namespace") and value is not None:
                value.namespace = self.namespace
        super().__set__(instance, value)

''',
    },
    'openpyxl.drawing.connector': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    String,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText

from .properties import (
    NonVisualDrawingProps,
    NonVisualDrawingShapeProps,
)
from .geometry import ShapeStyle

class Connection(Serialisable):

    id = Integer()
    idx = Integer()

    def __init__(self,
                 id=None,
                 idx=None,
                ):
        self.id = id
        self.idx = idx


class ConnectorLocking(Serialisable):

    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 extLst=None,
                ):
        self.extLst = extLst


class NonVisualConnectorProperties(Serialisable):

    cxnSpLocks = Typed(expected_type=ConnectorLocking, allow_none=True)
    stCxn = Typed(expected_type=Connection, allow_none=True)
    endCxn = Typed(expected_type=Connection, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 cxnSpLocks=None,
                 stCxn=None,
                 endCxn=None,
                 extLst=None,
                ):
        self.cxnSpLocks = cxnSpLocks
        self.stCxn = stCxn
        self.endCxn = endCxn
        self.extLst = extLst


class ConnectorNonVisual(Serialisable):

    cNvPr = Typed(expected_type=NonVisualDrawingProps, )
    cNvCxnSpPr = Typed(expected_type=NonVisualConnectorProperties, )

    __elements__ = ("cNvPr", "cNvCxnSpPr",)

    def __init__(self,
                 cNvPr=None,
                 cNvCxnSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvCxnSpPr = cNvCxnSpPr


class ConnectorShape(Serialisable):

    tagname = "cxnSp"

    nvCxnSpPr = Typed(expected_type=ConnectorNonVisual)
    spPr = Typed(expected_type=GraphicalProperties)
    style = Typed(expected_type=ShapeStyle, allow_none=True)
    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)

    def __init__(self,
                 nvCxnSpPr=None,
                 spPr=None,
                 style=None,
                 macro=None,
                 fPublished=None,
                 ):
        self.nvCxnSpPr = nvCxnSpPr
        self.spPr = spPr
        self.style = style
        self.macro = macro
        self.fPublished = fPublished


class ShapeMeta(Serialisable):

    tagname = "nvSpPr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps)
    cNvSpPr = Typed(expected_type=NonVisualDrawingShapeProps)

    def __init__(self, cNvPr=None, cNvSpPr=None):
        self.cNvPr = cNvPr
        self.cNvSpPr = cNvSpPr


class Shape(Serialisable):

    macro = String(allow_none=True)
    textlink = String(allow_none=True)
    fPublished = Bool(allow_none=True)
    fLocksText = Bool(allow_none=True)
    nvSpPr = Typed(expected_type=ShapeMeta, allow_none=True)
    meta = Alias("nvSpPr")
    spPr = Typed(expected_type=GraphicalProperties)
    graphicalProperties = Alias("spPr")
    style = Typed(expected_type=ShapeStyle, allow_none=True)
    txBody = Typed(expected_type=RichText, allow_none=True)

    def __init__(self,
                 macro=None,
                 textlink=None,
                 fPublished=None,
                 fLocksText=None,
                 nvSpPr=None,
                 spPr=None,
                 style=None,
                 txBody=None,
                ):
        self.macro = macro
        self.textlink = textlink
        self.fPublished = fPublished
        self.fLocksText = fLocksText
        self.nvSpPr = nvSpPr
        self.spPr = spPr
        self.style = style
        self.txBody = txBody

""",
    },
    'openpyxl.drawing.drawing': {
        'is_package': False,
        'source': r'''

# Copyright (c) 2010-2024 openpyxl

import math

from openpyxl.utils.units import pixels_to_EMU


class Drawing:
    """ a drawing object - eg container for shapes or charts
        we assume user specifies dimensions in pixels; units are
        converted to EMU in the drawing part
    """

    count = 0

    def __init__(self):

        self.name = ''
        self.description = ''
        self.coordinates = ((1, 2), (16, 8))
        self.left = 0
        self.top = 0
        self._width = 21 # default in px
        self._height = 192 #default in px
        self.resize_proportional = False
        self.rotation = 0
        self.anchortype = "absolute"
        self.anchorcol = 0 # left cell
        self.anchorrow = 0 # top row


    @property
    def width(self):
        return self._width


    @width.setter
    def width(self, w):
        if self.resize_proportional and w:
            ratio = self._height / self._width
            self._height = round(ratio * w)
        self._width = w


    @property
    def height(self):
        return self._height


    @height.setter
    def height(self, h):
        if self.resize_proportional and h:
            ratio = self._width / self._height
            self._width = round(ratio * h)
        self._height = h


    def set_dimension(self, w=0, h=0):

        xratio = w / self._width
        yratio = h / self._height

        if self.resize_proportional and w and h:
            if (xratio * self._height) < h:
                self._height = math.ceil(xratio * self._height)
                self._width = w
            else:
                self._width = math.ceil(yratio * self._width)
                self._height = h


    @property
    def anchor(self):
        from .spreadsheet_drawing import (
            OneCellAnchor,
            TwoCellAnchor,
            AbsoluteAnchor)
        if self.anchortype == "absolute":
            anchor = AbsoluteAnchor()
            anchor.pos.x = pixels_to_EMU(self.left)
            anchor.pos.y = pixels_to_EMU(self.top)

        elif self.anchortype == "oneCell":
            anchor = OneCellAnchor()
            anchor._from.col = self.anchorcol
            anchor._from.row = self.anchorrow

        anchor.ext.width = pixels_to_EMU(self._width)
        anchor.ext.height = pixels_to_EMU(self._height)

        return anchor

''',
    },
    'openpyxl.drawing.effect': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Set,
    Bool,
    Integer,
    Float,
)

from .colors import ColorChoice


class TintEffect(Serialisable):

    tagname = "tint"

    hue = Integer()
    amt = Integer()

    def __init__(self,
                 hue=0,
                 amt=0,
                ):
        self.hue = hue
        self.amt = amt


class LuminanceEffect(Serialisable):

    tagname = "lum"

    bright = Integer() #Pct ?
    contrast = Integer() #Pct#

    def __init__(self,
                 bright=0,
                 contrast=0,
                ):
        self.bright = bright
        self.contrast = contrast


class HSLEffect(Serialisable):

    hue = Integer()
    sat = Integer()
    lum = Integer()

    def __init__(self,
                 hue=None,
                 sat=None,
                 lum=None,
                ):
        self.hue = hue
        self.sat = sat
        self.lum = lum


class GrayscaleEffect(Serialisable):

    tagname = "grayscl"


class FillOverlayEffect(Serialisable):

    blend = Set(values=(['over', 'mult', 'screen', 'darken', 'lighten']))

    def __init__(self,
                 blend=None,
                ):
        self.blend = blend


class DuotoneEffect(Serialisable):

    pass

class ColorReplaceEffect(Serialisable):

    pass

class Color(Serialisable):

    pass

class ColorChangeEffect(Serialisable):

    useA = Bool(allow_none=True)
    clrFrom = Typed(expected_type=Color, )
    clrTo = Typed(expected_type=Color, )

    def __init__(self,
                 useA=None,
                 clrFrom=None,
                 clrTo=None,
                ):
        self.useA = useA
        self.clrFrom = clrFrom
        self.clrTo = clrTo


class BlurEffect(Serialisable):

    rad = Float()
    grow = Bool(allow_none=True)

    def __init__(self,
                 rad=None,
                 grow=None,
                ):
        self.rad = rad
        self.grow = grow


class BiLevelEffect(Serialisable):

    thresh = Integer()

    def __init__(self,
                 thresh=None,
                ):
        self.thresh = thresh


class AlphaReplaceEffect(Serialisable):

    a = Integer()

    def __init__(self,
                 a=None,
                ):
        self.a = a


class AlphaModulateFixedEffect(Serialisable):

    amt = Integer()

    def __init__(self,
                 amt=None,
                ):
        self.amt = amt


class EffectContainer(Serialisable):

    type = Set(values=(['sib', 'tree']))
    name = String(allow_none=True)

    def __init__(self,
                 type=None,
                 name=None,
                ):
        self.type = type
        self.name = name


class AlphaModulateEffect(Serialisable):

    cont = Typed(expected_type=EffectContainer, )

    def __init__(self,
                 cont=None,
                ):
        self.cont = cont


class AlphaInverseEffect(Serialisable):

    pass

class AlphaFloorEffect(Serialisable):

    pass

class AlphaCeilingEffect(Serialisable):

    pass

class AlphaBiLevelEffect(Serialisable):

    thresh = Integer()

    def __init__(self,
                 thresh=None,
                ):
        self.thresh = thresh


class GlowEffect(ColorChoice):

    rad = Float()
    # uses element group EG_ColorChoice
    scrgbClr = ColorChoice.scrgbClr
    srgbClr = ColorChoice.srgbClr
    hslClr = ColorChoice.hslClr
    sysClr = ColorChoice.sysClr
    schemeClr = ColorChoice.schemeClr
    prstClr = ColorChoice.prstClr

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 rad=None,
                 **kw
                ):
        self.rad = rad
        super().__init__(**kw)


class InnerShadowEffect(ColorChoice):

    blurRad = Float()
    dist = Float()
    dir = Integer()
    # uses element group EG_ColorChoice
    scrgbClr = ColorChoice.scrgbClr
    srgbClr = ColorChoice.srgbClr
    hslClr = ColorChoice.hslClr
    sysClr = ColorChoice.sysClr
    schemeClr = ColorChoice.schemeClr
    prstClr = ColorChoice.prstClr

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 blurRad=None,
                 dist=None,
                 dir=None,
                 **kw
                 ):
        self.blurRad = blurRad
        self.dist = dist
        self.dir = dir
        super().__init__(**kw)


class OuterShadow(ColorChoice):

    tagname = "outerShdw"

    blurRad = Float(allow_none=True)
    dist = Float(allow_none=True)
    dir = Integer(allow_none=True)
    sx = Integer(allow_none=True)
    sy = Integer(allow_none=True)
    kx = Integer(allow_none=True)
    ky = Integer(allow_none=True)
    algn = Set(values=['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'])
    rotWithShape = Bool(allow_none=True)
    # uses element group EG_ColorChoice
    scrgbClr = ColorChoice.scrgbClr
    srgbClr = ColorChoice.srgbClr
    hslClr = ColorChoice.hslClr
    sysClr = ColorChoice.sysClr
    schemeClr = ColorChoice.schemeClr
    prstClr = ColorChoice.prstClr

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 blurRad=None,
                 dist=None,
                 dir=None,
                 sx=None,
                 sy=None,
                 kx=None,
                 ky=None,
                 algn=None,
                 rotWithShape=None,
                 **kw
                ):
        self.blurRad = blurRad
        self.dist = dist
        self.dir = dir
        self.sx = sx
        self.sy = sy
        self.kx = kx
        self.ky = ky
        self.algn = algn
        self.rotWithShape = rotWithShape
        super().__init__(**kw)


class PresetShadowEffect(ColorChoice):

    prst = Set(values=(['shdw1', 'shdw2', 'shdw3', 'shdw4', 'shdw5', 'shdw6',
                        'shdw7', 'shdw8', 'shdw9', 'shdw10', 'shdw11', 'shdw12', 'shdw13',
                        'shdw14', 'shdw15', 'shdw16', 'shdw17', 'shdw18', 'shdw19', 'shdw20']))
    dist = Float()
    dir = Integer()
    # uses element group EG_ColorChoice
    scrgbClr = ColorChoice.scrgbClr
    srgbClr = ColorChoice.srgbClr
    hslClr = ColorChoice.hslClr
    sysClr = ColorChoice.sysClr
    schemeClr = ColorChoice.schemeClr
    prstClr = ColorChoice.prstClr

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 prst=None,
                 dist=None,
                 dir=None,
                 **kw
                ):
        self.prst = prst
        self.dist = dist
        self.dir = dir
        super().__init__(**kw)


class ReflectionEffect(Serialisable):

    blurRad = Float()
    stA = Integer()
    stPos = Integer()
    endA = Integer()
    endPos = Integer()
    dist = Float()
    dir = Integer()
    fadeDir = Integer()
    sx = Integer()
    sy = Integer()
    kx = Integer()
    ky = Integer()
    algn = Set(values=(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']))
    rotWithShape = Bool(allow_none=True)

    def __init__(self,
                 blurRad=None,
                 stA=None,
                 stPos=None,
                 endA=None,
                 endPos=None,
                 dist=None,
                 dir=None,
                 fadeDir=None,
                 sx=None,
                 sy=None,
                 kx=None,
                 ky=None,
                 algn=None,
                 rotWithShape=None,
                ):
        self.blurRad = blurRad
        self.stA = stA
        self.stPos = stPos
        self.endA = endA
        self.endPos = endPos
        self.dist = dist
        self.dir = dir
        self.fadeDir = fadeDir
        self.sx = sx
        self.sy = sy
        self.kx = kx
        self.ky = ky
        self.algn = algn
        self.rotWithShape = rotWithShape


class SoftEdgesEffect(Serialisable):

    rad = Float()

    def __init__(self,
                 rad=None,
                ):
        self.rad = rad


class EffectList(Serialisable):

    blur = Typed(expected_type=BlurEffect, allow_none=True)
    fillOverlay = Typed(expected_type=FillOverlayEffect, allow_none=True)
    glow = Typed(expected_type=GlowEffect, allow_none=True)
    innerShdw = Typed(expected_type=InnerShadowEffect, allow_none=True)
    outerShdw = Typed(expected_type=OuterShadow, allow_none=True)
    prstShdw = Typed(expected_type=PresetShadowEffect, allow_none=True)
    reflection = Typed(expected_type=ReflectionEffect, allow_none=True)
    softEdge = Typed(expected_type=SoftEdgesEffect, allow_none=True)

    __elements__ = ('blur', 'fillOverlay', 'glow', 'innerShdw', 'outerShdw',
                    'prstShdw', 'reflection', 'softEdge')

    def __init__(self,
                 blur=None,
                 fillOverlay=None,
                 glow=None,
                 innerShdw=None,
                 outerShdw=None,
                 prstShdw=None,
                 reflection=None,
                 softEdge=None,
                ):
        self.blur = blur
        self.fillOverlay = fillOverlay
        self.glow = glow
        self.innerShdw = innerShdw
        self.outerShdw = outerShdw
        self.prstShdw = prstShdw
        self.reflection = reflection
        self.softEdge = softEdge

""",
    },
    'openpyxl.drawing.fill': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Bool,
    Integer,
    Set,
    NoneSet,
    Typed,
    MinMax,
)
from openpyxl.descriptors.excel import (
    Relation,
    Percentage,
)
from openpyxl.descriptors.nested import NestedNoneSet, NestedValue
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from openpyxl.xml.constants import DRAWING_NS

from .colors import (
    ColorChoice,
    HSLColor,
    SystemColor,
    SchemeColor,
    PRESET_COLORS,
    RGBPercent,
)

from .effect import (
    AlphaBiLevelEffect,
    AlphaCeilingEffect,
    AlphaFloorEffect,
    AlphaInverseEffect,
    AlphaModulateEffect,
    AlphaModulateFixedEffect,
    AlphaReplaceEffect,
    BiLevelEffect,
    BlurEffect,
    ColorChangeEffect,
    ColorReplaceEffect,
    DuotoneEffect,
    FillOverlayEffect,
    GrayscaleEffect,
    HSLEffect,
    LuminanceEffect,
    TintEffect,
)

"""
Fill elements from drawing main schema
"""

class PatternFillProperties(Serialisable):

    tagname = "pattFill"
    namespace = DRAWING_NS

    prst = NoneSet(values=(['pct5', 'pct10', 'pct20', 'pct25', 'pct30',
                            'pct40', 'pct50', 'pct60', 'pct70', 'pct75', 'pct80', 'pct90', 'horz',
                            'vert', 'ltHorz', 'ltVert', 'dkHorz', 'dkVert', 'narHorz', 'narVert',
                            'dashHorz', 'dashVert', 'cross', 'dnDiag', 'upDiag', 'ltDnDiag',
                            'ltUpDiag', 'dkDnDiag', 'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'dashDnDiag',
                            'dashUpDiag', 'diagCross', 'smCheck', 'lgCheck', 'smGrid', 'lgGrid',
                            'dotGrid', 'smConfetti', 'lgConfetti', 'horzBrick', 'diagBrick',
                            'solidDmnd', 'openDmnd', 'dotDmnd', 'plaid', 'sphere', 'weave', 'divot',
                            'shingle', 'wave', 'trellis', 'zigZag']))
    preset = Alias("prst")
    fgClr = Typed(expected_type=ColorChoice, allow_none=True)
    foreground = Alias("fgClr")
    bgClr = Typed(expected_type=ColorChoice, allow_none=True)
    background = Alias("bgClr")

    __elements__ = ("fgClr", "bgClr")

    def __init__(self,
                 prst=None,
                 fgClr=None,
                 bgClr=None,
                ):
        self.prst = prst
        self.fgClr = fgClr
        self.bgClr = bgClr


class RelativeRect(Serialisable):

    tagname = "rect"
    namespace = DRAWING_NS

    l = Percentage(allow_none=True)
    left = Alias('l')
    t = Percentage(allow_none=True)
    top = Alias('t')
    r = Percentage(allow_none=True)
    right = Alias('r')
    b = Percentage(allow_none=True)
    bottom = Alias('b')

    def __init__(self,
                 l=None,
                 t=None,
                 r=None,
                 b=None,
                ):
        self.l = l
        self.t = t
        self.r = r
        self.b = b


class StretchInfoProperties(Serialisable):

    tagname = "stretch"
    namespace = DRAWING_NS

    fillRect = Typed(expected_type=RelativeRect, allow_none=True)

    def __init__(self,
                 fillRect=RelativeRect(),
                ):
        self.fillRect = fillRect


class GradientStop(Serialisable):

    tagname = "gs"
    namespace = DRAWING_NS

    pos = MinMax(min=0, max=100000, allow_none=True)
    # Color Choice Group
    scrgbClr = Typed(expected_type=RGBPercent, allow_none=True)
    RGBPercent = Alias('scrgbClr')
    srgbClr = NestedValue(expected_type=str, allow_none=True) # needs pattern and can have transform
    RGB = Alias('srgbClr')
    hslClr = Typed(expected_type=HSLColor, allow_none=True)
    sysClr = Typed(expected_type=SystemColor, allow_none=True)
    schemeClr = Typed(expected_type=SchemeColor, allow_none=True)
    prstClr = NestedNoneSet(values=PRESET_COLORS)

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 pos=None,
                 scrgbClr=None,
                 srgbClr=None,
                 hslClr=None,
                 sysClr=None,
                 schemeClr=None,
                 prstClr=None,
                ):
        if pos is None:
            pos = 0
        self.pos = pos

        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr


class LinearShadeProperties(Serialisable):

    tagname = "lin"
    namespace = DRAWING_NS

    ang = Integer()
    scaled = Bool(allow_none=True)

    def __init__(self,
                 ang=None,
                 scaled=None,
                ):
        self.ang = ang
        self.scaled = scaled


class PathShadeProperties(Serialisable):

    tagname = "path"
    namespace = DRAWING_NS

    path = Set(values=(['shape', 'circle', 'rect']))
    fillToRect = Typed(expected_type=RelativeRect, allow_none=True)

    def __init__(self,
                 path=None,
                 fillToRect=None,
                ):
        self.path = path
        self.fillToRect = fillToRect


class GradientFillProperties(Serialisable):

    tagname = "gradFill"
    namespace = DRAWING_NS

    flip = NoneSet(values=(['x', 'y', 'xy']))
    rotWithShape = Bool(allow_none=True)

    gsLst = NestedSequence(expected_type=GradientStop, count=False)
    stop_list = Alias("gsLst")

    lin = Typed(expected_type=LinearShadeProperties, allow_none=True)
    linear = Alias("lin")
    path = Typed(expected_type=PathShadeProperties, allow_none=True)

    tileRect = Typed(expected_type=RelativeRect, allow_none=True)

    __elements__ = ('gsLst', 'lin', 'path', 'tileRect')

    def __init__(self,
                 flip=None,
                 rotWithShape=None,
                 gsLst=(),
                 lin=None,
                 path=None,
                 tileRect=None,
                ):
        self.flip = flip
        self.rotWithShape = rotWithShape
        self.gsLst = gsLst
        self.lin = lin
        self.path = path
        self.tileRect = tileRect


class SolidColorFillProperties(Serialisable):

    tagname = "solidFill"

    # uses element group EG_ColorChoice
    scrgbClr = Typed(expected_type=RGBPercent, allow_none=True)
    RGBPercent = Alias('scrgbClr')
    srgbClr = NestedValue(expected_type=str, allow_none=True) # needs pattern and can have transform
    RGB = Alias('srgbClr')
    hslClr = Typed(expected_type=HSLColor, allow_none=True)
    sysClr = Typed(expected_type=SystemColor, allow_none=True)
    schemeClr = Typed(expected_type=SchemeColor, allow_none=True)
    prstClr = NestedNoneSet(values=PRESET_COLORS)

    __elements__ = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 scrgbClr=None,
                 srgbClr=None,
                 hslClr=None,
                 sysClr=None,
                 schemeClr=None,
                 prstClr=None,
                ):
        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr


class Blip(Serialisable):

    tagname = "blip"
    namespace = DRAWING_NS

    # Using attribute groupAG_Blob
    cstate = NoneSet(values=(['email', 'screen', 'print', 'hqprint']))
    embed = Relation() # rId
    link = Relation() # hyperlink
    noGrp = Bool(allow_none=True)
    noSelect = Bool(allow_none=True)
    noRot = Bool(allow_none=True)
    noChangeAspect = Bool(allow_none=True)
    noMove = Bool(allow_none=True)
    noResize = Bool(allow_none=True)
    noEditPoints = Bool(allow_none=True)
    noAdjustHandles = Bool(allow_none=True)
    noChangeArrowheads = Bool(allow_none=True)
    noChangeShapeType = Bool(allow_none=True)
    # some elements are choice
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)
    alphaBiLevel = Typed(expected_type=AlphaBiLevelEffect, allow_none=True)
    alphaCeiling = Typed(expected_type=AlphaCeilingEffect, allow_none=True)
    alphaFloor = Typed(expected_type=AlphaFloorEffect, allow_none=True)
    alphaInv = Typed(expected_type=AlphaInverseEffect, allow_none=True)
    alphaMod = Typed(expected_type=AlphaModulateEffect, allow_none=True)
    alphaModFix = Typed(expected_type=AlphaModulateFixedEffect, allow_none=True)
    alphaRepl = Typed(expected_type=AlphaReplaceEffect, allow_none=True)
    biLevel = Typed(expected_type=BiLevelEffect, allow_none=True)
    blur = Typed(expected_type=BlurEffect, allow_none=True)
    clrChange = Typed(expected_type=ColorChangeEffect, allow_none=True)
    clrRepl = Typed(expected_type=ColorReplaceEffect, allow_none=True)
    duotone = Typed(expected_type=DuotoneEffect, allow_none=True)
    fillOverlay = Typed(expected_type=FillOverlayEffect, allow_none=True)
    grayscl = Typed(expected_type=GrayscaleEffect, allow_none=True)
    hsl = Typed(expected_type=HSLEffect, allow_none=True)
    lum = Typed(expected_type=LuminanceEffect, allow_none=True)
    tint = Typed(expected_type=TintEffect, allow_none=True)

    __elements__ = ('alphaBiLevel', 'alphaCeiling', 'alphaFloor', 'alphaInv',
                    'alphaMod', 'alphaModFix', 'alphaRepl', 'biLevel', 'blur', 'clrChange',
                    'clrRepl', 'duotone', 'fillOverlay', 'grayscl', 'hsl', 'lum', 'tint')

    def __init__(self,
                 cstate=None,
                 embed=None,
                 link=None,
                 noGrp=None,
                 noSelect=None,
                 noRot=None,
                 noChangeAspect=None,
                 noMove=None,
                 noResize=None,
                 noEditPoints=None,
                 noAdjustHandles=None,
                 noChangeArrowheads=None,
                 noChangeShapeType=None,
                 extLst=None,
                 alphaBiLevel=None,
                 alphaCeiling=None,
                 alphaFloor=None,
                 alphaInv=None,
                 alphaMod=None,
                 alphaModFix=None,
                 alphaRepl=None,
                 biLevel=None,
                 blur=None,
                 clrChange=None,
                 clrRepl=None,
                 duotone=None,
                 fillOverlay=None,
                 grayscl=None,
                 hsl=None,
                 lum=None,
                 tint=None,
                ):
        self.cstate = cstate
        self.embed = embed
        self.link = link
        self.noGrp = noGrp
        self.noSelect = noSelect
        self.noRot = noRot
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.noEditPoints = noEditPoints
        self.noAdjustHandles = noAdjustHandles
        self.noChangeArrowheads = noChangeArrowheads
        self.noChangeShapeType = noChangeShapeType
        self.extLst = extLst
        self.alphaBiLevel = alphaBiLevel
        self.alphaCeiling = alphaCeiling
        self.alphaFloor = alphaFloor
        self.alphaInv = alphaInv
        self.alphaMod = alphaMod
        self.alphaModFix = alphaModFix
        self.alphaRepl = alphaRepl
        self.biLevel = biLevel
        self.blur = blur
        self.clrChange = clrChange
        self.clrRepl = clrRepl
        self.duotone = duotone
        self.fillOverlay = fillOverlay
        self.grayscl = grayscl
        self.hsl = hsl
        self.lum = lum
        self.tint = tint


class TileInfoProperties(Serialisable):

    tx = Integer(allow_none=True)
    ty = Integer(allow_none=True)
    sx = Integer(allow_none=True)
    sy = Integer(allow_none=True)
    flip = NoneSet(values=(['x', 'y', 'xy']))
    algn = Set(values=(['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br']))

    def __init__(self,
                 tx=None,
                 ty=None,
                 sx=None,
                 sy=None,
                 flip=None,
                 algn=None,
                ):
        self.tx = tx
        self.ty = ty
        self.sx = sx
        self.sy = sy
        self.flip = flip
        self.algn = algn


class BlipFillProperties(Serialisable):

    tagname = "blipFill"

    dpi = Integer(allow_none=True)
    rotWithShape = Bool(allow_none=True)

    blip = Typed(expected_type=Blip, allow_none=True)
    srcRect = Typed(expected_type=RelativeRect, allow_none=True)
    tile = Typed(expected_type=TileInfoProperties, allow_none=True)
    stretch = Typed(expected_type=StretchInfoProperties, allow_none=True)

    __elements__ = ("blip", "srcRect", "tile", "stretch")

    def __init__(self,
                 dpi=None,
                 rotWithShape=None,
                 blip=None,
                 tile=None,
                 stretch=StretchInfoProperties(),
                 srcRect=None,
                ):
        self.dpi = dpi
        self.rotWithShape = rotWithShape
        self.blip = blip
        self.tile = tile
        self.stretch = stretch
        self.srcRect = srcRect

''',
    },
    'openpyxl.drawing.geometry': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Integer,
    Bool,
    MinMax,
    Set,
    NoneSet,
    String,
    Alias,
)
from openpyxl.descriptors.excel import Coordinate, Percentage
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from .line import LineProperties

from openpyxl.styles.colors import Color
from openpyxl.xml.constants import DRAWING_NS


class Point2D(Serialisable):

    tagname = "off"
    namespace = DRAWING_NS

    x = Coordinate()
    y = Coordinate()

    def __init__(self,
                 x=None,
                 y=None,
                ):
        self.x = x
        self.y = y


class PositiveSize2D(Serialisable):

    tagname = "ext"
    namespace = DRAWING_NS

    """
    Dimensions in EMUs
    """

    cx = Integer()
    width = Alias('cx')
    cy = Integer()
    height = Alias('cy')

    def __init__(self,
                 cx=None,
                 cy=None,
                ):
        self.cx = cx
        self.cy = cy


class Transform2D(Serialisable):

    tagname = "xfrm"
    namespace = DRAWING_NS

    rot = Integer(allow_none=True)
    flipH = Bool(allow_none=True)
    flipV = Bool(allow_none=True)
    off = Typed(expected_type=Point2D, allow_none=True)
    ext = Typed(expected_type=PositiveSize2D, allow_none=True)
    chOff = Typed(expected_type=Point2D, allow_none=True)
    chExt = Typed(expected_type=PositiveSize2D, allow_none=True)

    __elements__ = ('off', 'ext', 'chOff', 'chExt')

    def __init__(self,
                 rot=None,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                 chOff=None,
                 chExt=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext
        self.chOff = chOff
        self.chExt = chExt


class GroupTransform2D(Serialisable):

    tagname = "xfrm"
    namespace = DRAWING_NS

    rot = Integer(allow_none=True)
    flipH = Bool(allow_none=True)
    flipV = Bool(allow_none=True)
    off = Typed(expected_type=Point2D, allow_none=True)
    ext = Typed(expected_type=PositiveSize2D, allow_none=True)
    chOff = Typed(expected_type=Point2D, allow_none=True)
    chExt = Typed(expected_type=PositiveSize2D, allow_none=True)

    __elements__ = ("off", "ext", "chOff", "chExt")

    def __init__(self,
                 rot=0,
                 flipH=None,
                 flipV=None,
                 off=None,
                 ext=None,
                 chOff=None,
                 chExt=None,
                ):
        self.rot = rot
        self.flipH = flipH
        self.flipV = flipV
        self.off = off
        self.ext = ext
        self.chOff = chOff
        self.chExt = chExt


class SphereCoords(Serialisable):

    tagname = "sphereCoords" # usually

    lat = Integer()
    lon = Integer()
    rev = Integer()

    def __init__(self,
                 lat=None,
                 lon=None,
                 rev=None,
                ):
        self.lat = lat
        self.lon = lon
        self.rev = rev


class Camera(Serialisable):

    tagname = "camera"

    prst = Set(values=[
        'legacyObliqueTopLeft', 'legacyObliqueTop', 'legacyObliqueTopRight', 'legacyObliqueLeft',
         'legacyObliqueFront', 'legacyObliqueRight', 'legacyObliqueBottomLeft',
         'legacyObliqueBottom', 'legacyObliqueBottomRight', 'legacyPerspectiveTopLeft',
         'legacyPerspectiveTop', 'legacyPerspectiveTopRight', 'legacyPerspectiveLeft',
         'legacyPerspectiveFront', 'legacyPerspectiveRight', 'legacyPerspectiveBottomLeft',
         'legacyPerspectiveBottom', 'legacyPerspectiveBottomRight', 'orthographicFront',
         'isometricTopUp', 'isometricTopDown', 'isometricBottomUp', 'isometricBottomDown',
         'isometricLeftUp', 'isometricLeftDown', 'isometricRightUp', 'isometricRightDown',
         'isometricOffAxis1Left', 'isometricOffAxis1Right', 'isometricOffAxis1Top',
         'isometricOffAxis2Left', 'isometricOffAxis2Right', 'isometricOffAxis2Top',
         'isometricOffAxis3Left', 'isometricOffAxis3Right', 'isometricOffAxis3Bottom',
         'isometricOffAxis4Left', 'isometricOffAxis4Right', 'isometricOffAxis4Bottom',
         'obliqueTopLeft',  'obliqueTop', 'obliqueTopRight', 'obliqueLeft', 'obliqueRight',
         'obliqueBottomLeft', 'obliqueBottom', 'obliqueBottomRight', 'perspectiveFront',
         'perspectiveLeft', 'perspectiveRight', 'perspectiveAbove', 'perspectiveBelow',
         'perspectiveAboveLeftFacing', 'perspectiveAboveRightFacing',
         'perspectiveContrastingLeftFacing', 'perspectiveContrastingRightFacing',
         'perspectiveHeroicLeftFacing', 'perspectiveHeroicRightFacing',
         'perspectiveHeroicExtremeLeftFacing', 'perspectiveHeroicExtremeRightFacing',
         'perspectiveRelaxed', 'perspectiveRelaxedModerately'])
    fov = Integer(allow_none=True)
    zoom = Typed(expected_type=Percentage, allow_none=True)
    rot = Typed(expected_type=SphereCoords, allow_none=True)


    def __init__(self,
                 prst=None,
                 fov=None,
                 zoom=None,
                 rot=None,
                ):
        self.prst = prst
        self.fov = fov
        self.zoom = zoom
        self.rot = rot


class LightRig(Serialisable):

    tagname = "lightRig"

    rig = Set(values=['legacyFlat1', 'legacyFlat2', 'legacyFlat3', 'legacyFlat4', 'legacyNormal1',
         'legacyNormal2', 'legacyNormal3', 'legacyNormal4', 'legacyHarsh1',
         'legacyHarsh2', 'legacyHarsh3', 'legacyHarsh4', 'threePt', 'balanced',
         'soft', 'harsh', 'flood', 'contrasting', 'morning', 'sunrise', 'sunset',
         'chilly', 'freezing', 'flat', 'twoPt', 'glow', 'brightRoom']
    )
    dir = Set(values=(['tl', 't', 'tr', 'l', 'r', 'bl', 'b', 'br']))
    rot = Typed(expected_type=SphereCoords, allow_none=True)

    def __init__(self,
                 rig=None,
                 dir=None,
                 rot=None,
                ):
        self.rig = rig
        self.dir = dir
        self.rot = rot


class Vector3D(Serialisable):

    tagname = "vector"

    dx = Integer() # can be in or universl measure :-/
    dy = Integer()
    dz = Integer()

    def __init__(self,
                 dx=None,
                 dy=None,
                 dz=None,
                ):
        self.dx = dx
        self.dy = dy
        self.dz = dz


class Point3D(Serialisable):

    tagname = "anchor"

    x = Integer()
    y = Integer()
    z = Integer()

    def __init__(self,
                 x=None,
                 y=None,
                 z=None,
                ):
        self.x = x
        self.y = y
        self.z = z


class Backdrop(Serialisable):

    anchor = Typed(expected_type=Point3D, )
    norm = Typed(expected_type=Vector3D, )
    up = Typed(expected_type=Vector3D, )
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 anchor=None,
                 norm=None,
                 up=None,
                 extLst=None,
                ):
        self.anchor = anchor
        self.norm = norm
        self.up = up
        self.extLst = extLst


class Scene3D(Serialisable):

    camera = Typed(expected_type=Camera, )
    lightRig = Typed(expected_type=LightRig, )
    backdrop = Typed(expected_type=Backdrop, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 camera=None,
                 lightRig=None,
                 backdrop=None,
                 extLst=None,
                ):
        self.camera = camera
        self.lightRig = lightRig
        self.backdrop = backdrop
        self.extLst = extLst


class Bevel(Serialisable):

    tagname = "bevel"

    w = Integer()
    h = Integer()
    prst = NoneSet(values=
               ['relaxedInset', 'circle', 'slope', 'cross', 'angle',
                'softRound', 'convex', 'coolSlant', 'divot', 'riblet',
                 'hardEdge', 'artDeco']
               )

    def __init__(self,
                 w=None,
                 h=None,
                 prst=None,
                ):
        self.w = w
        self.h = h
        self.prst = prst


class Shape3D(Serialisable):

    namespace = DRAWING_NS

    z = Typed(expected_type=Coordinate, allow_none=True)
    extrusionH = Integer(allow_none=True)
    contourW = Integer(allow_none=True)
    prstMaterial = NoneSet(values=[
        'legacyMatte','legacyPlastic', 'legacyMetal', 'legacyWireframe', 'matte', 'plastic',
        'metal', 'warmMatte', 'translucentPowder', 'powder', 'dkEdge',
        'softEdge', 'clear', 'flat', 'softmetal']
                       )
    bevelT = Typed(expected_type=Bevel, allow_none=True)
    bevelB = Typed(expected_type=Bevel, allow_none=True)
    extrusionClr = Typed(expected_type=Color, allow_none=True)
    contourClr = Typed(expected_type=Color, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 z=None,
                 extrusionH=None,
                 contourW=None,
                 prstMaterial=None,
                 bevelT=None,
                 bevelB=None,
                 extrusionClr=None,
                 contourClr=None,
                 extLst=None,
                ):
        self.z = z
        self.extrusionH = extrusionH
        self.contourW = contourW
        self.prstMaterial = prstMaterial
        self.bevelT = bevelT
        self.bevelB = bevelB
        self.extrusionClr = extrusionClr
        self.contourClr = contourClr
        self.extLst = extLst


class Path2D(Serialisable):

    w = Float()
    h = Float()
    fill = NoneSet(values=(['norm', 'lighten', 'lightenLess', 'darken', 'darkenLess']))
    stroke = Bool(allow_none=True)
    extrusionOk = Bool(allow_none=True)

    def __init__(self,
                 w=None,
                 h=None,
                 fill=None,
                 stroke=None,
                 extrusionOk=None,
                ):
        self.w = w
        self.h = h
        self.fill = fill
        self.stroke = stroke
        self.extrusionOk = extrusionOk


class Path2DList(Serialisable):

    path = Typed(expected_type=Path2D, allow_none=True)

    def __init__(self,
                 path=None,
                ):
        self.path = path


class GeomRect(Serialisable):

    l = Coordinate()
    t = Coordinate()
    r = Coordinate()
    b = Coordinate()

    def __init__(self,
                 l=None,
                 t=None,
                 r=None,
                 b=None,
                ):
        self.l = l
        self.t = t
        self.r = r
        self.b = b


class AdjPoint2D(Serialisable):

    x = Coordinate()
    y = Coordinate()

    def __init__(self,
                 x=None,
                 y=None,
                ):
        self.x = x
        self.y = y


class ConnectionSite(Serialisable):

    ang = MinMax(min=0, max=360) # guess work, can also be a name
    pos = Typed(expected_type=AdjPoint2D, )

    def __init__(self,
                 ang=None,
                 pos=None,
                ):
        self.ang = ang
        self.pos = pos


class ConnectionSiteList(Serialisable):

    cxn = Typed(expected_type=ConnectionSite, allow_none=True)

    def __init__(self,
                 cxn=None,
                ):
        self.cxn = cxn


class AdjustHandleList(Serialisable):

    pass

class GeomGuide(Serialisable):

    name = String()
    fmla = String()

    def __init__(self,
                 name=None,
                 fmla=None,
                ):
        self.name = name
        self.fmla = fmla


class GeomGuideList(Serialisable):

    gd = Typed(expected_type=GeomGuide, allow_none=True)

    def __init__(self,
                 gd=None,
                ):
        self.gd = gd


class CustomGeometry2D(Serialisable):

    avLst = Typed(expected_type=GeomGuideList, allow_none=True)
    gdLst = Typed(expected_type=GeomGuideList, allow_none=True)
    ahLst = Typed(expected_type=AdjustHandleList, allow_none=True)
    cxnLst = Typed(expected_type=ConnectionSiteList, allow_none=True)
    #rect = Typed(expected_type=GeomRect, allow_none=True)
    pathLst = Typed(expected_type=Path2DList, )

    def __init__(self,
                 avLst=None,
                 gdLst=None,
                 ahLst=None,
                 cxnLst=None,
                 rect=None,
                 pathLst=None,
                ):
        self.avLst = avLst
        self.gdLst = gdLst
        self.ahLst = ahLst
        self.cxnLst = cxnLst
        self.rect = None
        self.pathLst = pathLst


class PresetGeometry2D(Serialisable):

    namespace = DRAWING_NS

    prst = Set(values=(
        ['line', 'lineInv', 'triangle', 'rtTriangle', 'rect',
         'diamond', 'parallelogram', 'trapezoid', 'nonIsoscelesTrapezoid',
         'pentagon', 'hexagon', 'heptagon', 'octagon', 'decagon', 'dodecagon',
         'star4', 'star5', 'star6', 'star7', 'star8', 'star10', 'star12',
         'star16', 'star24', 'star32', 'roundRect', 'round1Rect',
         'round2SameRect', 'round2DiagRect', 'snipRoundRect', 'snip1Rect',
         'snip2SameRect', 'snip2DiagRect', 'plaque', 'ellipse', 'teardrop',
         'homePlate', 'chevron', 'pieWedge', 'pie', 'blockArc', 'donut',
         'noSmoking', 'rightArrow', 'leftArrow', 'upArrow', 'downArrow',
         'stripedRightArrow', 'notchedRightArrow', 'bentUpArrow',
         'leftRightArrow', 'upDownArrow', 'leftUpArrow', 'leftRightUpArrow',
         'quadArrow', 'leftArrowCallout', 'rightArrowCallout', 'upArrowCallout',
         'downArrowCallout', 'leftRightArrowCallout', 'upDownArrowCallout',
         'quadArrowCallout', 'bentArrow', 'uturnArrow', 'circularArrow',
         'leftCircularArrow', 'leftRightCircularArrow', 'curvedRightArrow',
         'curvedLeftArrow', 'curvedUpArrow', 'curvedDownArrow', 'swooshArrow',
         'cube', 'can', 'lightningBolt', 'heart', 'sun', 'moon', 'smileyFace',
         'irregularSeal1', 'irregularSeal2', 'foldedCorner', 'bevel', 'frame',
         'halfFrame', 'corner', 'diagStripe', 'chord', 'arc', 'leftBracket',
         'rightBracket', 'leftBrace', 'rightBrace', 'bracketPair', 'bracePair',
         'straightConnector1', 'bentConnector2', 'bentConnector3',
         'bentConnector4', 'bentConnector5', 'curvedConnector2',
         'curvedConnector3', 'curvedConnector4', 'curvedConnector5', 'callout1',
         'callout2', 'callout3', 'accentCallout1', 'accentCallout2',
         'accentCallout3', 'borderCallout1', 'borderCallout2', 'borderCallout3',
         'accentBorderCallout1', 'accentBorderCallout2', 'accentBorderCallout3',
         'wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout',
         'cloudCallout', 'cloud', 'ribbon', 'ribbon2', 'ellipseRibbon',
         'ellipseRibbon2', 'leftRightRibbon', 'verticalScroll',
         'horizontalScroll', 'wave', 'doubleWave', 'plus', 'flowChartProcess',
         'flowChartDecision', 'flowChartInputOutput',
         'flowChartPredefinedProcess', 'flowChartInternalStorage',
         'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator',
         'flowChartPreparation', 'flowChartManualInput',
         'flowChartManualOperation', 'flowChartConnector', 'flowChartPunchedCard',
         'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr',
         'flowChartCollate', 'flowChartSort', 'flowChartExtract',
         'flowChartMerge', 'flowChartOfflineStorage', 'flowChartOnlineStorage',
         'flowChartMagneticTape', 'flowChartMagneticDisk',
         'flowChartMagneticDrum', 'flowChartDisplay', 'flowChartDelay',
         'flowChartAlternateProcess', 'flowChartOffpageConnector',
         'actionButtonBlank', 'actionButtonHome', 'actionButtonHelp',
         'actionButtonInformation', 'actionButtonForwardNext',
         'actionButtonBackPrevious', 'actionButtonEnd', 'actionButtonBeginning',
         'actionButtonReturn', 'actionButtonDocument', 'actionButtonSound',
         'actionButtonMovie', 'gear6', 'gear9', 'funnel', 'mathPlus', 'mathMinus',
         'mathMultiply', 'mathDivide', 'mathEqual', 'mathNotEqual', 'cornerTabs',
         'squareTabs', 'plaqueTabs', 'chartX', 'chartStar', 'chartPlus']))
    avLst = Typed(expected_type=GeomGuideList, allow_none=True)

    def __init__(self,
                 prst=None,
                 avLst=None,
                ):
        self.prst = prst
        self.avLst = avLst


class FontReference(Serialisable):

    idx = NoneSet(values=(['major', 'minor']))

    def __init__(self,
                 idx=None,
                ):
        self.idx = idx


class StyleMatrixReference(Serialisable):

    idx = Integer()

    def __init__(self,
                 idx=None,
                ):
        self.idx = idx


class ShapeStyle(Serialisable):

    lnRef = Typed(expected_type=StyleMatrixReference, )
    fillRef = Typed(expected_type=StyleMatrixReference, )
    effectRef = Typed(expected_type=StyleMatrixReference, )
    fontRef = Typed(expected_type=FontReference, )

    def __init__(self,
                 lnRef=None,
                 fillRef=None,
                 effectRef=None,
                 fontRef=None,
                ):
        self.lnRef = lnRef
        self.fillRef = fillRef
        self.effectRef = effectRef
        self.fontRef = fontRef

''',
    },
    'openpyxl.drawing.graphic': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.xml.constants import CHART_NS, DRAWING_NS
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    String,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

from .effect import (
    EffectList,
    EffectContainer,
)
from .fill import (
    Blip,
    GradientFillProperties,
    BlipFillProperties,
)
from .picture import PictureFrame
from .properties import (
    NonVisualDrawingProps,
    NonVisualGroupShape,
    GroupShapeProperties,
)
from .relation import ChartRelation
from .xdr import XDRTransform2D


class GraphicFrameLocking(Serialisable):

    noGrp = Bool(allow_none=True)
    noDrilldown = Bool(allow_none=True)
    noSelect = Bool(allow_none=True)
    noChangeAspect = Bool(allow_none=True)
    noMove = Bool(allow_none=True)
    noResize = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 noGrp=None,
                 noDrilldown=None,
                 noSelect=None,
                 noChangeAspect=None,
                 noMove=None,
                 noResize=None,
                 extLst=None,
                ):
        self.noGrp = noGrp
        self.noDrilldown = noDrilldown
        self.noSelect = noSelect
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.extLst = extLst


class NonVisualGraphicFrameProperties(Serialisable):

    tagname = "cNvGraphicFramePr"

    graphicFrameLocks = Typed(expected_type=GraphicFrameLocking, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 graphicFrameLocks=None,
                 extLst=None,
                ):
        self.graphicFrameLocks = graphicFrameLocks
        self.extLst = extLst


class NonVisualGraphicFrame(Serialisable):

    tagname = "nvGraphicFramePr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps)
    cNvGraphicFramePr = Typed(expected_type=NonVisualGraphicFrameProperties)

    __elements__ = ('cNvPr', 'cNvGraphicFramePr')

    def __init__(self,
                 cNvPr=None,
                 cNvGraphicFramePr=None,
                ):
        if cNvPr is None:
            cNvPr = NonVisualDrawingProps(id=0, name="Chart 0")
        self.cNvPr = cNvPr
        if cNvGraphicFramePr is None:
            cNvGraphicFramePr = NonVisualGraphicFrameProperties()
        self.cNvGraphicFramePr = cNvGraphicFramePr


class GraphicData(Serialisable):

    tagname = "graphicData"
    namespace = DRAWING_NS

    uri = String()
    chart = Typed(expected_type=ChartRelation, allow_none=True)


    def __init__(self,
                 uri=CHART_NS,
                 chart=None,
                ):
        self.uri = uri
        self.chart = chart


class GraphicObject(Serialisable):

    tagname = "graphic"
    namespace = DRAWING_NS

    graphicData = Typed(expected_type=GraphicData)

    def __init__(self,
                 graphicData=None,
                ):
        if graphicData is None:
            graphicData = GraphicData()
        self.graphicData = graphicData


class GraphicFrame(Serialisable):

    tagname = "graphicFrame"

    nvGraphicFramePr = Typed(expected_type=NonVisualGraphicFrame)
    xfrm = Typed(expected_type=XDRTransform2D)
    graphic = Typed(expected_type=GraphicObject)
    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)

    __elements__ = ('nvGraphicFramePr', 'xfrm', 'graphic', 'macro', 'fPublished')

    def __init__(self,
                 nvGraphicFramePr=None,
                 xfrm=None,
                 graphic=None,
                 macro=None,
                 fPublished=None,
                 ):
        if nvGraphicFramePr is None:
            nvGraphicFramePr = NonVisualGraphicFrame()
        self.nvGraphicFramePr = nvGraphicFramePr
        if xfrm is None:
            xfrm = XDRTransform2D()
        self.xfrm = xfrm
        if graphic is None:
            graphic = GraphicObject()
        self.graphic = graphic
        self.macro = macro
        self.fPublished = fPublished


class GroupShape(Serialisable):

    nvGrpSpPr = Typed(expected_type=NonVisualGroupShape)
    nonVisualProperties = Alias("nvGrpSpPr")
    grpSpPr = Typed(expected_type=GroupShapeProperties)
    visualProperties = Alias("grpSpPr")
    pic = Typed(expected_type=PictureFrame, allow_none=True)

    __elements__ = ["nvGrpSpPr", "grpSpPr", "pic"]

    def __init__(self,
                 nvGrpSpPr=None,
                 grpSpPr=None,
                 pic=None,
                ):
        self.nvGrpSpPr = nvGrpSpPr
        self.grpSpPr = grpSpPr
        self.pic = pic

""",
    },
    'openpyxl.drawing.image': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from io import BytesIO

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = False


def _import_image(img):
    if not PILImage:
        raise ImportError('You must install Pillow to fetch image objects')

    if not isinstance(img, PILImage.Image):
        img = PILImage.open(img)

    return img


class Image:
    """Image in a spreadsheet"""

    _id = 1
    _path = "/xl/media/image{0}.{1}"
    anchor = "A1"

    def __init__(self, img):

        self.ref = img
        mark_to_close = isinstance(img, str)
        image = _import_image(img)
        self.width, self.height = image.size

        try:
            self.format = image.format.lower()
        except AttributeError:
            self.format = "png"
        if mark_to_close:
            # PIL instances created for metadata should be closed.
            image.close()


    def _data(self):
        """
        Return image data, convert to supported types if necessary
        """
        img = _import_image(self.ref)
        # don't convert these file formats
        if self.format in ['gif', 'jpeg', 'png']:
            img.fp.seek(0)
            fp = img.fp
        else:
            fp = BytesIO()
            img.save(fp, format="png")
            fp.seek(0)

        data = fp.read()
        fp.close()
        return data


    @property
    def path(self):
        return self._path.format(self._id, self.format)

''',
    },
    'openpyxl.drawing.line': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    MinMax,
    NoneSet,
    Alias,
    Sequence
)

from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedNoneSet,
    EmptyTag,
)
from openpyxl.xml.constants import DRAWING_NS

from .colors import ColorChoiceDescriptor
from .fill import GradientFillProperties, PatternFillProperties
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

"""
Line elements from drawing main schema
"""


class LineEndProperties(Serialisable):

    tagname = "end"
    namespace = DRAWING_NS

    type = NoneSet(values=(['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow']))
    w = NoneSet(values=(['sm', 'med', 'lg']))
    len = NoneSet(values=(['sm', 'med', 'lg']))

    def __init__(self,
                 type=None,
                 w=None,
                 len=None,
                ):
        self.type = type
        self.w = w
        self.len = len


class DashStop(Serialisable):

    tagname = "ds"
    namespace = DRAWING_NS

    d = Integer()
    length = Alias('d')
    sp = Integer()
    space = Alias('sp')

    def __init__(self,
                 d=0,
                 sp=0,
                ):
        self.d = d
        self.sp = sp


class DashStopList(Serialisable):

    ds = Sequence(expected_type=DashStop, allow_none=True)

    def __init__(self,
                 ds=None,
                ):
        self.ds = ds


class LineProperties(Serialisable):

    tagname = "ln"
    namespace = DRAWING_NS

    w = MinMax(min=0, max=20116800, allow_none=True) # EMU
    width = Alias('w')
    cap = NoneSet(values=(['rnd', 'sq', 'flat']))
    cmpd = NoneSet(values=(['sng', 'dbl', 'thickThin', 'thinThick', 'tri']))
    algn = NoneSet(values=(['ctr', 'in']))

    noFill = EmptyTag()
    solidFill = ColorChoiceDescriptor()
    gradFill = Typed(expected_type=GradientFillProperties, allow_none=True)
    pattFill = Typed(expected_type=PatternFillProperties, allow_none=True)

    prstDash = NestedNoneSet(values=(['solid', 'dot', 'dash', 'lgDash', 'dashDot',
                       'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot',
                       'sysDashDotDot']), namespace=namespace)
    dashStyle = Alias('prstDash')

    custDash = Typed(expected_type=DashStop, allow_none=True)

    round = EmptyTag()
    bevel = EmptyTag()
    miter = NestedInteger(allow_none=True, attribute="lim")

    headEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    tailEnd = Typed(expected_type=LineEndProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ('noFill', 'solidFill', 'gradFill', 'pattFill',
                    'prstDash', 'custDash', 'round', 'bevel', 'miter', 'headEnd', 'tailEnd')

    def __init__(self,
                 w=None,
                 cap=None,
                 cmpd=None,
                 algn=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 pattFill=None,
                 prstDash=None,
                 custDash=None,
                 round=None,
                 bevel=None,
                 miter=None,
                 headEnd=None,
                 tailEnd=None,
                 extLst=None,
                ):
        self.w = w
        self.cap = cap
        self.cmpd = cmpd
        self.algn = algn
        self.noFill = noFill
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.pattFill = pattFill
        if prstDash is None:
            prstDash = "solid"
        self.prstDash = prstDash
        self.custDash = custDash
        self.round = round
        self.bevel = bevel
        self.miter = miter
        self.headEnd = headEnd
        self.tailEnd = tailEnd

''',
    },
    'openpyxl.drawing.picture': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.xml.constants import DRAWING_NS

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    String,
    Alias,
)
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

from openpyxl.chart.shapes import GraphicalProperties

from .fill import BlipFillProperties
from .properties import NonVisualDrawingProps
from .geometry import ShapeStyle


class PictureLocking(Serialisable):

    tagname = "picLocks"
    namespace = DRAWING_NS

    # Using attribute group AG_Locking
    noCrop = Bool(allow_none=True)
    noGrp = Bool(allow_none=True)
    noSelect = Bool(allow_none=True)
    noRot = Bool(allow_none=True)
    noChangeAspect = Bool(allow_none=True)
    noMove = Bool(allow_none=True)
    noResize = Bool(allow_none=True)
    noEditPoints = Bool(allow_none=True)
    noAdjustHandles = Bool(allow_none=True)
    noChangeArrowheads = Bool(allow_none=True)
    noChangeShapeType = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 noCrop=None,
                 noGrp=None,
                 noSelect=None,
                 noRot=None,
                 noChangeAspect=None,
                 noMove=None,
                 noResize=None,
                 noEditPoints=None,
                 noAdjustHandles=None,
                 noChangeArrowheads=None,
                 noChangeShapeType=None,
                 extLst=None,
                ):
        self.noCrop = noCrop
        self.noGrp = noGrp
        self.noSelect = noSelect
        self.noRot = noRot
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.noEditPoints = noEditPoints
        self.noAdjustHandles = noAdjustHandles
        self.noChangeArrowheads = noChangeArrowheads
        self.noChangeShapeType = noChangeShapeType


class NonVisualPictureProperties(Serialisable):

    tagname = "cNvPicPr"

    preferRelativeResize = Bool(allow_none=True)
    picLocks = Typed(expected_type=PictureLocking, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ("picLocks",)

    def __init__(self,
                 preferRelativeResize=None,
                 picLocks=None,
                 extLst=None,
                ):
        self.preferRelativeResize = preferRelativeResize
        self.picLocks = picLocks


class PictureNonVisual(Serialisable):

    tagname = "nvPicPr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps, )
    cNvPicPr = Typed(expected_type=NonVisualPictureProperties, )

    __elements__ = ("cNvPr", "cNvPicPr")

    def __init__(self,
                 cNvPr=None,
                 cNvPicPr=None,
                ):
        if cNvPr is None:
            cNvPr = NonVisualDrawingProps(id=0, name="Image 1", descr="Name of file")
        self.cNvPr = cNvPr
        if cNvPicPr is None:
            cNvPicPr = NonVisualPictureProperties()
        self.cNvPicPr = cNvPicPr




class PictureFrame(Serialisable):

    tagname = "pic"

    macro = String(allow_none=True)
    fPublished = Bool(allow_none=True)
    nvPicPr = Typed(expected_type=PictureNonVisual, )
    blipFill = Typed(expected_type=BlipFillProperties, )
    spPr = Typed(expected_type=GraphicalProperties, )
    graphicalProperties = Alias('spPr')
    style = Typed(expected_type=ShapeStyle, allow_none=True)

    __elements__ = ("nvPicPr", "blipFill", "spPr", "style")

    def __init__(self,
                 macro=None,
                 fPublished=None,
                 nvPicPr=None,
                 blipFill=None,
                 spPr=None,
                 style=None,
                ):
        self.macro = macro
        self.fPublished = fPublished
        if nvPicPr is None:
            nvPicPr = PictureNonVisual()
        self.nvPicPr = nvPicPr
        if blipFill is None:
            blipFill = BlipFillProperties()
        self.blipFill = blipFill
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.style = style

""",
    },
    'openpyxl.drawing.properties': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.xml.constants import DRAWING_NS
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    Set,
    String,
    Alias,
    NoneSet,
)
from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

from .geometry import GroupTransform2D, Scene3D
from .text import Hyperlink


class GroupShapeProperties(Serialisable):

    tagname = "grpSpPr"

    bwMode = NoneSet(values=(['clr', 'auto', 'gray', 'ltGray', 'invGray',
                          'grayWhite', 'blackGray', 'blackWhite', 'black', 'white', 'hidden']))
    xfrm = Typed(expected_type=GroupTransform2D, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bwMode=None,
                 xfrm=None,
                 scene3d=None,
                 extLst=None,
                ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.scene3d = scene3d
        self.extLst = extLst


class GroupLocking(Serialisable):

    tagname = "grpSpLocks"
    namespace = DRAWING_NS

    noGrp = Bool(allow_none=True)
    noUngrp = Bool(allow_none=True)
    noSelect = Bool(allow_none=True)
    noRot = Bool(allow_none=True)
    noChangeAspect = Bool(allow_none=True)
    noMove = Bool(allow_none=True)
    noResize = Bool(allow_none=True)
    noChangeArrowheads = Bool(allow_none=True)
    noEditPoints = Bool(allow_none=True)
    noAdjustHandles = Bool(allow_none=True)
    noChangeArrowheads = Bool(allow_none=True)
    noChangeShapeType = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 noGrp=None,
                 noUngrp=None,
                 noSelect=None,
                 noRot=None,
                 noChangeAspect=None,
                 noChangeArrowheads=None,
                 noMove=None,
                 noResize=None,
                 noEditPoints=None,
                 noAdjustHandles=None,
                 noChangeShapeType=None,
                 extLst=None,
                ):
        self.noGrp = noGrp
        self.noUngrp = noUngrp
        self.noSelect = noSelect
        self.noRot = noRot
        self.noChangeAspect = noChangeAspect
        self.noChangeArrowheads = noChangeArrowheads
        self.noMove = noMove
        self.noResize = noResize
        self.noEditPoints = noEditPoints
        self.noAdjustHandles = noAdjustHandles
        self.noChangeShapeType = noChangeShapeType


class NonVisualGroupDrawingShapeProps(Serialisable):

    tagname = "cNvGrpSpPr"

    grpSpLocks = Typed(expected_type=GroupLocking, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ("grpSpLocks",)

    def __init__(self,
                 grpSpLocks=None,
                 extLst=None,
                ):
        self.grpSpLocks = grpSpLocks


class NonVisualDrawingShapeProps(Serialisable):

    tagname = "cNvSpPr"

    spLocks = Typed(expected_type=GroupLocking, allow_none=True)
    txBax = Bool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ("spLocks", "txBax")

    def __init__(self,
                 spLocks=None,
                 txBox=None,
                 extLst=None,
                ):
        self.spLocks = spLocks
        self.txBox = txBox


class NonVisualDrawingProps(Serialisable):

    tagname = "cNvPr"

    id = Integer()
    name = String()
    descr = String(allow_none=True)
    hidden = Bool(allow_none=True)
    title = String(allow_none=True)
    hlinkClick = Typed(expected_type=Hyperlink, allow_none=True)
    hlinkHover = Typed(expected_type=Hyperlink, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ["hlinkClick", "hlinkHover"]

    def __init__(self,
                 id=None,
                 name=None,
                 descr=None,
                 hidden=None,
                 title=None,
                 hlinkClick=None,
                 hlinkHover=None,
                 extLst=None,
                ):
        self.id = id
        self.name = name
        self.descr = descr
        self.hidden = hidden
        self.title = title
        self.hlinkClick = hlinkClick
        self.hlinkHover = hlinkHover
        self.extLst = extLst

class NonVisualGroupShape(Serialisable):

    tagname = "nvGrpSpPr"

    cNvPr = Typed(expected_type=NonVisualDrawingProps)
    cNvGrpSpPr = Typed(expected_type=NonVisualGroupDrawingShapeProps)

    __elements__ = ("cNvPr", "cNvGrpSpPr")

    def __init__(self,
                 cNvPr=None,
                 cNvGrpSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvGrpSpPr = cNvGrpSpPr


""",
    },
    'openpyxl.drawing.relation': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.xml.constants import CHART_NS

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.excel import Relation


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id = Relation()

    def __init__(self, id):
        self.id = id

""",
    },
    'openpyxl.drawing.spreadsheet_drawing': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    NoneSet,
    Integer,
    Sequence,
    Alias,
)
from openpyxl.descriptors.nested import (
    NestedText,
    NestedNoneSet,
)
from openpyxl.descriptors.excel import Relation

from openpyxl.packaging.relationship import (
    Relationship,
    RelationshipList,
)
from openpyxl.utils import coordinate_to_tuple
from openpyxl.utils.units import (
    cm_to_EMU,
    pixels_to_EMU,
)
from openpyxl.drawing.image import Image

from openpyxl.xml.constants import SHEET_DRAWING_NS

from openpyxl.chart._chart import ChartBase
from .xdr import (
    XDRPoint2D,
    XDRPositiveSize2D,
)
from .fill import Blip
from .connector import Shape
from .graphic import (
    GroupShape,
    GraphicFrame,
    )
from .geometry import PresetGeometry2D
from .picture import PictureFrame
from .relation import ChartRelation


class AnchorClientData(Serialisable):

    fLocksWithSheet = Bool(allow_none=True)
    fPrintsWithSheet = Bool(allow_none=True)

    def __init__(self,
                 fLocksWithSheet=None,
                 fPrintsWithSheet=None,
                 ):
        self.fLocksWithSheet = fLocksWithSheet
        self.fPrintsWithSheet = fPrintsWithSheet


class AnchorMarker(Serialisable):

    tagname = "marker"

    col = NestedText(expected_type=int)
    colOff = NestedText(expected_type=int)
    row = NestedText(expected_type=int)
    rowOff = NestedText(expected_type=int)

    def __init__(self,
                 col=0,
                 colOff=0,
                 row=0,
                 rowOff=0,
                 ):
        self.col = col
        self.colOff = colOff
        self.row = row
        self.rowOff = rowOff


class _AnchorBase(Serialisable):

    #one of
    sp = Typed(expected_type=Shape, allow_none=True)
    shape = Alias("sp")
    grpSp = Typed(expected_type=GroupShape, allow_none=True)
    groupShape = Alias("grpSp")
    graphicFrame = Typed(expected_type=GraphicFrame, allow_none=True)
    cxnSp = Typed(expected_type=Shape, allow_none=True)
    connectionShape = Alias("cxnSp")
    pic = Typed(expected_type=PictureFrame, allow_none=True)
    contentPart = Relation()

    clientData = Typed(expected_type=AnchorClientData)

    __elements__ = ('sp', 'grpSp', 'graphicFrame',
                    'cxnSp', 'pic', 'contentPart', 'clientData')

    def __init__(self,
                 clientData=None,
                 sp=None,
                 grpSp=None,
                 graphicFrame=None,
                 cxnSp=None,
                 pic=None,
                 contentPart=None
                 ):
        if clientData is None:
            clientData = AnchorClientData()
        self.clientData = clientData
        self.sp = sp
        self.grpSp = grpSp
        self.graphicFrame = graphicFrame
        self.cxnSp = cxnSp
        self.pic = pic
        self.contentPart = contentPart


class AbsoluteAnchor(_AnchorBase):

    tagname = "absoluteAnchor"

    pos = Typed(expected_type=XDRPoint2D)
    ext = Typed(expected_type=XDRPositiveSize2D)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('pos', 'ext') + _AnchorBase.__elements__

    def __init__(self,
                 pos=None,
                 ext=None,
                 **kw
                ):
        if pos is None:
            pos = XDRPoint2D(0, 0)
        self.pos = pos
        if ext is None:
            ext = XDRPositiveSize2D(0, 0)
        self.ext = ext
        super().__init__(**kw)


class OneCellAnchor(_AnchorBase):

    tagname = "oneCellAnchor"

    _from = Typed(expected_type=AnchorMarker)
    ext = Typed(expected_type=XDRPositiveSize2D)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('_from', 'ext') + _AnchorBase.__elements__


    def __init__(self,
                 _from=None,
                 ext=None,
                 **kw
                ):
        if _from is None:
            _from = AnchorMarker()
        self._from = _from
        if ext is None:
            ext = XDRPositiveSize2D(0, 0)
        self.ext = ext
        super().__init__(**kw)


class TwoCellAnchor(_AnchorBase):

    tagname = "twoCellAnchor"

    editAs = NoneSet(values=(['twoCell', 'oneCell', 'absolute']))
    _from = Typed(expected_type=AnchorMarker)
    to = Typed(expected_type=AnchorMarker)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('_from', 'to') + _AnchorBase.__elements__

    def __init__(self,
                 editAs=None,
                 _from=None,
                 to=None,
                 **kw
                 ):
        self.editAs = editAs
        if _from is None:
            _from = AnchorMarker()
        self._from = _from
        if to is None:
            to = AnchorMarker()
        self.to = to
        super().__init__(**kw)


def _check_anchor(obj):
    """
    Check whether an object has an existing Anchor object
    If not create a OneCellAnchor using the provided coordinate
    """
    anchor = obj.anchor
    if not isinstance(anchor, _AnchorBase):
        row, col = coordinate_to_tuple(anchor.upper())
        anchor = OneCellAnchor()
        anchor._from.row = row -1
        anchor._from.col = col -1
        if isinstance(obj, ChartBase):
            anchor.ext.width = cm_to_EMU(obj.width)
            anchor.ext.height = cm_to_EMU(obj.height)
        elif isinstance(obj, Image):
            anchor.ext.width = pixels_to_EMU(obj.width)
            anchor.ext.height = pixels_to_EMU(obj.height)
    return anchor


class SpreadsheetDrawing(Serialisable):

    tagname = "wsDr"
    mime_type = "application/vnd.openxmlformats-officedocument.drawing+xml"
    _rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    _path = PartName="/xl/drawings/drawing{0}.xml"
    _id = None

    twoCellAnchor = Sequence(expected_type=TwoCellAnchor, allow_none=True)
    oneCellAnchor = Sequence(expected_type=OneCellAnchor, allow_none=True)
    absoluteAnchor = Sequence(expected_type=AbsoluteAnchor, allow_none=True)

    __elements__ = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")

    def __init__(self,
                 twoCellAnchor=(),
                 oneCellAnchor=(),
                 absoluteAnchor=(),
                 ):
        self.twoCellAnchor = twoCellAnchor
        self.oneCellAnchor = oneCellAnchor
        self.absoluteAnchor = absoluteAnchor
        self.charts = []
        self.images = []
        self._rels = []


    def __hash__(self):
        """
        Just need to check for identity
        """
        return id(self)


    def __bool__(self):
        return bool(self.charts) or bool(self.images)



    def _write(self):
        """
        create required structure and the serialise
        """
        anchors = []
        for idx, obj in enumerate(self.charts + self.images, 1):
            anchor = _check_anchor(obj)
            if isinstance(obj, ChartBase):
                rel = Relationship(type="chart", Target=obj.path)
                anchor.graphicFrame = self._chart_frame(idx)
            elif isinstance(obj, Image):
                rel = Relationship(type="image", Target=obj.path)
                child = anchor.pic or anchor.groupShape and anchor.groupShape.pic
                if not child:
                    anchor.pic = self._picture_frame(idx)
                else:
                    child.blipFill.blip.embed = "rId{0}".format(idx)

            anchors.append(anchor)
            self._rels.append(rel)

        for a in anchors:
            if isinstance(a, OneCellAnchor):
                self.oneCellAnchor.append(a)
            elif isinstance(a, TwoCellAnchor):
                self.twoCellAnchor.append(a)
            else:
                self.absoluteAnchor.append(a)

        tree = self.to_tree()
        tree.set('xmlns', SHEET_DRAWING_NS)
        return tree


    def _chart_frame(self, idx):
        chart_rel = ChartRelation(f"rId{idx}")
        frame = GraphicFrame()
        nv = frame.nvGraphicFramePr.cNvPr
        nv.id = idx
        nv.name = "Chart {0}".format(idx)
        frame.graphic.graphicData.chart = chart_rel
        return frame


    def _picture_frame(self, idx):
        pic = PictureFrame()
        pic.nvPicPr.cNvPr.descr = "Picture"
        pic.nvPicPr.cNvPr.id = idx
        pic.nvPicPr.cNvPr.name = "Image {0}".format(idx)

        pic.blipFill.blip = Blip()
        pic.blipFill.blip.embed = "rId{0}".format(idx)
        pic.blipFill.blip.cstate = "print"

        pic.spPr.prstGeom = PresetGeometry2D(prst="rect")
        pic.spPr.ln = None
        return pic


    def _write_rels(self):
        rels = RelationshipList()
        for r in self._rels:
            rels.append(r)
        return rels.to_tree()


    @property
    def path(self):
        return self._path.format(self._id)


    @property
    def _chart_rels(self):
        """
        Get relationship information for each chart and bind anchor to it
        """
        rels = []
        anchors = self.absoluteAnchor + self.oneCellAnchor + self.twoCellAnchor
        for anchor in anchors:
            if anchor.graphicFrame is not None:
                graphic = anchor.graphicFrame.graphic
                rel = graphic.graphicData.chart
                if rel is not None:
                    rel.anchor = anchor
                    rel.anchor.graphicFrame = None
                    rels.append(rel)
        return rels


    @property
    def _blip_rels(self):
        """
        Get relationship information for each blip and bind anchor to it

        Images that are not part of the XLSX package will be ignored.
        """
        rels = []
        anchors = self.absoluteAnchor + self.oneCellAnchor + self.twoCellAnchor

        for anchor in anchors:
            child = anchor.pic or anchor.groupShape and anchor.groupShape.pic
            if child and child.blipFill:
                rel = child.blipFill.blip
                if rel is not None and rel.embed:
                    rel.anchor = anchor
                    rels.append(rel)

        return rels

''',
    },
    'openpyxl.drawing.text': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    Set,
    NoneSet,
    Sequence,
    String,
    Bool,
    MinMax,
    Integer
)
from openpyxl.descriptors.excel import (
    HexBinary,
    Coordinate,
    Relation,
)
from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedText,
    NestedValue,
    EmptyTag
)
from openpyxl.xml.constants import DRAWING_NS


from .colors import ColorChoiceDescriptor
from .effect import (
    EffectList,
    EffectContainer,
)
from .fill import(
    GradientFillProperties,
    BlipFillProperties,
    PatternFillProperties,
    Blip
)
from .geometry import (
    LineProperties,
    Color,
    Scene3D
)

from openpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from openpyxl.descriptors.nested import NestedBool


class EmbeddedWAVAudioFile(Serialisable):

    name = String(allow_none=True)

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Hyperlink(Serialisable):

    tagname = "hlinkClick"
    namespace = DRAWING_NS

    invalidUrl = String(allow_none=True)
    action = String(allow_none=True)
    tgtFrame = String(allow_none=True)
    tooltip = String(allow_none=True)
    history = Bool(allow_none=True)
    highlightClick = Bool(allow_none=True)
    endSnd = Bool(allow_none=True)
    snd = Typed(expected_type=EmbeddedWAVAudioFile, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)
    id = Relation(allow_none=True)

    __elements__ = ('snd',)

    def __init__(self,
                 invalidUrl=None,
                 action=None,
                 tgtFrame=None,
                 tooltip=None,
                 history=None,
                 highlightClick=None,
                 endSnd=None,
                 snd=None,
                 extLst=None,
                 id=None,
                ):
        self.invalidUrl = invalidUrl
        self.action = action
        self.tgtFrame = tgtFrame
        self.tooltip = tooltip
        self.history = history
        self.highlightClick = highlightClick
        self.endSnd = endSnd
        self.snd = snd
        self.id = id


class Font(Serialisable):

    tagname = "latin"
    namespace = DRAWING_NS

    typeface = String()
    panose = HexBinary(allow_none=True)
    pitchFamily = MinMax(min=0, max=52, allow_none=True)
    charset = Integer(allow_none=True)

    def __init__(self,
                 typeface=None,
                 panose=None,
                 pitchFamily=None,
                 charset=None,
                ):
        self.typeface = typeface
        self.panose = panose
        self.pitchFamily = pitchFamily
        self.charset = charset


class CharacterProperties(Serialisable):

    tagname = "defRPr"
    namespace = DRAWING_NS

    kumimoji = Bool(allow_none=True)
    lang = String(allow_none=True)
    altLang = String(allow_none=True)
    sz = MinMax(allow_none=True, min=100, max=400000) # 100ths of a point
    b = Bool(allow_none=True)
    i = Bool(allow_none=True)
    u = NoneSet(values=(['words', 'sng', 'dbl', 'heavy', 'dotted',
                         'dottedHeavy', 'dash', 'dashHeavy', 'dashLong', 'dashLongHeavy',
                         'dotDash', 'dotDashHeavy', 'dotDotDash', 'dotDotDashHeavy', 'wavy',
                         'wavyHeavy', 'wavyDbl']))
    strike = NoneSet(values=(['noStrike', 'sngStrike', 'dblStrike']))
    kern = Integer(allow_none=True)
    cap = NoneSet(values=(['small', 'all']))
    spc = Integer(allow_none=True)
    normalizeH = Bool(allow_none=True)
    baseline = Integer(allow_none=True)
    noProof = Bool(allow_none=True)
    dirty = Bool(allow_none=True)
    err = Bool(allow_none=True)
    smtClean = Bool(allow_none=True)
    smtId = Integer(allow_none=True)
    bmk = String(allow_none=True)
    ln = Typed(expected_type=LineProperties, allow_none=True)
    highlight = Typed(expected_type=Color, allow_none=True)
    latin = Typed(expected_type=Font, allow_none=True)
    ea = Typed(expected_type=Font, allow_none=True)
    cs = Typed(expected_type=Font, allow_none=True)
    sym = Typed(expected_type=Font, allow_none=True)
    hlinkClick = Typed(expected_type=Hyperlink, allow_none=True)
    hlinkMouseOver = Typed(expected_type=Hyperlink, allow_none=True)
    rtl = NestedBool(allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)
    # uses element group EG_FillProperties
    noFill = EmptyTag(namespace=DRAWING_NS)
    solidFill = ColorChoiceDescriptor()
    gradFill = Typed(expected_type=GradientFillProperties, allow_none=True)
    blipFill = Typed(expected_type=BlipFillProperties, allow_none=True)
    pattFill = Typed(expected_type=PatternFillProperties, allow_none=True)
    grpFill = EmptyTag(namespace=DRAWING_NS)
    # uses element group EG_EffectProperties
    effectLst = Typed(expected_type=EffectList, allow_none=True)
    effectDag = Typed(expected_type=EffectContainer, allow_none=True)
    # uses element group EG_TextUnderlineLine
    uLnTx = EmptyTag()
    uLn = Typed(expected_type=LineProperties, allow_none=True)
    # uses element group EG_TextUnderlineFill
    uFillTx = EmptyTag()
    uFill = EmptyTag()

    __elements__ = ('ln', 'noFill', 'solidFill', 'gradFill', 'blipFill',
                    'pattFill', 'grpFill', 'effectLst', 'effectDag', 'highlight','uLnTx',
                    'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick',
                    'hlinkMouseOver', 'rtl', )

    def __init__(self,
                 kumimoji=None,
                 lang=None,
                 altLang=None,
                 sz=None,
                 b=None,
                 i=None,
                 u=None,
                 strike=None,
                 kern=None,
                 cap=None,
                 spc=None,
                 normalizeH=None,
                 baseline=None,
                 noProof=None,
                 dirty=None,
                 err=None,
                 smtClean=None,
                 smtId=None,
                 bmk=None,
                 ln=None,
                 highlight=None,
                 latin=None,
                 ea=None,
                 cs=None,
                 sym=None,
                 hlinkClick=None,
                 hlinkMouseOver=None,
                 rtl=None,
                 extLst=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 blipFill=None,
                 pattFill=None,
                 grpFill=None,
                 effectLst=None,
                 effectDag=None,
                 uLnTx=None,
                 uLn=None,
                 uFillTx=None,
                 uFill=None,
                ):
        self.kumimoji = kumimoji
        self.lang = lang
        self.altLang = altLang
        self.sz = sz
        self.b = b
        self.i = i
        self.u = u
        self.strike = strike
        self.kern = kern
        self.cap = cap
        self.spc = spc
        self.normalizeH = normalizeH
        self.baseline = baseline
        self.noProof = noProof
        self.dirty = dirty
        self.err = err
        self.smtClean = smtClean
        self.smtId = smtId
        self.bmk = bmk
        self.ln = ln
        self.highlight = highlight
        self.latin = latin
        self.ea = ea
        self.cs = cs
        self.sym = sym
        self.hlinkClick = hlinkClick
        self.hlinkMouseOver = hlinkMouseOver
        self.rtl = rtl
        self.noFill = noFill
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.blipFill = blipFill
        self.pattFill = pattFill
        self.grpFill = grpFill
        self.effectLst = effectLst
        self.effectDag = effectDag
        self.uLnTx = uLnTx
        self.uLn = uLn
        self.uFillTx = uFillTx
        self.uFill = uFill


class TabStop(Serialisable):

    pos = Typed(expected_type=Coordinate, allow_none=True)
    algn = Typed(expected_type=Set(values=(['l', 'ctr', 'r', 'dec'])))

    def __init__(self,
                 pos=None,
                 algn=None,
                ):
        self.pos = pos
        self.algn = algn


class TabStopList(Serialisable):

    tab = Typed(expected_type=TabStop, allow_none=True)

    def __init__(self,
                 tab=None,
                ):
        self.tab = tab


class Spacing(Serialisable):

    spcPct = NestedInteger(allow_none=True)
    spcPts = NestedInteger(allow_none=True)

    __elements__ = ('spcPct', 'spcPts')

    def __init__(self,
                 spcPct=None,
                 spcPts=None,
                 ):
        self.spcPct = spcPct
        self.spcPts = spcPts


class AutonumberBullet(Serialisable):

    type = Set(values=(['alphaLcParenBoth', 'alphaUcParenBoth',
                        'alphaLcParenR', 'alphaUcParenR', 'alphaLcPeriod', 'alphaUcPeriod',
                        'arabicParenBoth', 'arabicParenR', 'arabicPeriod', 'arabicPlain',
                        'romanLcParenBoth', 'romanUcParenBoth', 'romanLcParenR', 'romanUcParenR',
                        'romanLcPeriod', 'romanUcPeriod', 'circleNumDbPlain',
                        'circleNumWdBlackPlain', 'circleNumWdWhitePlain', 'arabicDbPeriod',
                        'arabicDbPlain', 'ea1ChsPeriod', 'ea1ChsPlain', 'ea1ChtPeriod',
                        'ea1ChtPlain', 'ea1JpnChsDbPeriod', 'ea1JpnKorPlain', 'ea1JpnKorPeriod',
                        'arabic1Minus', 'arabic2Minus', 'hebrew2Minus', 'thaiAlphaPeriod',
                        'thaiAlphaParenR', 'thaiAlphaParenBoth', 'thaiNumPeriod',
                        'thaiNumParenR', 'thaiNumParenBoth', 'hindiAlphaPeriod',
                        'hindiNumPeriod', 'hindiNumParenR', 'hindiAlpha1Period']))
    startAt = Integer()

    def __init__(self,
                 type=None,
                 startAt=None,
                ):
        self.type = type
        self.startAt = startAt


class ParagraphProperties(Serialisable):

    tagname = "pPr"
    namespace = DRAWING_NS

    marL = Integer(allow_none=True)
    marR = Integer(allow_none=True)
    lvl = Integer(allow_none=True)
    indent = Integer(allow_none=True)
    algn = NoneSet(values=(['l', 'ctr', 'r', 'just', 'justLow', 'dist', 'thaiDist']))
    defTabSz = Integer(allow_none=True)
    rtl = Bool(allow_none=True)
    eaLnBrk = Bool(allow_none=True)
    fontAlgn = NoneSet(values=(['auto', 't', 'ctr', 'base', 'b']))
    latinLnBrk = Bool(allow_none=True)
    hangingPunct = Bool(allow_none=True)

    # uses element group EG_TextBulletColor
    # uses element group EG_TextBulletSize
    # uses element group EG_TextBulletTypeface
    # uses element group EG_TextBullet
    lnSpc = Typed(expected_type=Spacing, allow_none=True)
    spcBef = Typed(expected_type=Spacing, allow_none=True)
    spcAft = Typed(expected_type=Spacing, allow_none=True)
    tabLst = Typed(expected_type=TabStopList, allow_none=True)
    defRPr = Typed(expected_type=CharacterProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)
    buClrTx = EmptyTag()
    buClr = Typed(expected_type=Color, allow_none=True)
    buSzTx = EmptyTag()
    buSzPct = NestedInteger(allow_none=True)
    buSzPts = NestedInteger(allow_none=True)
    buFontTx = EmptyTag()
    buFont = Typed(expected_type=Font, allow_none=True)
    buNone = EmptyTag()
    buAutoNum = EmptyTag()
    buChar = NestedValue(expected_type=str, attribute="char", allow_none=True)
    buBlip = NestedValue(expected_type=Blip, attribute="blip", allow_none=True)

    __elements__ = ('lnSpc', 'spcBef', 'spcAft', 'tabLst', 'defRPr',
                    'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont',
                    'buNone', 'buAutoNum', 'buChar', 'buBlip')

    def __init__(self,
                 marL=None,
                 marR=None,
                 lvl=None,
                 indent=None,
                 algn=None,
                 defTabSz=None,
                 rtl=None,
                 eaLnBrk=None,
                 fontAlgn=None,
                 latinLnBrk=None,
                 hangingPunct=None,
                 lnSpc=None,
                 spcBef=None,
                 spcAft=None,
                 tabLst=None,
                 defRPr=None,
                 extLst=None,
                 buClrTx=None,
                 buClr=None,
                 buSzTx=None,
                 buSzPct=None,
                 buSzPts=None,
                 buFontTx=None,
                 buFont=None,
                 buNone=None,
                 buAutoNum=None,
                 buChar=None,
                 buBlip=None,
                 ):
        self.marL = marL
        self.marR = marR
        self.lvl = lvl
        self.indent = indent
        self.algn = algn
        self.defTabSz = defTabSz
        self.rtl = rtl
        self.eaLnBrk = eaLnBrk
        self.fontAlgn = fontAlgn
        self.latinLnBrk = latinLnBrk
        self.hangingPunct = hangingPunct
        self.lnSpc = lnSpc
        self.spcBef = spcBef
        self.spcAft = spcAft
        self.tabLst = tabLst
        self.defRPr = defRPr
        self.buClrTx = buClrTx
        self.buClr = buClr
        self.buSzTx = buSzTx
        self.buSzPct = buSzPct
        self.buSzPts = buSzPts
        self.buFontTx = buFontTx
        self.buFont = buFont
        self.buNone = buNone
        self.buAutoNum = buAutoNum
        self.buChar = buChar
        self.buBlip = buBlip
        self.defRPr = defRPr


class ListStyle(Serialisable):

    tagname = "lstStyle"
    namespace = DRAWING_NS

    defPPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl1pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl2pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl3pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl4pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl5pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl6pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl7pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl8pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    lvl9pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    __elements__ = ("defPPr", "lvl1pPr", "lvl2pPr", "lvl3pPr", "lvl4pPr",
                    "lvl5pPr", "lvl6pPr", "lvl7pPr", "lvl8pPr", "lvl9pPr")

    def __init__(self,
                 defPPr=None,
                 lvl1pPr=None,
                 lvl2pPr=None,
                 lvl3pPr=None,
                 lvl4pPr=None,
                 lvl5pPr=None,
                 lvl6pPr=None,
                 lvl7pPr=None,
                 lvl8pPr=None,
                 lvl9pPr=None,
                 extLst=None,
                ):
        self.defPPr = defPPr
        self.lvl1pPr = lvl1pPr
        self.lvl2pPr = lvl2pPr
        self.lvl3pPr = lvl3pPr
        self.lvl4pPr = lvl4pPr
        self.lvl5pPr = lvl5pPr
        self.lvl6pPr = lvl6pPr
        self.lvl7pPr = lvl7pPr
        self.lvl8pPr = lvl8pPr
        self.lvl9pPr = lvl9pPr


class RegularTextRun(Serialisable):

    tagname = "r"
    namespace = DRAWING_NS

    rPr = Typed(expected_type=CharacterProperties, allow_none=True)
    properties = Alias("rPr")
    t = NestedText(expected_type=str)
    value = Alias("t")

    __elements__ = ('rPr', 't')

    def __init__(self,
                 rPr=None,
                 t="",
                ):
        self.rPr = rPr
        self.t = t


class LineBreak(Serialisable):

    tagname = "br"
    namespace = DRAWING_NS

    rPr = Typed(expected_type=CharacterProperties, allow_none=True)

    __elements__ = ('rPr',)

    def __init__(self,
                 rPr=None,
                ):
        self.rPr = rPr


class TextField(Serialisable):

    id = String()
    type = String(allow_none=True)
    rPr = Typed(expected_type=CharacterProperties, allow_none=True)
    pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    t = String(allow_none=True)

    __elements__ = ('rPr', 'pPr')

    def __init__(self,
                 id=None,
                 type=None,
                 rPr=None,
                 pPr=None,
                 t=None,
                ):
        self.id = id
        self.type = type
        self.rPr = rPr
        self.pPr = pPr
        self.t = t


class Paragraph(Serialisable):

    tagname = "p"
    namespace = DRAWING_NS

    # uses element group EG_TextRun
    pPr = Typed(expected_type=ParagraphProperties, allow_none=True)
    properties = Alias("pPr")
    endParaRPr = Typed(expected_type=CharacterProperties, allow_none=True)
    r = Sequence(expected_type=RegularTextRun)
    text = Alias('r')
    br = Typed(expected_type=LineBreak, allow_none=True)
    fld = Typed(expected_type=TextField, allow_none=True)

    __elements__ = ('pPr', 'r', 'br', 'fld', 'endParaRPr')

    def __init__(self,
                 pPr=None,
                 endParaRPr=None,
                 r=None,
                 br=None,
                 fld=None,
                 ):
        self.pPr = pPr
        self.endParaRPr = endParaRPr
        if r is None:
            r = [RegularTextRun()]
        self.r = r
        self.br = br
        self.fld = fld


class GeomGuide(Serialisable):

    name = String(())
    fmla = String(())

    def __init__(self,
                 name=None,
                 fmla=None,
                ):
        self.name = name
        self.fmla = fmla


class GeomGuideList(Serialisable):

    gd = Sequence(expected_type=GeomGuide, allow_none=True)

    def __init__(self,
                 gd=None,
                ):
        self.gd = gd


class PresetTextShape(Serialisable):

    prst = Typed(expected_type=Set(values=(
        ['textNoShape', 'textPlain','textStop', 'textTriangle', 'textTriangleInverted', 'textChevron',
         'textChevronInverted', 'textRingInside', 'textRingOutside', 'textArchUp',
         'textArchDown', 'textCircle', 'textButton', 'textArchUpPour',
         'textArchDownPour', 'textCirclePour', 'textButtonPour', 'textCurveUp',
         'textCurveDown', 'textCanUp', 'textCanDown', 'textWave1', 'textWave2',
         'textDoubleWave1', 'textWave4', 'textInflate', 'textDeflate',
         'textInflateBottom', 'textDeflateBottom', 'textInflateTop',
         'textDeflateTop', 'textDeflateInflate', 'textDeflateInflateDeflate',
         'textFadeRight', 'textFadeLeft', 'textFadeUp', 'textFadeDown',
         'textSlantUp', 'textSlantDown', 'textCascadeUp', 'textCascadeDown'
         ]
    )))
    avLst = Typed(expected_type=GeomGuideList, allow_none=True)

    def __init__(self,
                 prst=None,
                 avLst=None,
                ):
        self.prst = prst
        self.avLst = avLst


class TextNormalAutofit(Serialisable):

    fontScale = Integer()
    lnSpcReduction = Integer()

    def __init__(self,
                 fontScale=None,
                 lnSpcReduction=None,
                ):
        self.fontScale = fontScale
        self.lnSpcReduction = lnSpcReduction


class RichTextProperties(Serialisable):

    tagname = "bodyPr"
    namespace = DRAWING_NS

    rot = Integer(allow_none=True)
    spcFirstLastPara = Bool(allow_none=True)
    vertOverflow = NoneSet(values=(['overflow', 'ellipsis', 'clip']))
    horzOverflow = NoneSet(values=(['overflow', 'clip']))
    vert = NoneSet(values=(['horz', 'vert', 'vert270', 'wordArtVert',
                            'eaVert', 'mongolianVert', 'wordArtVertRtl']))
    wrap = NoneSet(values=(['none', 'square']))
    lIns = Integer(allow_none=True)
    tIns = Integer(allow_none=True)
    rIns = Integer(allow_none=True)
    bIns = Integer(allow_none=True)
    numCol = Integer(allow_none=True)
    spcCol = Integer(allow_none=True)
    rtlCol = Bool(allow_none=True)
    fromWordArt = Bool(allow_none=True)
    anchor = NoneSet(values=(['t', 'ctr', 'b', 'just', 'dist']))
    anchorCtr = Bool(allow_none=True)
    forceAA = Bool(allow_none=True)
    upright = Bool(allow_none=True)
    compatLnSpc = Bool(allow_none=True)
    prstTxWarp = Typed(expected_type=PresetTextShape, allow_none=True)
    scene3d = Typed(expected_type=Scene3D, allow_none=True)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)
    noAutofit = EmptyTag()
    normAutofit = EmptyTag()
    spAutoFit = EmptyTag()
    flatTx = NestedInteger(attribute="z", allow_none=True)

    __elements__ = ('prstTxWarp', 'scene3d', 'noAutofit', 'normAutofit', 'spAutoFit')

    def __init__(self,
                 rot=None,
                 spcFirstLastPara=None,
                 vertOverflow=None,
                 horzOverflow=None,
                 vert=None,
                 wrap=None,
                 lIns=None,
                 tIns=None,
                 rIns=None,
                 bIns=None,
                 numCol=None,
                 spcCol=None,
                 rtlCol=None,
                 fromWordArt=None,
                 anchor=None,
                 anchorCtr=None,
                 forceAA=None,
                 upright=None,
                 compatLnSpc=None,
                 prstTxWarp=None,
                 scene3d=None,
                 extLst=None,
                 noAutofit=None,
                 normAutofit=None,
                 spAutoFit=None,
                 flatTx=None,
                ):
        self.rot = rot
        self.spcFirstLastPara = spcFirstLastPara
        self.vertOverflow = vertOverflow
        self.horzOverflow = horzOverflow
        self.vert = vert
        self.wrap = wrap
        self.lIns = lIns
        self.tIns = tIns
        self.rIns = rIns
        self.bIns = bIns
        self.numCol = numCol
        self.spcCol = spcCol
        self.rtlCol = rtlCol
        self.fromWordArt = fromWordArt
        self.anchor = anchor
        self.anchorCtr = anchorCtr
        self.forceAA = forceAA
        self.upright = upright
        self.compatLnSpc = compatLnSpc
        self.prstTxWarp = prstTxWarp
        self.scene3d = scene3d
        self.noAutofit = noAutofit
        self.normAutofit = normAutofit
        self.spAutoFit = spAutoFit
        self.flatTx = flatTx

""",
    },
    'openpyxl.drawing.xdr': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Spreadsheet Drawing has some copies of Drawing ML elements
"""

from .geometry import Point2D, PositiveSize2D, Transform2D


class XDRPoint2D(Point2D):

    namespace = None
    x = Point2D.x
    y = Point2D.y


class XDRPositiveSize2D(PositiveSize2D):

    namespace = None
    cx = PositiveSize2D.cx
    cy = PositiveSize2D.cy


class XDRTransform2D(Transform2D):

    namespace = None
    rot = Transform2D.rot
    flipH = Transform2D.flipH
    flipV = Transform2D.flipV
    off = Transform2D.off
    ext = Transform2D.ext
    chOff = Transform2D.chOff
    chExt = Transform2D.chExt

''',
    },
    'openpyxl.formatting': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .rule import Rule

""",
    },
    'openpyxl.formatting.formatting': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from collections import OrderedDict

from openpyxl.descriptors import (
    Bool,
    Sequence,
    Alias,
    Convertible,
)
from openpyxl.descriptors.serialisable import Serialisable

from .rule import Rule

from openpyxl.worksheet.cell_range import MultiCellRange

class ConditionalFormatting(Serialisable):

    tagname = "conditionalFormatting"

    sqref = Convertible(expected_type=MultiCellRange)
    cells = Alias("sqref")
    pivot = Bool(allow_none=True)
    cfRule = Sequence(expected_type=Rule)
    rules = Alias("cfRule")


    def __init__(self, sqref=(), pivot=None, cfRule=(), extLst=None):
        self.sqref = sqref
        self.pivot = pivot
        self.cfRule = cfRule


    def __eq__(self, other):
        if not isinstance(other, self.__class__):
            return False
        return self.sqref == other.sqref


    def __hash__(self):
        return hash(self.sqref)


    def __repr__(self):
        return "<{cls} {cells}>".format(cls=self.__class__.__name__, cells=self.sqref)


    def __contains__(self, coord):
        """
        Check whether a certain cell is affected by the formatting
        """
        return coord in self.sqref


class ConditionalFormattingList:
    """Conditional formatting rules."""


    def __init__(self):
        self._cf_rules = OrderedDict()
        self.max_priority = 0


    def add(self, range_string, cfRule):
        """Add a rule such as ColorScaleRule, FormulaRule or CellIsRule

         The priority will be added automatically.
        """
        cf = range_string
        if isinstance(range_string, str):
            cf = ConditionalFormatting(range_string)
        if not isinstance(cfRule, Rule):
            raise ValueError("Only instances of openpyxl.formatting.rule.Rule may be added")
        rule = cfRule
        self.max_priority += 1
        if not rule.priority:
            rule.priority = self.max_priority

        self._cf_rules.setdefault(cf, []).append(rule)


    def __bool__(self):
        return bool(self._cf_rules)


    def __len__(self):
        return len(self._cf_rules)


    def __iter__(self):
        for cf, rules in self._cf_rules.items():
            cf.rules = rules
            yield cf


    def __getitem__(self, key):
        """
        Get the rules for a cell range
        """
        if isinstance(key, str):
            key = ConditionalFormatting(sqref=key)
        return self._cf_rules[key]


    def __delitem__(self, key):
        key = ConditionalFormatting(sqref=key)
        del self._cf_rules[key]


    def __setitem__(self, key, rule):
        """
        Add a rule for a cell range
        """
        self.add(key, rule)

''',
    },
    'openpyxl.formatting.rule': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Sequence,
    Bool,
    NoneSet,
    Set,
    Integer,
    Float,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.styles.colors import Color, ColorDescriptor
from openpyxl.styles.differential import DifferentialStyle

from openpyxl.utils.cell import COORD_RE


class ValueDescriptor(Float):
    """
    Expected type depends upon type attribute of parent :-(

    Most values should be numeric BUT they can also be cell references
    """

    def __set__(self, instance, value):
        ref = None
        if value is not None and isinstance(value, str):
            ref = COORD_RE.match(value)
        if instance.type == "formula" or ref:
            self.expected_type = str
        else:
            self.expected_type = float
        super().__set__(instance, value)


class FormatObject(Serialisable):

    tagname = "cfvo"

    type = Set(values=(['num', 'percent', 'max', 'min', 'formula', 'percentile']))
    val = ValueDescriptor(allow_none=True)
    gte = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 type,
                 val=None,
                 gte=None,
                 extLst=None,
                ):
        self.type = type
        self.val = val
        self.gte = gte


class RuleType(Serialisable):

    cfvo = Sequence(expected_type=FormatObject)


class IconSet(RuleType):

    tagname = "iconSet"

    iconSet = NoneSet(values=(['3Arrows', '3ArrowsGray', '3Flags',
                           '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
                           '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
                           '5Arrows', '5ArrowsGray', '5Rating', '5Quarters']))
    showValue = Bool(allow_none=True)
    percent = Bool(allow_none=True)
    reverse = Bool(allow_none=True)

    __elements__ = ("cfvo",)

    def __init__(self,
                 iconSet=None,
                 showValue=None,
                 percent=None,
                 reverse=None,
                 cfvo=None,
                ):
        self.iconSet = iconSet
        self.showValue = showValue
        self.percent = percent
        self.reverse = reverse
        self.cfvo = cfvo


class DataBar(RuleType):

    tagname = "dataBar"

    minLength = Integer(allow_none=True)
    maxLength = Integer(allow_none=True)
    showValue = Bool(allow_none=True)
    color = ColorDescriptor()

    __elements__ = ('cfvo', 'color')

    def __init__(self,
                 minLength=None,
                 maxLength=None,
                 showValue=None,
                 cfvo=None,
                 color=None,
                ):
        self.minLength = minLength
        self.maxLength = maxLength
        self.showValue = showValue
        self.cfvo = cfvo
        self.color = color


class ColorScale(RuleType):

    tagname = "colorScale"

    color = Sequence(expected_type=Color)

    __elements__ = ('cfvo', 'color')

    def __init__(self,
                 cfvo=None,
                 color=None,
                ):
        self.cfvo = cfvo
        self.color = color


class Rule(Serialisable):

    tagname = "cfRule"

    type = Set(values=(['expression', 'cellIs', 'colorScale', 'dataBar',
                        'iconSet', 'top10', 'uniqueValues', 'duplicateValues', 'containsText',
                        'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks',
                        'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod',
                        'aboveAverage']))
    dxfId = Integer(allow_none=True)
    priority = Integer()
    stopIfTrue = Bool(allow_none=True)
    aboveAverage = Bool(allow_none=True)
    percent = Bool(allow_none=True)
    bottom = Bool(allow_none=True)
    operator = NoneSet(values=(['lessThan', 'lessThanOrEqual', 'equal',
                            'notEqual', 'greaterThanOrEqual', 'greaterThan', 'between', 'notBetween',
                            'containsText', 'notContains', 'beginsWith', 'endsWith']))
    text = String(allow_none=True)
    timePeriod = NoneSet(values=(['today', 'yesterday', 'tomorrow', 'last7Days',
                              'thisMonth', 'lastMonth', 'nextMonth', 'thisWeek', 'lastWeek',
                              'nextWeek']))
    rank = Integer(allow_none=True)
    stdDev = Integer(allow_none=True)
    equalAverage = Bool(allow_none=True)
    formula = Sequence(expected_type=str)
    colorScale = Typed(expected_type=ColorScale, allow_none=True)
    dataBar = Typed(expected_type=DataBar, allow_none=True)
    iconSet = Typed(expected_type=IconSet, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    dxf = Typed(expected_type=DifferentialStyle, allow_none=True)

    __elements__ = ('colorScale', 'dataBar', 'iconSet', 'formula')
    __attrs__ = ('type', 'rank', 'priority', 'equalAverage', 'operator',
                 'aboveAverage', 'dxfId', 'stdDev', 'stopIfTrue', 'timePeriod', 'text',
                 'percent', 'bottom')


    def __init__(self,
                 type,
                 dxfId=None,
                 priority=0,
                 stopIfTrue=None,
                 aboveAverage=None,
                 percent=None,
                 bottom=None,
                 operator=None,
                 text=None,
                 timePeriod=None,
                 rank=None,
                 stdDev=None,
                 equalAverage=None,
                 formula=(),
                 colorScale=None,
                 dataBar=None,
                 iconSet=None,
                 extLst=None,
                 dxf=None,
                ):
        self.type = type
        self.dxfId = dxfId
        self.priority = priority
        self.stopIfTrue = stopIfTrue
        self.aboveAverage = aboveAverage
        self.percent = percent
        self.bottom = bottom
        self.operator = operator
        self.text = text
        self.timePeriod = timePeriod
        self.rank = rank
        self.stdDev = stdDev
        self.equalAverage = equalAverage
        self.formula = formula
        self.colorScale = colorScale
        self.dataBar = dataBar
        self.iconSet = iconSet
        self.dxf = dxf


def ColorScaleRule(start_type=None,
                 start_value=None,
                 start_color=None,
                 mid_type=None,
                 mid_value=None,
                 mid_color=None,
                 end_type=None,
                 end_value=None,
                 end_color=None):

    """Backwards compatibility"""
    formats = []
    if start_type is not None:
        formats.append(FormatObject(type=start_type, val=start_value))
    if mid_type is not None:
        formats.append(FormatObject(type=mid_type, val=mid_value))
    if end_type is not None:
        formats.append(FormatObject(type=end_type, val=end_value))
    colors = []
    for v in (start_color, mid_color, end_color):
        if v is not None:
            if not isinstance(v, Color):
                v = Color(v)
            colors.append(v)
    cs = ColorScale(cfvo=formats, color=colors)
    rule = Rule(type="colorScale", colorScale=cs)
    return rule


def FormulaRule(formula=None, stopIfTrue=None, font=None, border=None,
                fill=None):
    """
    Conditional formatting with custom differential style
    """
    rule = Rule(type="expression", formula=formula, stopIfTrue=stopIfTrue)
    rule.dxf =  DifferentialStyle(font=font, border=border, fill=fill)
    return rule


def CellIsRule(operator=None, formula=None, stopIfTrue=None, font=None, border=None, fill=None):
    """
    Conditional formatting rule based on cell contents.
    """
    # Excel doesn't use >, >=, etc, but allow for ease of python development
    expand = {">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan", "<=": "lessThanOrEqual",
              "=": "equal", "==": "equal", "!=": "notEqual"}

    operator = expand.get(operator, operator)

    rule = Rule(type='cellIs', operator=operator, formula=formula, stopIfTrue=stopIfTrue)
    rule.dxf = DifferentialStyle(font=font, border=border, fill=fill)

    return rule


def IconSetRule(icon_style=None, type=None, values=None, showValue=None, percent=None, reverse=None):
    """
    Convenience function for creating icon set rules
    """
    cfvo = []
    for val in values:
        cfvo.append(FormatObject(type, val))
    icon_set = IconSet(iconSet=icon_style, cfvo=cfvo, showValue=showValue,
                       percent=percent, reverse=reverse)
    rule = Rule(type='iconSet', iconSet=icon_set)

    return rule


def DataBarRule(start_type=None, start_value=None, end_type=None,
                end_value=None, color=None, showValue=None, minLength=None, maxLength=None):
    start = FormatObject(start_type, start_value)
    end = FormatObject(end_type, end_value)
    data_bar = DataBar(cfvo=[start, end], color=color, showValue=showValue,
                       minLength=minLength, maxLength=maxLength)
    rule = Rule(type='dataBar', dataBar=data_bar)

    return rule

''',
    },
    'openpyxl.formula': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .tokenizer import Tokenizer

""",
    },
    'openpyxl.formula.tokenizer': {
        'is_package': False,
        'source': r'''
"""
This module contains a tokenizer for Excel formulae.

The tokenizer is based on the Javascript tokenizer found at
http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html written by Eric
Bachtal
"""

import re


class TokenizerError(Exception):
    """Base class for all Tokenizer errors."""


class Tokenizer:

    """
    A tokenizer for Excel worksheet formulae.

    Converts a str string representing an Excel formula (in A1 notation)
    into a sequence of `Token` objects.

    `formula`: The str string to tokenize

    Tokenizer defines a method `._parse()` to parse the formula into tokens,
    which can then be accessed through the `.items` attribute.

    """

    SN_RE = re.compile("^[1-9](\\.[0-9]+)?[Ee]$")  # Scientific notation
    WSPACE_RE = re.compile(r"[ \n]+")
    STRING_REGEXES = {
        # Inside a string, all characters are treated as literals, except for
        # the quote character used to start the string. That character, when
        # doubled is treated as a single character in the string. If an
        # unmatched quote appears, the string is terminated.
        '"': re.compile('"(?:[^"]*"")*[^"]*"(?!")'),
        "'": re.compile("'(?:[^']*'')*[^']*'(?!')"),
    }
    ERROR_CODES = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?",
                   "#NUM!", "#N/A", "#GETTING_DATA")
    TOKEN_ENDERS = ',;}) +-*/^&=><%'  # Each of these characters, marks the
                                       # end of an operand token

    def __init__(self, formula):
        self.formula = formula
        self.items = []
        self.token_stack = []  # Used to keep track of arrays, functions, and
                               # parentheses
        self.offset = 0  # How many chars have we read
        self.token = []  # Used to build up token values char by char
        self._parse()

    def _parse(self):
        """Populate self.items with the tokens from the formula."""
        if self.offset:
            return  # Already parsed!
        if not self.formula:
            return
        elif self.formula[0] == '=':
            self.offset += 1
        else:
            self.items.append(Token(self.formula, Token.LITERAL))
            return
        consumers = (
            ('"\'', self._parse_string),
            ('[', self._parse_brackets),
            ('#', self._parse_error),
            (' ', self._parse_whitespace),
            ('\n', self._parse_whitespace),
            ('+-*/^&=><%', self._parse_operator),
            ('{(', self._parse_opener),
            (')}', self._parse_closer),
            (';,', self._parse_separator),
        )
        dispatcher = {}  # maps chars to the specific parsing function
        for chars, consumer in consumers:
            dispatcher.update(dict.fromkeys(chars, consumer))
        while self.offset < len(self.formula):
            if self.check_scientific_notation():  # May consume one character
                continue
            curr_char = self.formula[self.offset]
            if curr_char in self.TOKEN_ENDERS:
                self.save_token()
            if curr_char in dispatcher:
                self.offset += dispatcher[curr_char]()
            else:
                # TODO: this can probably be sped up using a regex to get to
                # the next interesting character
                self.token.append(curr_char)
                self.offset += 1
        self.save_token()

    def _parse_string(self):
        """
        Parse a "-delimited string or '-delimited link.

        The offset must be pointing to either a single quote ("'") or double
        quote ('"') character. The strings are parsed according to Excel
        rules where to escape the delimiter you just double it up. E.g.,
        "abc""def" in Excel is parsed as 'abc"def' in Python.

        Returns the number of characters matched. (Does not update
        self.offset)

        """
        self.assert_empty_token(can_follow=':')
        delim = self.formula[self.offset]
        assert delim in ('"', "'")
        regex = self.STRING_REGEXES[delim]
        match = regex.match(self.formula[self.offset:])
        if match is None:
            subtype = "string" if delim == '"' else 'link'
            raise TokenizerError(f"Reached end of formula while parsing {subtype} in {self.formula}")
        match = match.group(0)
        if delim == '"':
            self.items.append(Token.make_operand(match))
        else:
            self.token.append(match)
        return len(match)

    def _parse_brackets(self):
        """
        Consume all the text between square brackets [].

        Returns the number of characters matched. (Does not update
        self.offset)

        """
        assert self.formula[self.offset] == '['
        lefts = [(t.start(), 1) for t in
                 re.finditer(r"\[", self.formula[self.offset:])]
        rights = [(t.start(), -1) for t in
                  re.finditer(r"\]", self.formula[self.offset:])]

        open_count = 0
        for idx, open_close in sorted(lefts + rights):
            open_count += open_close
            if open_count == 0:
                outer_right = idx + 1
                self.token.append(
                    self.formula[self.offset:self.offset + outer_right])
                return outer_right

        raise TokenizerError(f"Encountered unmatched '[' in {self.formula}")

    def _parse_error(self):
        """
        Consume the text following a '#' as an error.

        Looks for a match in self.ERROR_CODES and returns the number of
        characters matched. (Does not update self.offset)

        """
        self.assert_empty_token(can_follow='!')
        assert self.formula[self.offset] == '#'
        subformula = self.formula[self.offset:]
        for err in self.ERROR_CODES:
            if subformula.startswith(err):
                self.items.append(Token.make_operand(''.join(self.token) + err))
                del self.token[:]
                return len(err)
        raise TokenizerError(f"Invalid error code at position {self.offset} in '{self.formula}'")

    def _parse_whitespace(self):
        """
        Consume a string of consecutive spaces.

        Returns the number of spaces found. (Does not update self.offset).

        """
        assert self.formula[self.offset] in (' ', '\n')
        self.items.append(Token(self.formula[self.offset], Token.WSPACE))
        return self.WSPACE_RE.match(self.formula[self.offset:]).end()

    def _parse_operator(self):
        """
        Consume the characters constituting an operator.

        Returns the number of characters consumed. (Does not update
        self.offset)

        """
        if self.formula[self.offset:self.offset + 2] in ('>=', '<=', '<>'):
            self.items.append(Token(
                self.formula[self.offset:self.offset + 2],
                Token.OP_IN
            ))
            return 2
        curr_char = self.formula[self.offset]  # guaranteed to be 1 char
        assert curr_char in '%*/^&=><+-'
        if curr_char == '%':
            token = Token('%', Token.OP_POST)
        elif curr_char in "*/^&=><":
            token = Token(curr_char, Token.OP_IN)
        # From here on, curr_char is guaranteed to be in '+-'
        elif not self.items:
            token = Token(curr_char, Token.OP_PRE)
        else:
            prev = next((i for i in reversed(self.items)
                         if i.type != Token.WSPACE), None)
            is_infix = prev and (
                prev.subtype == Token.CLOSE
                or prev.type == Token.OP_POST
                or prev.type == Token.OPERAND
            )
            if is_infix:
                token = Token(curr_char, Token.OP_IN)
            else:
                token = Token(curr_char, Token.OP_PRE)
        self.items.append(token)
        return 1

    def _parse_opener(self):
        """
        Consumes a ( or { character.

        Returns the number of characters consumed. (Does not update
        self.offset)

        """
        assert self.formula[self.offset] in ('(', '{')
        if self.formula[self.offset] == '{':
            self.assert_empty_token()
            token = Token.make_subexp("{")
        elif self.token:
            token_value = "".join(self.token) + '('
            del self.token[:]
            token = Token.make_subexp(token_value)
        else:
            token = Token.make_subexp("(")
        self.items.append(token)
        self.token_stack.append(token)
        return 1

    def _parse_closer(self):
        """
        Consumes a } or ) character.

        Returns the number of characters consumed. (Does not update
        self.offset)

        """
        assert self.formula[self.offset] in (')', '}')
        token = self.token_stack.pop().get_closer()
        if token.value != self.formula[self.offset]:
            raise TokenizerError(
                "Mismatched ( and { pair in '%s'" % self.formula)
        self.items.append(token)
        return 1

    def _parse_separator(self):
        """
        Consumes a ; or , character.

        Returns the number of characters consumed. (Does not update
        self.offset)

        """
        curr_char = self.formula[self.offset]
        assert curr_char in (';', ',')
        if curr_char == ';':
            token = Token.make_separator(";")
        else:
            try:
                top_type = self.token_stack[-1].type
            except IndexError:
                token = Token(",", Token.OP_IN)  # Range Union operator
            else:
                if top_type == Token.PAREN:
                    token = Token(",", Token.OP_IN)  # Range Union operator
                else:
                    token = Token.make_separator(",")
        self.items.append(token)
        return 1

    def check_scientific_notation(self):
        """
        Consumes a + or - character if part of a number in sci. notation.

        Returns True if the character was consumed and self.offset was
        updated, False otherwise.

        """
        curr_char = self.formula[self.offset]
        if (curr_char in '+-'
                and len(self.token) >= 1
                and self.SN_RE.match("".join(self.token))):
            self.token.append(curr_char)
            self.offset += 1
            return True
        return False

    def assert_empty_token(self, can_follow=()):
        """
        Ensure that there's no token currently being parsed.

        Or if there is a token being parsed, it must end with a character in
        can_follow.

        If there are unconsumed token contents, it means we hit an unexpected
        token transition. In this case, we raise a TokenizerError

        """
        if self.token and self.token[-1] not in can_follow:
            raise TokenizerError(f"Unexpected character at position {self.offset} in '{self.formula}'")

    def save_token(self):
        """If there's a token being parsed, add it to the item list."""
        if self.token:
            self.items.append(Token.make_operand("".join(self.token)))
            del self.token[:]

    def render(self):
        """Convert the parsed tokens back to a string."""
        if not self.items:
            return ""
        elif self.items[0].type == Token.LITERAL:
            return self.items[0].value
        return "=" + "".join(token.value for token in self.items)


class Token:

    """
    A token in an Excel formula.

    Tokens have three attributes:

    * `value`: The string value parsed that led to this token
    * `type`: A string identifying the type of token
    * `subtype`: A string identifying subtype of the token (optional, and
                 defaults to "")

    """

    __slots__ = ['value', 'type', 'subtype']

    LITERAL = "LITERAL"
    OPERAND = "OPERAND"
    FUNC = "FUNC"
    ARRAY = "ARRAY"
    PAREN = "PAREN"
    SEP = "SEP"
    OP_PRE = "OPERATOR-PREFIX"
    OP_IN = "OPERATOR-INFIX"
    OP_POST = "OPERATOR-POSTFIX"
    WSPACE = "WHITE-SPACE"

    def __init__(self, value, type_, subtype=""):
        self.value = value
        self.type = type_
        self.subtype = subtype

    # Literal operands:
    #
    # Literal operands are always of type 'OPERAND' and can be of subtype
    # 'TEXT' (for text strings), 'NUMBER' (for all numeric types), 'LOGICAL'
    # (for TRUE and FALSE), 'ERROR' (for literal error values), or 'RANGE'
    # (for all range references).

    TEXT = 'TEXT'
    NUMBER = 'NUMBER'
    LOGICAL = 'LOGICAL'
    ERROR = 'ERROR'
    RANGE = 'RANGE'

    def __repr__(self):
        return u"{0} {1} {2}:".format(self.type, self.subtype, self.value)

    @classmethod
    def make_operand(cls, value):
        """Create an operand token."""
        if value.startswith('"'):
            subtype = cls.TEXT
        elif value.startswith('#'):
            subtype = cls.ERROR
        elif value in ('TRUE', 'FALSE'):
            subtype = cls.LOGICAL
        else:
            try:
                float(value)
                subtype = cls.NUMBER
            except ValueError:
                subtype = cls.RANGE
        return cls(value, cls.OPERAND, subtype)


    # Subexpresssions
    #
    # There are 3 types of `Subexpressions`: functions, array literals, and
    # parentheticals. Subexpressions have 'OPEN' and 'CLOSE' tokens. 'OPEN'
    # is used when parsing the initial expression token (i.e., '(' or '{')
    # and 'CLOSE' is used when parsing the closing expression token ('}' or
    # ')').

    OPEN = "OPEN"
    CLOSE = "CLOSE"

    @classmethod
    def make_subexp(cls, value, func=False):
        """
        Create a subexpression token.

        `value`: The value of the token
        `func`: If True, force the token to be of type FUNC

        """
        assert value[-1] in ('{', '}', '(', ')')
        if func:
            assert re.match('.+\\(|\\)', value)
            type_ = Token.FUNC
        elif value in '{}':
            type_ = Token.ARRAY
        elif value in '()':
            type_ = Token.PAREN
        else:
            type_ = Token.FUNC
        subtype = cls.CLOSE if value in ')}' else cls.OPEN
        return cls(value, type_, subtype)

    def get_closer(self):
        """Return a closing token that matches this token's type."""
        assert self.type in (self.FUNC, self.ARRAY, self.PAREN)
        assert self.subtype == self.OPEN
        value = "}" if self.type == self.ARRAY else ")"
        return self.make_subexp(value, func=self.type == self.FUNC)

    # Separator tokens
    #
    # Argument separators always have type 'SEP' and can have one of two
    # subtypes: 'ARG', 'ROW'. 'ARG' is used for the ',' token, when used to
    # delimit either function arguments or array elements. 'ROW' is used for
    # the ';' token, which is always used to delimit rows in an array
    # literal.

    ARG = "ARG"
    ROW = "ROW"

    @classmethod
    def make_separator(cls, value):
        """Create a separator token"""
        assert value in (',', ';')
        subtype = cls.ARG if value == ',' else cls.ROW
        return cls(value, cls.SEP, subtype)

''',
    },
    'openpyxl.formula.translate': {
        'is_package': False,
        'source': r'''
"""
This module contains code to translate formulae across cells in a worksheet.

The idea is that if A1 has formula "=B1+C1", then translating it to cell A2
results in formula "=B2+C2". The algorithm relies on the formula tokenizer
to identify the parts of the formula that need to change.

"""

import re
from .tokenizer import Tokenizer, Token
from openpyxl.utils import (
    coordinate_to_tuple,
    column_index_from_string,
    get_column_letter
)

class TranslatorError(Exception):
    """
    Raised when a formula can't be translated across cells.

    This error arises when a formula's references would be translated outside
    the worksheet's bounds on the top or left. Excel represents these
    situations with a #REF! literal error. E.g., if the formula at B2 is
    '=A1', attempting to translate the formula to B1 raises TranslatorError,
    since there's no cell above A1. Similarly, translating the same formula
    from B2 to A2 raises TranslatorError, since there's no cell to the left of
    A1.

    """


class Translator:

    """
    Modifies a formula so that it can be translated from one cell to another.

    `formula`: The str string to translate. Must include the leading '='
               character.
    `origin`: The cell address (in A1 notation) where this formula was
              defined (excluding the worksheet name).

    """

    def __init__(self, formula, origin):
        # Excel errors out when a workbook has formulae in R1C1 notation,
        # regardless of the calcPr:refMode setting, so I'm assuming the
        # formulae stored in the workbook must be in A1 notation.
        self.row, self.col = coordinate_to_tuple(origin)
        self.tokenizer = Tokenizer(formula)

    def get_tokens(self):
        "Returns a list with the tokens comprising the formula."
        return self.tokenizer.items

    ROW_RANGE_RE = re.compile(r"(\$?[1-9][0-9]{0,6}):(\$?[1-9][0-9]{0,6})$")
    COL_RANGE_RE = re.compile(r"(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})$")
    CELL_REF_RE = re.compile(r"(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

    @staticmethod
    def translate_row(row_str, rdelta):
        """
        Translate a range row-snippet by the given number of rows.
        """
        if row_str.startswith('$'):
            return row_str
        else:
            new_row = int(row_str) + rdelta
            if new_row <= 0:
                raise TranslatorError("Formula out of range")
            return str(new_row)

    @staticmethod
    def translate_col(col_str, cdelta):
        """
        Translate a range col-snippet by the given number of columns
        """
        if col_str.startswith('$'):
            return col_str
        else:
            try:
                return get_column_letter(
                    column_index_from_string(col_str) + cdelta)
            except ValueError:
                raise TranslatorError("Formula out of range")

    @staticmethod
    def strip_ws_name(range_str):
        "Splits out the worksheet reference, if any, from a range reference."
        # This code assumes that named ranges cannot contain any exclamation
        # marks. Excel refuses to create these (even using VBA), and
        # complains of a corrupt workbook when there are names with
        # exclamation marks. The ECMA spec only states that named ranges will
        # be of `ST_Xstring` type, which in theory allows '!' (char code
        # 0x21) per http://www.w3.org/TR/xml/#charsets
        if '!' in range_str:
            sheet, range_str = range_str.rsplit('!', 1)
            return sheet + "!", range_str
        return "", range_str

    @classmethod
    def translate_range(cls, range_str, rdelta, cdelta):
        """
        Translate an A1-style range reference to the destination cell.

        `rdelta`: the row offset to add to the range
        `cdelta`: the column offset to add to the range
        `range_str`: an A1-style reference to a range. Potentially includes
                     the worksheet reference. Could also be a named range.

        """
        ws_part, range_str = cls.strip_ws_name(range_str)
        match = cls.ROW_RANGE_RE.match(range_str)  # e.g. `3:4`
        if match is not None:
            return (ws_part + cls.translate_row(match.group(1), rdelta) + ":"
                    + cls.translate_row(match.group(2), rdelta))
        match = cls.COL_RANGE_RE.match(range_str)  # e.g. `A:BC`
        if match is not None:
            return (ws_part + cls.translate_col(match.group(1), cdelta) + ':'
                    + cls.translate_col(match.group(2), cdelta))
        if ':' in range_str: # e.g. `A1:B5`
            # The check is necessarily general because range references can
            # have one or both endpoints specified by named ranges. I.e.,
            # `named_range:C2`, `C2:named_range`, and `name1:name2` are all
            # valid references. Further, Excel allows chaining multiple
            # colons together (with unclear meaning)
            return ws_part + ":".join(
                cls.translate_range(piece, rdelta, cdelta)
                for piece in range_str.split(':'))
        match = cls.CELL_REF_RE.match(range_str)
        if match is None:  # Must be a named range
            return range_str
        return (ws_part + cls.translate_col(match.group(1), cdelta)
                + cls.translate_row(match.group(2), rdelta))

    def translate_formula(self, dest=None, row_delta=0, col_delta=0):
        """
        Convert the formula into A1 notation, or as row and column coordinates

        The formula is converted into A1 assuming it is assigned to the cell
        whose address is `dest` (no worksheet name).

        """
        tokens = self.get_tokens()
        if not tokens:
            return ""
        elif tokens[0].type == Token.LITERAL:
            return tokens[0].value
        out = ['=']
        # per the spec:
        # A compliant producer or consumer considers a defined name in the
        # range A1-XFD1048576 to be an error. All other names outside this
        # range can be defined as names and overrides a cell reference if an
        # ambiguity exists. (I.18.2.5)
        if dest:
            row, col = coordinate_to_tuple(dest)
            row_delta = row - self.row
            col_delta = col - self.col
        for token in tokens:
            if (token.type == Token.OPERAND
                and token.subtype == Token.RANGE):
                out.append(self.translate_range(token.value, row_delta,
                                                col_delta))
            else:
                out.append(token.value)
        return "".join(out)

''',
    },
    'openpyxl.packaging': {
        'is_package': True,
        'source': r'''
"""
Stuff related to Office OpenXML packaging: relationships, archive, content types.
"""

''',
    },
    'openpyxl.packaging.core': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import datetime

from openpyxl.descriptors import (
    DateTime,
    Alias,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.nested import NestedText
from openpyxl.xml.functions import (
    Element,
    QName,
)
from openpyxl.xml.constants import (
    COREPROPS_NS,
    DCORE_NS,
    XSI_NS,
    DCTERMS_NS,
)


class NestedDateTime(DateTime, NestedText):

    expected_type = datetime.datetime

    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if namespace is not None:
            tagname = "{%s}%s" % (namespace, tagname)
        el = Element(tagname)
        if value is not None:
            value = value.replace(tzinfo=None)
            el.text = value.isoformat(timespec="seconds") + 'Z'
            return el


class QualifiedDateTime(NestedDateTime):

    """In certain situations Excel will complain if the additional type
    attribute isn't set"""

    def to_tree(self, tagname=None, value=None, namespace=None):
        el = super().to_tree(tagname, value, namespace)
        el.set("{%s}type" % XSI_NS, QName(DCTERMS_NS, "W3CDTF"))
        return el


class DocumentProperties(Serialisable):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    tagname = "coreProperties"
    namespace = COREPROPS_NS

    category = NestedText(expected_type=str, allow_none=True)
    contentStatus = NestedText(expected_type=str, allow_none=True)
    keywords = NestedText(expected_type=str, allow_none=True)
    lastModifiedBy = NestedText(expected_type=str, allow_none=True)
    lastPrinted = NestedDateTime(allow_none=True)
    revision = NestedText(expected_type=str, allow_none=True)
    version = NestedText(expected_type=str, allow_none=True)
    last_modified_by = Alias("lastModifiedBy")

    # Dublin Core Properties
    subject = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    title = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    creator = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    description = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    identifier = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    language = NestedText(expected_type=str, allow_none=True, namespace=DCORE_NS)
    # Dublin Core Terms
    created = QualifiedDateTime(allow_none=True, namespace=DCTERMS_NS) # assumed to be UTC
    modified = QualifiedDateTime(allow_none=True, namespace=DCTERMS_NS) # assumed to be UTC

    __elements__ = ("creator", "title", "description", "subject","identifier",
                    "language", "created", "modified", "lastModifiedBy", "category",
                    "contentStatus", "version", "revision", "keywords", "lastPrinted",
                    )


    def __init__(self,
                 category=None,
                 contentStatus=None,
                 keywords=None,
                 lastModifiedBy=None,
                 lastPrinted=None,
                 revision=None,
                 version=None,
                 created=None,
                 creator="openpyxl",
                 description=None,
                 identifier=None,
                 language=None,
                 modified=None,
                 subject=None,
                 title=None,
                 ):
        now = datetime.datetime.now(tz=datetime.timezone.utc).replace(tzinfo=None)
        self.contentStatus = contentStatus
        self.lastPrinted = lastPrinted
        self.revision = revision
        self.version = version
        self.creator = creator
        self.lastModifiedBy = lastModifiedBy
        self.modified = modified or now
        self.created = created or now
        self.title = title
        self.subject = subject
        self.description = description
        self.identifier = identifier
        self.language = language
        self.keywords = keywords
        self.category = category

''',
    },
    'openpyxl.packaging.custom': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Implementation of custom properties see § 22.3 in the specification"""


from warnings import warn

from openpyxl.descriptors import Strict
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors import (
    Alias,
    String,
    Integer,
    Float,
    DateTime,
    Bool,
)
from openpyxl.descriptors.nested import (
    NestedText,
)

from openpyxl.xml.constants import (
    CUSTPROPS_NS,
    VTYPES_NS,
    CPROPS_FMTID,
)

from .core import NestedDateTime


class NestedBoolText(Bool, NestedText):
    """
    Descriptor for handling nested elements with the value stored in the text part
    """

    pass


class _CustomDocumentProperty(Serialisable):

    """
    Low-level representation of a Custom Document Property.
    Not used directly
    Must always contain a child element, even if this is empty
    """

    tagname = "property"
    _typ = None

    name = String(allow_none=True)
    lpwstr = NestedText(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    i4 = NestedText(expected_type=int, allow_none=True, namespace=VTYPES_NS)
    r8 = NestedText(expected_type=float, allow_none=True, namespace=VTYPES_NS)
    filetime = NestedDateTime(allow_none=True, namespace=VTYPES_NS)
    bool = NestedBoolText(expected_type=bool, allow_none=True, namespace=VTYPES_NS)
    linkTarget = String(expected_type=str, allow_none=True)
    fmtid = String()
    pid = Integer()

    def __init__(self,
                 name=None,
                 pid=0,
                 fmtid=CPROPS_FMTID,
                 linkTarget=None,
                 **kw):
        self.fmtid = fmtid
        self.pid = pid
        self.name = name
        self._typ = None
        self.linkTarget = linkTarget

        for k, v in kw.items():
            setattr(self, k, v)
            setattr(self, "_typ", k) # ugh!
        for e in self.__elements__:
            if e not in kw:
                setattr(self, e, None)


    @property
    def type(self):
        if self._typ is not None:
            return self._typ
        for a in self.__elements__:
            if getattr(self, a) is not None:
                return a
        if self.linkTarget is not None:
            return "linkTarget"


    def to_tree(self, tagname=None, idx=None, namespace=None):
        child = getattr(self, self._typ, None)
        if child is None:
            setattr(self, self._typ, "")

        return super().to_tree(tagname=None, idx=None, namespace=None)


class _CustomDocumentPropertyList(Serialisable):

    """
    Parses and seriliases property lists but is not used directly
    """

    tagname = "Properties"

    property = Sequence(expected_type=_CustomDocumentProperty, namespace=CUSTPROPS_NS)
    customProps = Alias("property")


    def __init__(self, property=()):
        self.property = property


    def __len__(self):
        return len(self.property)


    def to_tree(self, tagname=None, idx=None, namespace=None):
        for idx, p in enumerate(self.property, 2):
            p.pid = idx
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", CUSTPROPS_NS)

        return tree


class _TypedProperty(Strict):

    name = String()

    def __init__(self,
                 name,
                 value):
        self.name = name
        self.value = value


    def __eq__(self, other):
        return self.name == other.name and self.value == other.value


    def __repr__(self):
        return f"{self.__class__.__name__}, name={self.name}, value={self.value}"


class IntProperty(_TypedProperty):

    value = Integer()


class FloatProperty(_TypedProperty):

    value = Float()


class StringProperty(_TypedProperty):

    value = String(allow_none=True)


class DateTimeProperty(_TypedProperty):

    value = DateTime()


class BoolProperty(_TypedProperty):

    value = Bool()


class LinkProperty(_TypedProperty):

    value = String()


# from Python
CLASS_MAPPING = {
    StringProperty: "lpwstr",
    IntProperty: "i4",
    FloatProperty: "r8",
    DateTimeProperty: "filetime",
    BoolProperty: "bool",
    LinkProperty: "linkTarget"
}

XML_MAPPING = {v:k for k,v in CLASS_MAPPING.items()}


class CustomPropertyList(Strict):


    props = Sequence(expected_type=_TypedProperty)

    def __init__(self):
        self.props = []


    @classmethod
    def from_tree(cls, tree):
        """
        Create list from OOXML element
        """
        prop_list = _CustomDocumentPropertyList.from_tree(tree)
        props = []

        for prop in prop_list.property:
            attr = prop.type

            typ = XML_MAPPING.get(attr, None)
            if not typ:
                warn(f"Unknown type for {prop.name}")
                continue
            value = getattr(prop, attr)
            link = prop.linkTarget
            if link is not None:
                typ = LinkProperty
                value = prop.linkTarget

            new_prop = typ(name=prop.name, value=value)
            props.append(new_prop)

        new_prop_list = cls()
        new_prop_list.props = props
        return new_prop_list


    def append(self, prop):
        if prop.name in self.names:
            raise ValueError(f"Property with name {prop.name} already exists")

        self.props.append(prop)


    def to_tree(self):
        props = []

        for p in self.props:
            attr = CLASS_MAPPING.get(p.__class__, None)
            if not attr:
                raise TypeError("Unknown adapter for {p}")
            np = _CustomDocumentProperty(name=p.name, **{attr:p.value})
            if isinstance(p, LinkProperty):
                np._typ = "lpwstr"
                #np.lpwstr = ""
            props.append(np)

        prop_list = _CustomDocumentPropertyList(property=props)
        return prop_list.to_tree()


    def __len__(self):
        return len(self.props)


    @property
    def names(self):
        """List of property names"""
        return [p.name for p in self.props]


    def __getitem__(self, name):
        """
        Get property by name
        """
        for p in self.props:
            if p.name == name:
                return p
        raise KeyError(f"Property with name {name} not found")


    def __delitem__(self, name):
        """
        Delete a propery by name
        """
        for idx, p in enumerate(self.props):
            if p.name == name:
                self.props.pop(idx)
                return
        raise KeyError(f"Property with name {name} not found")


    def __repr__(self):
        return f"{self.__class__.__name__} containing {self.props}"


    def __iter__(self):
        return iter(self.props)

''',
    },
    'openpyxl.packaging.extended': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
)
from openpyxl.descriptors.nested import (
    NestedText,
)

from openpyxl.xml.constants import XPROPS_NS
from openpyxl import __version__


class DigSigBlob(Serialisable):

    __elements__ = __attrs__ = ()


class VectorLpstr(Serialisable):

    __elements__ = __attrs__ = ()


class VectorVariant(Serialisable):

    __elements__ = __attrs__ = ()


class ExtendedProperties(Serialisable):

    """
    See 22.2

    Most of this is irrelevant but Excel is very picky about the version number

    It uses XX.YYYY (Version.Build) and expects everyone else to

    We provide Major.Minor and the full version in the application name
    """

    tagname = "Properties"

    Template = NestedText(expected_type=str, allow_none=True)
    Manager = NestedText(expected_type=str, allow_none=True)
    Company = NestedText(expected_type=str, allow_none=True)
    Pages = NestedText(expected_type=int, allow_none=True)
    Words = NestedText(expected_type=int,allow_none=True)
    Characters = NestedText(expected_type=int, allow_none=True)
    PresentationFormat = NestedText(expected_type=str, allow_none=True)
    Lines = NestedText(expected_type=int, allow_none=True)
    Paragraphs = NestedText(expected_type=int, allow_none=True)
    Slides = NestedText(expected_type=int, allow_none=True)
    Notes = NestedText(expected_type=int, allow_none=True)
    TotalTime = NestedText(expected_type=int, allow_none=True)
    HiddenSlides = NestedText(expected_type=int, allow_none=True)
    MMClips = NestedText(expected_type=int, allow_none=True)
    ScaleCrop = NestedText(expected_type=bool, allow_none=True)
    HeadingPairs = Typed(expected_type=VectorVariant, allow_none=True)
    TitlesOfParts = Typed(expected_type=VectorLpstr, allow_none=True)
    LinksUpToDate = NestedText(expected_type=bool, allow_none=True)
    CharactersWithSpaces = NestedText(expected_type=int, allow_none=True)
    SharedDoc = NestedText(expected_type=bool, allow_none=True)
    HyperlinkBase = NestedText(expected_type=str, allow_none=True)
    HLinks = Typed(expected_type=VectorVariant, allow_none=True)
    HyperlinksChanged = NestedText(expected_type=bool, allow_none=True)
    DigSig = Typed(expected_type=DigSigBlob, allow_none=True)
    Application = NestedText(expected_type=str, allow_none=True)
    AppVersion = NestedText(expected_type=str, allow_none=True)
    DocSecurity = NestedText(expected_type=int, allow_none=True)

    __elements__ = ('Application', 'AppVersion', 'DocSecurity', 'ScaleCrop',
                    'LinksUpToDate', 'SharedDoc', 'HyperlinksChanged')

    def __init__(self,
                 Template=None,
                 Manager=None,
                 Company=None,
                 Pages=None,
                 Words=None,
                 Characters=None,
                 PresentationFormat=None,
                 Lines=None,
                 Paragraphs=None,
                 Slides=None,
                 Notes=None,
                 TotalTime=None,
                 HiddenSlides=None,
                 MMClips=None,
                 ScaleCrop=None,
                 HeadingPairs=None,
                 TitlesOfParts=None,
                 LinksUpToDate=None,
                 CharactersWithSpaces=None,
                 SharedDoc=None,
                 HyperlinkBase=None,
                 HLinks=None,
                 HyperlinksChanged=None,
                 DigSig=None,
                 Application=None,
                 AppVersion=None,
                 DocSecurity=None,
                ):
        self.Template = Template
        self.Manager = Manager
        self.Company = Company
        self.Pages = Pages
        self.Words = Words
        self.Characters = Characters
        self.PresentationFormat = PresentationFormat
        self.Lines = Lines
        self.Paragraphs = Paragraphs
        self.Slides = Slides
        self.Notes = Notes
        self.TotalTime = TotalTime
        self.HiddenSlides = HiddenSlides
        self.MMClips = MMClips
        self.ScaleCrop = ScaleCrop
        self.HeadingPairs = None
        self.TitlesOfParts = None
        self.LinksUpToDate = LinksUpToDate
        self.CharactersWithSpaces = CharactersWithSpaces
        self.SharedDoc = SharedDoc
        self.HyperlinkBase = HyperlinkBase
        self.HLinks = None
        self.HyperlinksChanged = HyperlinksChanged
        self.DigSig = None
        self.Application = f"Microsoft Excel Compatible / Openpyxl {__version__}"
        self.AppVersion = ".".join(__version__.split(".")[:-1])
        self.DocSecurity = DocSecurity


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", XPROPS_NS)
        return tree

''',
    },
    'openpyxl.packaging.interface': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from abc import abstractproperty
from openpyxl.compat.abc import ABC


class ISerialisableFile(ABC):

    """
    Interface for Serialisable classes that represent files in the archive
    """


    @abstractproperty
    def id(self):
        """
        Object id making it unique
        """
        pass


    @abstractproperty
    def _path(self):
        """
        File path in the archive
        """
        pass


    @abstractproperty
    def _namespace(self):
        """
        Qualified namespace when serialised
        """
        pass


    @abstractproperty
    def _type(self):
        """
        The content type for the manifest
        """


    @abstractproperty
    def _rel_type(self):
        """
        The content type for relationships
        """


    @abstractproperty
    def _rel_id(self):
        """
        Links object with parent
        """

''',
    },
    'openpyxl.packaging.manifest': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
File manifest
"""
from mimetypes import MimeTypes
import os.path

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Sequence
from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_THEME,
    ARC_STYLE,
    THEME_TYPE,
    STYLES_TYPE,
    CONTYPES_NS,
    ACTIVEX,
    CTRL,
    VBA,
)
from openpyxl.xml.functions import tostring

# initialise mime-types
mimetypes = MimeTypes()
mimetypes.add_type('application/xml', ".xml")
mimetypes.add_type('application/vnd.openxmlformats-package.relationships+xml', ".rels")
mimetypes.add_type("application/vnd.ms-office.vbaProject", ".bin")
mimetypes.add_type("application/vnd.openxmlformats-officedocument.vmlDrawing", ".vml")
mimetypes.add_type("image/x-emf", ".emf")


class FileExtension(Serialisable):

    tagname = "Default"

    Extension = String()
    ContentType = String()

    def __init__(self, Extension, ContentType):
        self.Extension = Extension
        self.ContentType = ContentType


class Override(Serialisable):

    tagname = "Override"

    PartName = String()
    ContentType = String()

    def __init__(self, PartName, ContentType):
        self.PartName = PartName
        self.ContentType = ContentType


DEFAULT_TYPES = [
    FileExtension("rels", "application/vnd.openxmlformats-package.relationships+xml"),
    FileExtension("xml", "application/xml"),
]

DEFAULT_OVERRIDE = [
    Override("/" + ARC_STYLE, STYLES_TYPE), # Styles
    Override("/" + ARC_THEME, THEME_TYPE), # Theme
    Override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    Override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
]


class Manifest(Serialisable):

    tagname = "Types"

    Default = Sequence(expected_type=FileExtension, unique=True)
    Override = Sequence(expected_type=Override, unique=True)
    path = "[Content_Types].xml"

    __elements__ = ("Default", "Override")

    def __init__(self,
                 Default=(),
                 Override=(),
                 ):
        if not Default:
            Default = DEFAULT_TYPES
        self.Default = Default
        if not Override:
            Override = DEFAULT_OVERRIDE
        self.Override = Override


    @property
    def filenames(self):
        return [part.PartName for part in self.Override]


    @property
    def extensions(self):
        """
        Map content types to file extensions
        Skip parts without extensions
        """
        exts = {os.path.splitext(part.PartName)[-1] for part in self.Override}
        return [(ext[1:], mimetypes.types_map[True][ext]) for ext in sorted(exts) if ext]


    def to_tree(self):
        """
        Custom serialisation method to allow setting a default namespace
        """
        defaults = [t.Extension for t in self.Default]
        for ext, mime in self.extensions:
            if ext not in defaults:
                mime = FileExtension(ext, mime)
                self.Default.append(mime)
        tree = super().to_tree()
        tree.set("xmlns", CONTYPES_NS)
        return tree


    def __contains__(self, content_type):
        """
        Check whether a particular content type is contained
        """
        for t in self.Override:
            if t.ContentType == content_type:
                return True


    def find(self, content_type):
        """
        Find specific content-type
        """
        try:
            return next(self.findall(content_type))
        except StopIteration:
            return


    def findall(self, content_type):
        """
        Find all elements of a specific content-type
        """
        for t in self.Override:
            if t.ContentType == content_type:
                yield t


    def append(self, obj):
        """
        Add content object to the package manifest
        # needs a contract...
        """
        ct = Override(PartName=obj.path, ContentType=obj.mime_type)
        self.Override.append(ct)


    def _write(self, archive, workbook):
        """
        Write manifest to the archive
        """
        self.append(workbook)
        self._write_vba(workbook)
        self._register_mimetypes(filenames=archive.namelist())
        archive.writestr(self.path, tostring(self.to_tree()))


    def _register_mimetypes(self, filenames):
        """
        Make sure that the mime type for all file extensions is registered
        """
        for fn in filenames:
            ext = os.path.splitext(fn)[-1]
            if not ext:
                continue
            mime = mimetypes.types_map[True][ext]
            fe = FileExtension(ext[1:], mime)
            self.Default.append(fe)


    def _write_vba(self, workbook):
        """
        Add content types from cached workbook when keeping VBA
        """
        if workbook.vba_archive:
            node = fromstring(workbook.vba_archive.read(ARC_CONTENT_TYPES))
            mf = Manifest.from_tree(node)
            filenames = self.filenames
            for override in mf.Override:
                if override.PartName not in (ACTIVEX, CTRL, VBA):
                    continue
                if override.PartName not in filenames:
                    self.Override.append(override)

''',
    },
    'openpyxl.packaging.relationship': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import posixpath
from warnings import warn

from openpyxl.descriptors import (
    String,
    Alias,
    Sequence,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.container import ElementList

from openpyxl.xml.constants import REL_NS, PKG_REL_NS
from openpyxl.xml.functions import (
    Element,
    fromstring,
)


class Relationship(Serialisable):
    """Represents many kinds of relationships."""

    tagname = "Relationship"

    Type = String()
    Target = String()
    target = Alias("Target")
    TargetMode = String(allow_none=True)
    Id = String(allow_none=True)
    id = Alias("Id")


    def __init__(self,
                 Id=None,
                 Type=None,
                 type=None,
                 Target=None,
                 TargetMode=None
                 ):
        """
        `type` can be used as a shorthand with the default relationships namespace
        otherwise the `Type` must be a fully qualified URL
        """
        if type is not None:
            Type = "{0}/{1}".format(REL_NS, type)
        self.Type = Type
        self.Target = Target
        self.TargetMode = TargetMode
        self.Id = Id


class RelationshipList(ElementList):

    tagname = "Relationships"
    expected_type = Relationship


    def append(self, value):
        super().append(value)
        if not value.Id:
            value.Id = f"rId{len(self)}"


    def find(self, content_type):
        """
        Find relationships by content-type
        NB. these content-types namespaced objects and different to the MIME-types
        in the package manifest :-(
        """
        for r in self:
            if r.Type == content_type:
                yield r


    def get(self, key):
        for r in self:
            if r.Id == key:
                return r
        raise KeyError("Unknown relationship: {0}".format(key))


    def to_dict(self):
        """Return a dictionary of relations keyed by id"""
        return {r.id:r for r in self}


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", PKG_REL_NS)
        return tree


def get_rels_path(path):
    """
    Convert relative path to absolutes that can be loaded from a zip
    archive.
    The path to be passed in is that of containing object (workbook,
    worksheet, etc.)
    """
    folder, obj = posixpath.split(path)
    filename = posixpath.join(folder, '_rels', '{0}.rels'.format(obj))
    return filename


def get_dependents(archive, filename):
    """
    Normalise dependency file paths to absolute ones

    Relative paths are relative to parent object
    """
    src = archive.read(filename)
    node = fromstring(src)
    try:
        rels = RelationshipList.from_tree(node)
    except TypeError:
        msg = "{0} contains invalid dependency definitions".format(filename)
        warn(msg)
        rels = RelationshipList()
    folder = posixpath.dirname(filename)
    parent = posixpath.split(folder)[0]
    for r in rels:
        if r.TargetMode == "External":
            continue
        elif r.target.startswith("/"):
            r.target = r.target[1:]
        else:
            pth = posixpath.join(parent, r.target)
            r.target = posixpath.normpath(pth)
    return rels


def get_rel(archive, deps, id=None, cls=None):
    """
    Get related object based on id or rel_type
    """
    if not any([id, cls]):
        raise ValueError("Either the id or the content type are required")
    if id is not None:
        rel = deps.get(id)
    else:
        try:
            rel = next(deps.find(cls.rel_type))
        except StopIteration: # no known dependency
            return

    path = rel.target
    src = archive.read(path)
    tree = fromstring(src)
    obj = cls.from_tree(tree)

    rels_path = get_rels_path(path)
    try:
        obj.deps = get_dependents(archive, rels_path)
    except KeyError:
        obj.deps = []

    return obj

''',
    },
    'openpyxl.packaging.workbook': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    String,
    Integer,
    Bool,
    NoneSet,
)
from openpyxl.descriptors.excel import ExtensionList, Relation
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.descriptors.nested import NestedString

from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.workbook.defined_name import DefinedNameList
from openpyxl.workbook.external_reference import ExternalReference
from openpyxl.workbook.function_group import FunctionGroupList
from openpyxl.workbook.properties import WorkbookProperties, CalcProperties, FileVersion
from openpyxl.workbook.protection import WorkbookProtection, FileSharing
from openpyxl.workbook.smart_tags import SmartTagList, SmartTagProperties
from openpyxl.workbook.views import CustomWorkbookView, BookView
from openpyxl.workbook.web import WebPublishing, WebPublishObjectList


class FileRecoveryProperties(Serialisable):

    tagname = "fileRecoveryPr"

    autoRecover = Bool(allow_none=True)
    crashSave = Bool(allow_none=True)
    dataExtractLoad = Bool(allow_none=True)
    repairLoad = Bool(allow_none=True)

    def __init__(self,
                 autoRecover=None,
                 crashSave=None,
                 dataExtractLoad=None,
                 repairLoad=None,
                ):
        self.autoRecover = autoRecover
        self.crashSave = crashSave
        self.dataExtractLoad = dataExtractLoad
        self.repairLoad = repairLoad


class ChildSheet(Serialisable):
    """
    Represents a reference to a worksheet or chartsheet in workbook.xml

    It contains the title, order and state but only an indirect reference to
    the objects themselves.
    """

    tagname = "sheet"

    name = String()
    sheetId = Integer()
    state = NoneSet(values=(['visible', 'hidden', 'veryHidden']))
    id = Relation()

    def __init__(self,
                 name=None,
                 sheetId=None,
                 state="visible",
                 id=None,
                ):
        self.name = name
        self.sheetId = sheetId
        self.state = state
        self.id = id


class PivotCache(Serialisable):

    tagname = "pivotCache"

    cacheId = Integer()
    id = Relation()

    def __init__(self,
                 cacheId=None,
                 id=None
                ):
        self.cacheId = cacheId
        self.id = id


class WorkbookPackage(Serialisable):

    """
    Represent the workbook file in the archive
    """

    tagname = "workbook"

    conformance = NoneSet(values=['strict', 'transitional'])
    fileVersion = Typed(expected_type=FileVersion, allow_none=True)
    fileSharing = Typed(expected_type=FileSharing, allow_none=True)
    workbookPr = Typed(expected_type=WorkbookProperties, allow_none=True)
    properties = Alias("workbookPr")
    workbookProtection = Typed(expected_type=WorkbookProtection, allow_none=True)
    bookViews = NestedSequence(expected_type=BookView)
    sheets = NestedSequence(expected_type=ChildSheet)
    functionGroups = Typed(expected_type=FunctionGroupList, allow_none=True)
    externalReferences = NestedSequence(expected_type=ExternalReference)
    definedNames = Typed(expected_type=DefinedNameList, allow_none=True)
    calcPr = Typed(expected_type=CalcProperties, allow_none=True)
    oleSize = NestedString(allow_none=True, attribute="ref")
    customWorkbookViews = NestedSequence(expected_type=CustomWorkbookView)
    pivotCaches = NestedSequence(expected_type=PivotCache, allow_none=True)
    smartTagPr = Typed(expected_type=SmartTagProperties, allow_none=True)
    smartTagTypes = Typed(expected_type=SmartTagList, allow_none=True)
    webPublishing = Typed(expected_type=WebPublishing, allow_none=True)
    fileRecoveryPr = Typed(expected_type=FileRecoveryProperties, allow_none=True)
    webPublishObjects = Typed(expected_type=WebPublishObjectList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    Ignorable = NestedString(namespace="http://schemas.openxmlformats.org/markup-compatibility/2006", allow_none=True)

    __elements__ = ('fileVersion', 'fileSharing', 'workbookPr',
                    'workbookProtection', 'bookViews', 'sheets', 'functionGroups',
                    'externalReferences', 'definedNames', 'calcPr', 'oleSize',
                    'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes',
                    'webPublishing', 'fileRecoveryPr', 'webPublishObjects')

    def __init__(self,
                 conformance=None,
                 fileVersion=None,
                 fileSharing=None,
                 workbookPr=None,
                 workbookProtection=None,
                 bookViews=(),
                 sheets=(),
                 functionGroups=None,
                 externalReferences=(),
                 definedNames=None,
                 calcPr=None,
                 oleSize=None,
                 customWorkbookViews=(),
                 pivotCaches=(),
                 smartTagPr=None,
                 smartTagTypes=None,
                 webPublishing=None,
                 fileRecoveryPr=None,
                 webPublishObjects=None,
                 extLst=None,
                 Ignorable=None,
                ):
        self.conformance = conformance
        self.fileVersion = fileVersion
        self.fileSharing = fileSharing
        if workbookPr is None:
            workbookPr = WorkbookProperties()
        self.workbookPr = workbookPr
        self.workbookProtection = workbookProtection
        self.bookViews = bookViews
        self.sheets = sheets
        self.functionGroups = functionGroups
        self.externalReferences = externalReferences
        self.definedNames = definedNames
        self.calcPr = calcPr
        self.oleSize = oleSize
        self.customWorkbookViews = customWorkbookViews
        self.pivotCaches = pivotCaches
        self.smartTagPr = smartTagPr
        self.smartTagTypes = smartTagTypes
        self.webPublishing = webPublishing
        self.fileRecoveryPr = fileRecoveryPr
        self.webPublishObjects = webPublishObjects


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def active(self):
        for view in self.bookViews:
            if view.activeTab is not None:
                return view.activeTab
        return 0

''',
    },
    'openpyxl.pivot': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

""",
    },
    'openpyxl.pivot.cache': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Float,
    Set,
    NoneSet,
    String,
    Integer,
    DateTime,
    Sequence,
)

from openpyxl.descriptors.excel import (
    HexBinary,
    ExtensionList,
    Relation,
)
from openpyxl.descriptors.nested import NestedInteger
from openpyxl.descriptors.sequence import (
    NestedSequence,
    MultiSequence,
    MultiSequencePart,
)
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring
from openpyxl.packaging.relationship import (
    RelationshipList,
    Relationship,
    get_rels_path
)

from .table import (
    PivotArea,
    Reference,
)
from .fields import (
    Boolean,
    Error,
    Missing,
    Number,
    Text,
    TupleList,
    DateTimeField,
)

class MeasureDimensionMap(Serialisable):

    tagname = "map"

    measureGroup = Integer(allow_none=True)
    dimension = Integer(allow_none=True)

    def __init__(self,
                 measureGroup=None,
                 dimension=None,
                ):
        self.measureGroup = measureGroup
        self.dimension = dimension


class MeasureGroup(Serialisable):

    tagname = "measureGroup"

    name = String()
    caption = String()

    def __init__(self,
                 name=None,
                 caption=None,
                ):
        self.name = name
        self.caption = caption


class PivotDimension(Serialisable):

    tagname = "dimension"

    measure = Bool()
    name = String()
    uniqueName = String()
    caption = String()

    def __init__(self,
                 measure=None,
                 name=None,
                 uniqueName=None,
                 caption=None,
                ):
        self.measure = measure
        self.name = name
        self.uniqueName = uniqueName
        self.caption = caption


class CalculatedMember(Serialisable):

    tagname = "calculatedMember"

    name = String()
    mdx = String()
    memberName = String(allow_none=True)
    hierarchy = String(allow_none=True)
    parent = String(allow_none=True)
    solveOrder = Integer(allow_none=True)
    set = Bool()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 name=None,
                 mdx=None,
                 memberName=None,
                 hierarchy=None,
                 parent=None,
                 solveOrder=None,
                 set=None,
                 extLst=None,
                ):
        self.name = name
        self.mdx = mdx
        self.memberName = memberName
        self.hierarchy = hierarchy
        self.parent = parent
        self.solveOrder = solveOrder
        self.set = set
        #self.extLst = extLst


class CalculatedItem(Serialisable):

    tagname = "calculatedItem"

    field = Integer(allow_none=True)
    formula = String()
    pivotArea = Typed(expected_type=PivotArea, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotArea', 'extLst')

    def __init__(self,
                 field=None,
                 formula=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.field = field
        self.formula = formula
        self.pivotArea = pivotArea
        self.extLst = extLst


class ServerFormat(Serialisable):

    tagname = "serverFormat"

    culture = String(allow_none=True)
    format = String(allow_none=True)

    def __init__(self,
                 culture=None,
                 format=None,
                ):
        self.culture = culture
        self.format = format


class Query(Serialisable):

    tagname = "query"

    mdx = String()
    tpls = Typed(expected_type=TupleList, allow_none=True)

    __elements__ = ('tpls',)

    def __init__(self,
                 mdx=None,
                 tpls=None,
                ):
        self.mdx = mdx
        self.tpls = tpls


class OLAPSet(Serialisable):

    tagname = "set"

    count = Integer()
    maxRank = Integer()
    setDefinition = String()
    sortType = NoneSet(values=(['ascending', 'descending', 'ascendingAlpha',
                                'descendingAlpha', 'ascendingNatural', 'descendingNatural']))
    queryFailed = Bool()
    tpls = Typed(expected_type=TupleList, allow_none=True)
    sortByTuple = Typed(expected_type=TupleList, allow_none=True)

    __elements__ = ('tpls', 'sortByTuple')

    def __init__(self,
                 count=None,
                 maxRank=None,
                 setDefinition=None,
                 sortType=None,
                 queryFailed=None,
                 tpls=None,
                 sortByTuple=None,
                ):
        self.count = count
        self.maxRank = maxRank
        self.setDefinition = setDefinition
        self.sortType = sortType
        self.queryFailed = queryFailed
        self.tpls = tpls
        self.sortByTuple = sortByTuple


class PCDSDTCEntries(Serialisable):
    # Implements CT_PCDSDTCEntries

    tagname = "entries"

    count = Integer(allow_none=True)
    # elements are choice
    m = Typed(expected_type=Missing, allow_none=True)
    n = Typed(expected_type=Number, allow_none=True)
    e = Typed(expected_type=Error, allow_none=True)
    s = Typed(expected_type=Text, allow_none=True)

    __elements__ = ('m', 'n', 'e', 's')

    def __init__(self,
                 count=None,
                 m=None,
                 n=None,
                 e=None,
                 s=None,
                ):
        self.count = count
        self.m = m
        self.n = n
        self.e = e
        self.s = s


class TupleCache(Serialisable):

    tagname = "tupleCache"

    entries = Typed(expected_type=PCDSDTCEntries, allow_none=True)
    sets = NestedSequence(expected_type=OLAPSet, count=True)
    queryCache = NestedSequence(expected_type=Query, count=True)
    serverFormats = NestedSequence(expected_type=ServerFormat, count=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('entries', 'sets', 'queryCache', 'serverFormats', 'extLst')

    def __init__(self,
                 entries=None,
                 sets=(),
                 queryCache=(),
                 serverFormats=(),
                 extLst=None,
                ):
        self.entries = entries
        self.sets = sets
        self.queryCache = queryCache
        self.serverFormats = serverFormats
        self.extLst = extLst


class OLAPKPI(Serialisable):

    tagname = "kpi"

    uniqueName = String()
    caption = String(allow_none=True)
    displayFolder = String(allow_none=True)
    measureGroup = String(allow_none=True)
    parent = String(allow_none=True)
    value = String()
    goal = String(allow_none=True)
    status = String(allow_none=True)
    trend = String(allow_none=True)
    weight = String(allow_none=True)
    time = String(allow_none=True)

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 displayFolder=None,
                 measureGroup=None,
                 parent=None,
                 value=None,
                 goal=None,
                 status=None,
                 trend=None,
                 weight=None,
                 time=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.displayFolder = displayFolder
        self.measureGroup = measureGroup
        self.parent = parent
        self.value = value
        self.goal = goal
        self.status = status
        self.trend = trend
        self.weight = weight
        self.time = time


class GroupMember(Serialisable):

    tagname = "groupMember"

    uniqueName = String()
    group = Bool()

    def __init__(self,
                 uniqueName=None,
                 group=None,
                ):
        self.uniqueName = uniqueName
        self.group = group


class LevelGroup(Serialisable):

    tagname = "group"

    name = String()
    uniqueName = String()
    caption = String()
    uniqueParent = String()
    id = Integer()
    groupMembers = NestedSequence(expected_type=GroupMember, count=True)

    __elements__ = ('groupMembers',)

    def __init__(self,
                 name=None,
                 uniqueName=None,
                 caption=None,
                 uniqueParent=None,
                 id=None,
                 groupMembers=(),
                ):
        self.name = name
        self.uniqueName = uniqueName
        self.caption = caption
        self.uniqueParent = uniqueParent
        self.id = id
        self.groupMembers = groupMembers


class GroupLevel(Serialisable):

    tagname = "groupLevel"

    uniqueName = String()
    caption = String()
    user = Bool()
    customRollUp = Bool()
    groups = NestedSequence(expected_type=LevelGroup, count=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('groups', 'extLst')

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 user=None,
                 customRollUp=None,
                 groups=(),
                 extLst=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.user = user
        self.customRollUp = customRollUp
        self.groups = groups
        self.extLst = extLst


class FieldUsage(Serialisable):

    tagname = "fieldUsage"

    x = Integer()

    def __init__(self,
                 x=None,
                ):
        self.x = x


class CacheHierarchy(Serialisable):

    tagname = "cacheHierarchy"

    uniqueName = String()
    caption = String(allow_none=True)
    measure = Bool()
    set = Bool()
    parentSet = Integer(allow_none=True)
    iconSet = Integer()
    attribute = Bool()
    time = Bool()
    keyAttribute = Bool()
    defaultMemberUniqueName = String(allow_none=True)
    allUniqueName = String(allow_none=True)
    allCaption = String(allow_none=True)
    dimensionUniqueName = String(allow_none=True)
    displayFolder = String(allow_none=True)
    measureGroup = String(allow_none=True)
    measures = Bool()
    count = Integer()
    oneField = Bool()
    memberValueDatatype = Integer(allow_none=True)
    unbalanced = Bool(allow_none=True)
    unbalancedGroup = Bool(allow_none=True)
    hidden = Bool()
    fieldsUsage = NestedSequence(expected_type=FieldUsage, count=True)
    groupLevels = NestedSequence(expected_type=GroupLevel, count=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('fieldsUsage', 'groupLevels')

    def __init__(self,
                 uniqueName="",
                 caption=None,
                 measure=None,
                 set=None,
                 parentSet=None,
                 iconSet=0,
                 attribute=None,
                 time=None,
                 keyAttribute=None,
                 defaultMemberUniqueName=None,
                 allUniqueName=None,
                 allCaption=None,
                 dimensionUniqueName=None,
                 displayFolder=None,
                 measureGroup=None,
                 measures=None,
                 count=None,
                 oneField=None,
                 memberValueDatatype=None,
                 unbalanced=None,
                 unbalancedGroup=None,
                 hidden=None,
                 fieldsUsage=(),
                 groupLevels=(),
                 extLst=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.measure = measure
        self.set = set
        self.parentSet = parentSet
        self.iconSet = iconSet
        self.attribute = attribute
        self.time = time
        self.keyAttribute = keyAttribute
        self.defaultMemberUniqueName = defaultMemberUniqueName
        self.allUniqueName = allUniqueName
        self.allCaption = allCaption
        self.dimensionUniqueName = dimensionUniqueName
        self.displayFolder = displayFolder
        self.measureGroup = measureGroup
        self.measures = measures
        self.count = count
        self.oneField = oneField
        self.memberValueDatatype = memberValueDatatype
        self.unbalanced = unbalanced
        self.unbalancedGroup = unbalancedGroup
        self.hidden = hidden
        self.fieldsUsage = fieldsUsage
        self.groupLevels = groupLevels
        self.extLst = extLst


class GroupItems(Serialisable):

    tagname = "groupItems"

    m = Sequence(expected_type=Missing)
    n = Sequence(expected_type=Number)
    b = Sequence(expected_type=Boolean)
    e = Sequence(expected_type=Error)
    s = Sequence(expected_type=Text)
    d = Sequence(expected_type=DateTimeField,)

    __elements__ = ('m', 'n', 'b', 'e', 's', 'd')
    __attrs__ = ("count", )

    def __init__(self,
                 count=None,
                 m=(),
                 n=(),
                 b=(),
                 e=(),
                 s=(),
                 d=(),
                ):
        self.m = m
        self.n = n
        self.b = b
        self.e = e
        self.s = s
        self.d = d


    @property
    def count(self):
        return len(self.m + self.n + self.b + self.e + self.s + self.d)


class RangePr(Serialisable):

    tagname = "rangePr"

    autoStart = Bool(allow_none=True)
    autoEnd = Bool(allow_none=True)
    groupBy = NoneSet(values=(['range', 'seconds', 'minutes', 'hours', 'days',
                           'months', 'quarters', 'years']))
    startNum = Float(allow_none=True)
    endNum = Float(allow_none=True)
    startDate = DateTime(allow_none=True)
    endDate = DateTime(allow_none=True)
    groupInterval = Float(allow_none=True)

    def __init__(self,
                 autoStart=True,
                 autoEnd=True,
                 groupBy="range",
                 startNum=None,
                 endNum=None,
                 startDate=None,
                 endDate=None,
                 groupInterval=1,
                ):
        self.autoStart = autoStart
        self.autoEnd = autoEnd
        self.groupBy = groupBy
        self.startNum = startNum
        self.endNum = endNum
        self.startDate = startDate
        self.endDate = endDate
        self.groupInterval = groupInterval


class FieldGroup(Serialisable):

    tagname = "fieldGroup"

    par = Integer(allow_none=True)
    base = Integer(allow_none=True)
    rangePr = Typed(expected_type=RangePr, allow_none=True)
    discretePr = NestedSequence(expected_type=NestedInteger, count=True)
    groupItems = Typed(expected_type=GroupItems, allow_none=True)

    __elements__ = ('rangePr', 'discretePr', 'groupItems')

    def __init__(self,
                 par=None,
                 base=None,
                 rangePr=None,
                 discretePr=(),
                 groupItems=None,
                ):
        self.par = par
        self.base = base
        self.rangePr = rangePr
        self.discretePr = discretePr
        self.groupItems = groupItems


class SharedItems(Serialisable):

    tagname = "sharedItems"

    _fields = MultiSequence()
    m = MultiSequencePart(expected_type=Missing, store="_fields")
    n = MultiSequencePart(expected_type=Number, store="_fields")
    b = MultiSequencePart(expected_type=Boolean, store="_fields")
    e = MultiSequencePart(expected_type=Error, store="_fields")
    s = MultiSequencePart(expected_type=Text,  store="_fields")
    d = MultiSequencePart(expected_type=DateTimeField, store="_fields")
    # attributes are optional and must be derived from associated cache records
    containsSemiMixedTypes = Bool(allow_none=True)
    containsNonDate = Bool(allow_none=True)
    containsDate = Bool(allow_none=True)
    containsString = Bool(allow_none=True)
    containsBlank = Bool(allow_none=True)
    containsMixedTypes = Bool(allow_none=True)
    containsNumber = Bool(allow_none=True)
    containsInteger = Bool(allow_none=True)
    minValue = Float(allow_none=True)
    maxValue = Float(allow_none=True)
    minDate = DateTime(allow_none=True)
    maxDate = DateTime(allow_none=True)
    longText = Bool(allow_none=True)

    __attrs__ = ('count', 'containsBlank', 'containsDate', 'containsInteger',
                 'containsMixedTypes', 'containsNonDate', 'containsNumber',
                 'containsSemiMixedTypes', 'containsString', 'minValue', 'maxValue',
                 'minDate', 'maxDate', 'longText')

    def __init__(self,
                 _fields=(),
                 containsSemiMixedTypes=None,
                 containsNonDate=None,
                 containsDate=None,
                 containsString=None,
                 containsBlank=None,
                 containsMixedTypes=None,
                 containsNumber=None,
                 containsInteger=None,
                 minValue=None,
                 maxValue=None,
                 minDate=None,
                 maxDate=None,
                 count=None,
                 longText=None,
                ):
        self._fields = _fields
        self.containsBlank = containsBlank
        self.containsDate = containsDate
        self.containsNonDate = containsNonDate
        self.containsString = containsString
        self.containsMixedTypes = containsMixedTypes
        self.containsSemiMixedTypes = containsSemiMixedTypes
        self.containsNumber = containsNumber
        self.containsInteger = containsInteger
        self.minValue = minValue
        self.maxValue = maxValue
        self.minDate = minDate
        self.maxDate = maxDate
        self.longText = longText


    @property
    def count(self):
        return len(self._fields)


class CacheField(Serialisable):

    tagname = "cacheField"

    sharedItems = Typed(expected_type=SharedItems, allow_none=True)
    fieldGroup = Typed(expected_type=FieldGroup, allow_none=True)
    mpMap = NestedInteger(allow_none=True, attribute="v")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    name = String()
    caption = String(allow_none=True)
    propertyName = String(allow_none=True)
    serverField = Bool(allow_none=True)
    uniqueList = Bool(allow_none=True)
    numFmtId = Integer(allow_none=True)
    formula = String(allow_none=True)
    sqlType = Integer(allow_none=True)
    hierarchy = Integer(allow_none=True)
    level = Integer(allow_none=True)
    databaseField = Bool(allow_none=True)
    mappingCount = Integer(allow_none=True)
    memberPropertyField = Bool(allow_none=True)

    __elements__ = ('sharedItems', 'fieldGroup', 'mpMap')

    def __init__(self,
                 sharedItems=None,
                 fieldGroup=None,
                 mpMap=None,
                 extLst=None,
                 name=None,
                 caption=None,
                 propertyName=None,
                 serverField=None,
                 uniqueList=True,
                 numFmtId=None,
                 formula=None,
                 sqlType=0,
                 hierarchy=0,
                 level=0,
                 databaseField=True,
                 mappingCount=None,
                 memberPropertyField=None,
                ):
        self.sharedItems = sharedItems
        self.fieldGroup = fieldGroup
        self.mpMap = mpMap
        self.extLst = extLst
        self.name = name
        self.caption = caption
        self.propertyName = propertyName
        self.serverField = serverField
        self.uniqueList = uniqueList
        self.numFmtId = numFmtId
        self.formula = formula
        self.sqlType = sqlType
        self.hierarchy = hierarchy
        self.level = level
        self.databaseField = databaseField
        self.mappingCount = mappingCount
        self.memberPropertyField = memberPropertyField


class RangeSet(Serialisable):

    tagname = "rangeSet"

    i1 = Integer(allow_none=True)
    i2 = Integer(allow_none=True)
    i3 = Integer(allow_none=True)
    i4 = Integer(allow_none=True)
    ref = String()
    name = String(allow_none=True)
    sheet = String(allow_none=True)

    def __init__(self,
                 i1=None,
                 i2=None,
                 i3=None,
                 i4=None,
                 ref=None,
                 name=None,
                 sheet=None,
                ):
        self.i1 = i1
        self.i2 = i2
        self.i3 = i3
        self.i4 = i4
        self.ref = ref
        self.name = name
        self.sheet = sheet


class PageItem(Serialisable):

    tagname = "pageItem"

    name = String()

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Consolidation(Serialisable):

    tagname = "consolidation"

    autoPage = Bool(allow_none=True)
    pages = NestedSequence(expected_type=PageItem, count=True)
    rangeSets = NestedSequence(expected_type=RangeSet, count=True)

    __elements__ = ('pages', 'rangeSets')

    def __init__(self,
                 autoPage=None,
                 pages=(),
                 rangeSets=(),
                ):
        self.autoPage = autoPage
        self.pages = pages
        self.rangeSets = rangeSets


class WorksheetSource(Serialisable):

    tagname = "worksheetSource"

    ref = String(allow_none=True)
    name = String(allow_none=True)
    sheet = String(allow_none=True)

    def __init__(self,
                 ref=None,
                 name=None,
                 sheet=None,
                ):
        self.ref = ref
        self.name = name
        self.sheet = sheet


class CacheSource(Serialisable):

    tagname = "cacheSource"

    type = Set(values=(['worksheet', 'external', 'consolidation', 'scenario']))
    connectionId = Integer(allow_none=True)
    # some elements are choice
    worksheetSource = Typed(expected_type=WorksheetSource, allow_none=True)
    consolidation = Typed(expected_type=Consolidation, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('worksheetSource', 'consolidation',)

    def __init__(self,
                 type=None,
                 connectionId=None,
                 worksheetSource=None,
                 consolidation=None,
                 extLst=None,
                ):
        self.type = type
        self.connectionId = connectionId
        self.worksheetSource = worksheetSource
        self.consolidation = consolidation


class CacheDefinition(Serialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
    _id = 1
    _path = "/xl/pivotCache/pivotCacheDefinition{0}.xml"
    records = None

    tagname = "pivotCacheDefinition"

    invalid = Bool(allow_none=True)
    saveData = Bool(allow_none=True)
    refreshOnLoad = Bool(allow_none=True)
    optimizeMemory = Bool(allow_none=True)
    enableRefresh = Bool(allow_none=True)
    refreshedBy = String(allow_none=True)
    refreshedDate = Float(allow_none=True)
    refreshedDateIso = DateTime(allow_none=True)
    backgroundQuery = Bool(allow_none=True)
    missingItemsLimit = Integer(allow_none=True)
    createdVersion = Integer(allow_none=True)
    refreshedVersion = Integer(allow_none=True)
    minRefreshableVersion = Integer(allow_none=True)
    recordCount = Integer(allow_none=True)
    upgradeOnRefresh = Bool(allow_none=True)
    supportSubquery = Bool(allow_none=True)
    supportAdvancedDrill = Bool(allow_none=True)
    cacheSource = Typed(expected_type=CacheSource)
    cacheFields = NestedSequence(expected_type=CacheField, count=True)
    cacheHierarchies = NestedSequence(expected_type=CacheHierarchy, allow_none=True)
    kpis = NestedSequence(expected_type=OLAPKPI, count=True)
    tupleCache = Typed(expected_type=TupleCache, allow_none=True)
    calculatedItems = NestedSequence(expected_type=CalculatedItem, count=True)
    calculatedMembers = NestedSequence(expected_type=CalculatedMember, count=True)
    dimensions = NestedSequence(expected_type=PivotDimension, allow_none=True)
    measureGroups = NestedSequence(expected_type=MeasureGroup, count=True)
    maps = NestedSequence(expected_type=MeasureDimensionMap, count=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    id = Relation()

    __elements__ = ('cacheSource', 'cacheFields', 'cacheHierarchies', 'kpis',
                    'tupleCache', 'calculatedItems', 'calculatedMembers', 'dimensions',
                    'measureGroups', 'maps',)

    def __init__(self,
                 invalid=None,
                 saveData=None,
                 refreshOnLoad=None,
                 optimizeMemory=None,
                 enableRefresh=None,
                 refreshedBy=None,
                 refreshedDate=None,
                 refreshedDateIso=None,
                 backgroundQuery=None,
                 missingItemsLimit=None,
                 createdVersion=None,
                 refreshedVersion=None,
                 minRefreshableVersion=None,
                 recordCount=None,
                 upgradeOnRefresh=None,
                 tupleCache=None,
                 supportSubquery=None,
                 supportAdvancedDrill=None,
                 cacheSource=None,
                 cacheFields=(),
                 cacheHierarchies=(),
                 kpis=(),
                 calculatedItems=(),
                 calculatedMembers=(),
                 dimensions=(),
                 measureGroups=(),
                 maps=(),
                 extLst=None,
                 id = None,
                ):
        self.invalid = invalid
        self.saveData = saveData
        self.refreshOnLoad = refreshOnLoad
        self.optimizeMemory = optimizeMemory
        self.enableRefresh = enableRefresh
        self.refreshedBy = refreshedBy
        self.refreshedDate = refreshedDate
        self.refreshedDateIso = refreshedDateIso
        self.backgroundQuery = backgroundQuery
        self.missingItemsLimit = missingItemsLimit
        self.createdVersion = createdVersion
        self.refreshedVersion = refreshedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.recordCount = recordCount
        self.upgradeOnRefresh = upgradeOnRefresh
        self.supportSubquery = supportSubquery
        self.supportAdvancedDrill = supportAdvancedDrill
        self.cacheSource = cacheSource
        self.cacheFields = cacheFields
        self.cacheHierarchies = cacheHierarchies
        self.kpis = kpis
        self.tupleCache = tupleCache
        self.calculatedItems = calculatedItems
        self.calculatedMembers = calculatedMembers
        self.dimensions = dimensions
        self.measureGroups = measureGroups
        self.maps = maps
        self.id = id


    def to_tree(self):
        node = super().to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Add to zipfile and update manifest
        """
        self._write_rels(archive, manifest)
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        """
        Write the relevant child objects and add links
        """
        if self.records is None:
            return

        rels = RelationshipList()
        r = Relationship(Type=self.records.rel_type, Target=self.records.path)
        rels.append(r)
        self.id = r.id
        self.records._id = self._id
        self.records._write(archive, manifest)

        path = get_rels_path(self.path)
        xml = tostring(rels.to_tree())
        archive.writestr(path[1:], xml)

''',
    },
    'openpyxl.pivot.fields': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    DateTime,
    Bool,
    Float,
    String,
    Integer,
    Sequence,
)
from openpyxl.descriptors.excel import HexBinary

class Index(Serialisable):

    tagname = "x"

    v = Integer(allow_none=True)

    def __init__(self,
                 v=0,
                ):
        self.v = v


class Tuple(Serialisable):

    tagname = "tpl"

    fld = Integer(allow_none=True)
    hier = Integer(allow_none=True)
    item = Integer()

    def __init__(self,
                 fld=None,
                 hier=None,
                 item=None,
                ):
        self.fld = fld
        self.hier = hier
        self.item = item


class TupleList(Serialisable):

    tagname = "tpls"

    c = Integer(allow_none=True)
    tpl = Typed(expected_type=Tuple, )

    __elements__ = ('tpl',)

    def __init__(self,
                 c=None,
                 tpl=None,
                ):
        self.c = c
        self.tpl = tpl


class Missing(Serialisable):

    tagname = "m"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=Index)
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = x
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Number(Serialisable):

    tagname = "n"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=Index)
    v = Float()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Error(Serialisable):

    tagname = "e"

    tpls = Typed(expected_type=TupleList, allow_none=True)
    x = Sequence(expected_type=Index)
    v = String()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=None,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Boolean(Serialisable):

    tagname = "b"

    x = Sequence(expected_type=Index)
    v = Bool()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                ):
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp


class Text(Serialisable):

    tagname = "s"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=Index)
    v = String()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary(allow_none=True)
    fc = HexBinary(allow_none=True)
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                 ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class DateTimeField(Serialisable):

    tagname = "d"

    x = Sequence(expected_type=Index)
    v = DateTime()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 ):
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp

""",
    },
    'openpyxl.pivot.record': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    Sequence,
)
from openpyxl.descriptors.sequence import (
    MultiSequence,
    MultiSequencePart,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.nested import (
    NestedInteger,
    NestedBool,
)

from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring

from .fields import (
    Boolean,
    Error,
    Missing,
    Number,
    Text,
    TupleList,
    DateTimeField,
    Index,
)


class Record(Serialisable):

    tagname = "r"

    _fields = MultiSequence()
    m = MultiSequencePart(expected_type=Missing, store="_fields")
    n = MultiSequencePart(expected_type=Number, store="_fields")
    b = MultiSequencePart(expected_type=Boolean, store="_fields")
    e = MultiSequencePart(expected_type=Error, store="_fields")
    s = MultiSequencePart(expected_type=Text,  store="_fields")
    d = MultiSequencePart(expected_type=DateTimeField, store="_fields")
    x = MultiSequencePart(expected_type=Index, store="_fields")


    def __init__(self,
                 _fields=(),
                 m=None,
                 n=None,
                 b=None,
                 e=None,
                 s=None,
                 d=None,
                 x=None,
                ):
        self._fields = _fields


class RecordList(Serialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
    _id = 1
    _path = "/xl/pivotCache/pivotCacheRecords{0}.xml"

    tagname ="pivotCacheRecords"

    r = Sequence(expected_type=Record, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('r', )
    __attrs__ = ('count', )

    def __init__(self,
                 count=None,
                 r=(),
                 extLst=None,
                ):
        self.r = r
        self.extLst = extLst


    @property
    def count(self):
        return len(self.r)


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Write to zipfile and update manifest
        """
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        pass

''',
    },
    'openpyxl.pivot.table': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


from collections import defaultdict
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    NoneSet,
    Set,
    Bool,
    String,
    Bool,
    Sequence,
)

from openpyxl.descriptors.excel import ExtensionList, Relation
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import tostring
from openpyxl.packaging.relationship import (
    RelationshipList,
    Relationship,
    get_rels_path
)
from .fields import Index

from openpyxl.worksheet.filters import (
    AutoFilter,
)


class HierarchyUsage(Serialisable):

    tagname = "hierarchyUsage"

    hierarchyUsage = Integer()

    def __init__(self,
                 hierarchyUsage=None,
                ):
        self.hierarchyUsage = hierarchyUsage


class ColHierarchiesUsage(Serialisable):

    tagname = "colHierarchiesUsage"

    colHierarchyUsage = Sequence(expected_type=HierarchyUsage, )

    __elements__ = ('colHierarchyUsage',)
    __attrs__ = ('count', )

    def __init__(self,
                 count=None,
                 colHierarchyUsage=(),
                ):
        self.colHierarchyUsage = colHierarchyUsage


    @property
    def count(self):
        return len(self.colHierarchyUsage)


class RowHierarchiesUsage(Serialisable):

    tagname = "rowHierarchiesUsage"

    rowHierarchyUsage = Sequence(expected_type=HierarchyUsage, )

    __elements__ = ('rowHierarchyUsage',)
    __attrs__ = ('count', )

    def __init__(self,
                 count=None,
                 rowHierarchyUsage=(),
                ):
        self.rowHierarchyUsage = rowHierarchyUsage

    @property
    def count(self):
        return len(self.rowHierarchyUsage)


class PivotFilter(Serialisable):

    tagname = "filter"

    fld = Integer()
    mpFld = Integer(allow_none=True)
    type = Set(values=(['unknown', 'count', 'percent', 'sum', 'captionEqual',
                        'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith',
                        'captionEndsWith', 'captionNotEndsWith', 'captionContains',
                        'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual',
                        'captionLessThan', 'captionLessThanOrEqual', 'captionBetween',
                        'captionNotBetween', 'valueEqual', 'valueNotEqual', 'valueGreaterThan',
                        'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual',
                        'valueBetween', 'valueNotBetween', 'dateEqual', 'dateNotEqual',
                        'dateOlderThan', 'dateOlderThanOrEqual', 'dateNewerThan',
                        'dateNewerThanOrEqual', 'dateBetween', 'dateNotBetween', 'tomorrow',
                        'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth',
                        'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter',
                        'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4',
                        'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11',
                        'M12']))
    evalOrder = Integer(allow_none=True)
    id = Integer()
    iMeasureHier = Integer(allow_none=True)
    iMeasureFld = Integer(allow_none=True)
    name = String(allow_none=True)
    description = String(allow_none=True)
    stringValue1 = String(allow_none=True)
    stringValue2 = String(allow_none=True)
    autoFilter = Typed(expected_type=AutoFilter, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('autoFilter',)

    def __init__(self,
                 fld=None,
                 mpFld=None,
                 type=None,
                 evalOrder=None,
                 id=None,
                 iMeasureHier=None,
                 iMeasureFld=None,
                 name=None,
                 description=None,
                 stringValue1=None,
                 stringValue2=None,
                 autoFilter=None,
                 extLst=None,
                ):
        self.fld = fld
        self.mpFld = mpFld
        self.type = type
        self.evalOrder = evalOrder
        self.id = id
        self.iMeasureHier = iMeasureHier
        self.iMeasureFld = iMeasureFld
        self.name = name
        self.description = description
        self.stringValue1 = stringValue1
        self.stringValue2 = stringValue2
        self.autoFilter = autoFilter


class PivotFilters(Serialisable):

    count = Integer()
    filter = Typed(expected_type=PivotFilter, allow_none=True)

    __elements__ = ('filter',)

    def __init__(self,
                 count=None,
                 filter=None,
                ):
        self.filter = filter


class PivotTableStyle(Serialisable):

    tagname = "pivotTableStyleInfo"

    name = String(allow_none=True)
    showRowHeaders = Bool()
    showColHeaders = Bool()
    showRowStripes = Bool()
    showColStripes = Bool()
    showLastColumn = Bool()

    def __init__(self,
                 name=None,
                 showRowHeaders=None,
                 showColHeaders=None,
                 showRowStripes=None,
                 showColStripes=None,
                 showLastColumn=None,
                ):
        self.name = name
        self.showRowHeaders = showRowHeaders
        self.showColHeaders = showColHeaders
        self.showRowStripes = showRowStripes
        self.showColStripes = showColStripes
        self.showLastColumn = showLastColumn


class MemberList(Serialisable):

    tagname = "members"

    level = Integer(allow_none=True)
    member = NestedSequence(expected_type=String, attribute="name")

    __elements__ = ('member',)

    def __init__(self,
                 count=None,
                 level=None,
                 member=(),
                ):
        self.level = level
        self.member = member

    @property
    def count(self):
        return len(self.member)


class MemberProperty(Serialisable):

    tagname = "mps"

    name = String(allow_none=True)
    showCell = Bool(allow_none=True)
    showTip = Bool(allow_none=True)
    showAsCaption = Bool(allow_none=True)
    nameLen = Integer(allow_none=True)
    pPos = Integer(allow_none=True)
    pLen = Integer(allow_none=True)
    level = Integer(allow_none=True)
    field = Integer()

    def __init__(self,
                 name=None,
                 showCell=None,
                 showTip=None,
                 showAsCaption=None,
                 nameLen=None,
                 pPos=None,
                 pLen=None,
                 level=None,
                 field=None,
                ):
        self.name = name
        self.showCell = showCell
        self.showTip = showTip
        self.showAsCaption = showAsCaption
        self.nameLen = nameLen
        self.pPos = pPos
        self.pLen = pLen
        self.level = level
        self.field = field


class PivotHierarchy(Serialisable):

    tagname = "pivotHierarchy"

    outline = Bool()
    multipleItemSelectionAllowed = Bool()
    subtotalTop = Bool()
    showInFieldList = Bool()
    dragToRow = Bool()
    dragToCol = Bool()
    dragToPage = Bool()
    dragToData = Bool()
    dragOff = Bool()
    includeNewItemsInFilter = Bool()
    caption = String(allow_none=True)
    mps = NestedSequence(expected_type=MemberProperty, count=True)
    members = Typed(expected_type=MemberList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('mps', 'members',)

    def __init__(self,
                 outline=None,
                 multipleItemSelectionAllowed=None,
                 subtotalTop=None,
                 showInFieldList=None,
                 dragToRow=None,
                 dragToCol=None,
                 dragToPage=None,
                 dragToData=None,
                 dragOff=None,
                 includeNewItemsInFilter=None,
                 caption=None,
                 mps=(),
                 members=None,
                 extLst=None,
                ):
        self.outline = outline
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.subtotalTop = subtotalTop
        self.showInFieldList = showInFieldList
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.caption = caption
        self.mps = mps
        self.members = members
        self.extLst = extLst


class Reference(Serialisable):

    tagname = "reference"

    field = Integer(allow_none=True)
    selected = Bool(allow_none=True)
    byPosition = Bool(allow_none=True)
    relative = Bool(allow_none=True)
    defaultSubtotal = Bool(allow_none=True)
    sumSubtotal = Bool(allow_none=True)
    countASubtotal = Bool(allow_none=True)
    avgSubtotal = Bool(allow_none=True)
    maxSubtotal = Bool(allow_none=True)
    minSubtotal = Bool(allow_none=True)
    productSubtotal = Bool(allow_none=True)
    countSubtotal = Bool(allow_none=True)
    stdDevSubtotal = Bool(allow_none=True)
    stdDevPSubtotal = Bool(allow_none=True)
    varSubtotal = Bool(allow_none=True)
    varPSubtotal = Bool(allow_none=True)
    x = Sequence(expected_type=Index)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 field=None,
                 count=None,
                 selected=None,
                 byPosition=None,
                 relative=None,
                 defaultSubtotal=None,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 x=(),
                 extLst=None,
                ):
        self.field = field
        self.selected = selected
        self.byPosition = byPosition
        self.relative = relative
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.x = x


    @property
    def count(self):
        return len(self.field)


class PivotArea(Serialisable):

    tagname = "pivotArea"

    references = NestedSequence(expected_type=Reference, count=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    field = Integer(allow_none=True)
    type = NoneSet(values=(['normal', 'data', 'all', 'origin', 'button',
                            'topEnd', 'topRight']))
    dataOnly = Bool(allow_none=True)
    labelOnly = Bool(allow_none=True)
    grandRow = Bool(allow_none=True)
    grandCol = Bool(allow_none=True)
    cacheIndex = Bool(allow_none=True)
    outline = Bool(allow_none=True)
    offset = String(allow_none=True)
    collapsedLevelsAreSubtotals = Bool(allow_none=True)
    axis = NoneSet(values=(['axisRow', 'axisCol', 'axisPage', 'axisValues']))
    fieldPosition = Integer(allow_none=True)

    __elements__ = ('references',)

    def __init__(self,
                 references=(),
                 extLst=None,
                 field=None,
                 type="normal",
                 dataOnly=True,
                 labelOnly=None,
                 grandRow=None,
                 grandCol=None,
                 cacheIndex=None,
                 outline=True,
                 offset=None,
                 collapsedLevelsAreSubtotals=None,
                 axis=None,
                 fieldPosition=None,
                ):
        self.references = references
        self.extLst = extLst
        self.field = field
        self.type = type
        self.dataOnly = dataOnly
        self.labelOnly = labelOnly
        self.grandRow = grandRow
        self.grandCol = grandCol
        self.cacheIndex = cacheIndex
        self.outline = outline
        self.offset = offset
        self.collapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals
        self.axis = axis
        self.fieldPosition = fieldPosition


class ChartFormat(Serialisable):

    tagname = "chartFormat"

    chart = Integer()
    format = Integer()
    series = Bool()
    pivotArea = Typed(expected_type=PivotArea, )

    __elements__ = ('pivotArea',)

    def __init__(self,
                 chart=None,
                 format=None,
                 series=None,
                 pivotArea=None,
                ):
        self.chart = chart
        self.format = format
        self.series = series
        self.pivotArea = pivotArea


class ConditionalFormat(Serialisable):

    tagname = "conditionalFormat"

    scope = Set(values=(['selection', 'data', 'field']))
    type = NoneSet(values=(['all', 'row', 'column']))
    priority = Integer()
    pivotAreas = NestedSequence(expected_type=PivotArea)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotAreas',)

    def __init__(self,
                 scope="selection",
                 type=None,
                 priority=None,
                 pivotAreas=(),
                 extLst=None,
                ):
        self.scope = scope
        self.type = type
        self.priority = priority
        self.pivotAreas = pivotAreas
        self.extLst = extLst


class ConditionalFormatList(Serialisable):

    tagname = "conditionalFormats"

    conditionalFormat = Sequence(expected_type=ConditionalFormat)

    __attrs__ = ("count",)

    def __init__(self, conditionalFormat=(), count=None):
        self.conditionalFormat = conditionalFormat


    def by_priority(self):
        """
        Return a dictionary of format objects keyed by (field id and format property).
        This can be used to map the formats to field but also to dedupe to match
        worksheet definitions which are grouped by cell range
        """

        fmts = {}
        for fmt in self.conditionalFormat:
            for area in fmt.pivotAreas:
                for ref in area.references:
                    for field in ref.x:
                        key = (field.v, fmt.priority)
                        fmts[key] = fmt

        return fmts


    def _dedupe(self):
        """
        Group formats by field index and priority.
        Sorted to match sorting and grouping for corresponding worksheet formats

        The implemtenters notes contain significant deviance from the OOXML
        specification, in particular how conditional formats in tables relate to
        those defined in corresponding worksheets and how to determine which
        format applies to which fields.

        There are some magical interdependencies:

        * Every pivot table fmt must have a worksheet cxf with the same priority.

        * In the reference part the field 4294967294 refers to a data field, the
        spec says -2

        * Data fields are referenced by the 0-index reference.x.v value

        Things are made more complicated by the fact that field items behave
        diffently if the parent is a reference or shared item: "In Office if the
        parent is the reference element, then restrictions of this value are
        defined by reference@field. If the parent is the tables element, then
        this value specifies the index into the table tag position in @url."
        Yeah, right!
        """
        fmts = self.by_priority()
        # sort by priority in order, keeping the highest numerical priority, least when
        # actually applied
        # this is not documented but it's what Excel is happy with
        fmts = {field:fmt for (field, priority), fmt in sorted(fmts.items(), reverse=True)}
        #fmts = {field:fmt for (field, priority), fmt in fmts.items()}
        if fmts:
            self.conditionalFormat = list(fmts.values())


    @property
    def count(self):
        return len(self.conditionalFormat)


    def to_tree(self, tagname=None):
        self._dedupe()
        return super().to_tree(tagname)


class Format(Serialisable):

    tagname = "format"

    action = NoneSet(values=(['blank', 'formatting', 'drill', 'formula']))
    dxfId = Integer(allow_none=True)
    pivotArea = Typed(expected_type=PivotArea, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotArea',)

    def __init__(self,
                 action="formatting",
                 dxfId=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.action = action
        self.dxfId = dxfId
        self.pivotArea = pivotArea
        self.extLst = extLst


class DataField(Serialisable):

    tagname = "dataField"

    name = String(allow_none=True)
    fld = Integer()
    subtotal = Set(values=(['average', 'count', 'countNums', 'max', 'min',
                            'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp']))
    showDataAs = Set(values=(['normal', 'difference', 'percent',
                              'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol',
                              'percentOfTotal', 'index']))
    baseField = Integer()
    baseItem = Integer()
    numFmtId = Integer(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()


    def __init__(self,
                 name=None,
                 fld=None,
                 subtotal="sum",
                 showDataAs="normal",
                 baseField=-1,
                 baseItem=1048832,
                 numFmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fld = fld
        self.subtotal = subtotal
        self.showDataAs = showDataAs
        self.baseField = baseField
        self.baseItem = baseItem
        self.numFmtId = numFmtId
        self.extLst = extLst


class PageField(Serialisable):

    tagname = "pageField"

    fld = Integer()
    item = Integer(allow_none=True)
    hier = Integer(allow_none=True)
    name = String(allow_none=True)
    cap = String(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 fld=None,
                 item=None,
                 hier=None,
                 name=None,
                 cap=None,
                 extLst=None,
                ):
        self.fld = fld
        self.item = item
        self.hier = hier
        self.name = name
        self.cap = cap
        self.extLst = extLst


class RowColItem(Serialisable):

    tagname = "i"

    t = Set(values=(['data', 'default', 'sum', 'countA', 'avg', 'max', 'min',
                     'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand',
                     'blank']))
    r = Integer()
    i = Integer()
    x = Sequence(expected_type=Index, attribute="v")

    __elements__ = ('x',)

    def __init__(self,
                 t="data",
                 r=0,
                 i=0,
                 x=(),
                ):
        self.t = t
        self.r = r
        self.i = i
        self.x = x


class RowColField(Serialisable):

    tagname = "field"

    x = Integer()

    def __init__(self,
                 x=None,
                ):
        self.x = x


class AutoSortScope(Serialisable):

    pivotArea = Typed(expected_type=PivotArea, )

    __elements__ = ('pivotArea',)

    def __init__(self,
                 pivotArea=None,
                ):
        self.pivotArea = pivotArea


class FieldItem(Serialisable):

    tagname = "item"

    n = String(allow_none=True)
    t = Set(values=(['data', 'default', 'sum', 'countA', 'avg', 'max', 'min',
                     'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand',
                     'blank']))
    h = Bool(allow_none=True)
    s = Bool(allow_none=True)
    sd = Bool(allow_none=True)
    f = Bool(allow_none=True)
    m = Bool(allow_none=True)
    c = Bool(allow_none=True)
    x = Integer(allow_none=True)
    d = Bool(allow_none=True)
    e = Bool(allow_none=True)

    def __init__(self,
                 n=None,
                 t="data",
                 h=None,
                 s=None,
                 sd=True,
                 f=None,
                 m=None,
                 c=None,
                 x=None,
                 d=None,
                 e=None,
                ):
        self.n = n
        self.t = t
        self.h = h
        self.s = s
        self.sd = sd
        self.f = f
        self.m = m
        self.c = c
        self.x = x
        self.d = d
        self.e = e


class PivotField(Serialisable):

    tagname = "pivotField"

    items = NestedSequence(expected_type=FieldItem, count=True)
    autoSortScope = Typed(expected_type=AutoSortScope, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    name = String(allow_none=True)
    axis = NoneSet(values=(['axisRow', 'axisCol', 'axisPage', 'axisValues']))
    dataField = Bool(allow_none=True)
    subtotalCaption = String(allow_none=True)
    showDropDowns = Bool(allow_none=True)
    hiddenLevel = Bool(allow_none=True)
    uniqueMemberProperty = String(allow_none=True)
    compact = Bool(allow_none=True)
    allDrilled = Bool(allow_none=True)
    numFmtId = Integer(allow_none=True)
    outline = Bool(allow_none=True)
    subtotalTop = Bool(allow_none=True)
    dragToRow = Bool(allow_none=True)
    dragToCol = Bool(allow_none=True)
    multipleItemSelectionAllowed = Bool(allow_none=True)
    dragToPage = Bool(allow_none=True)
    dragToData = Bool(allow_none=True)
    dragOff = Bool(allow_none=True)
    showAll = Bool(allow_none=True)
    insertBlankRow = Bool(allow_none=True)
    serverField = Bool(allow_none=True)
    insertPageBreak = Bool(allow_none=True)
    autoShow = Bool(allow_none=True)
    topAutoShow = Bool(allow_none=True)
    hideNewItems = Bool(allow_none=True)
    measureFilter = Bool(allow_none=True)
    includeNewItemsInFilter = Bool(allow_none=True)
    itemPageCount = Integer(allow_none=True)
    sortType = Set(values=(['manual', 'ascending', 'descending']))
    dataSourceSort = Bool(allow_none=True)
    nonAutoSortDefault = Bool(allow_none=True)
    rankBy = Integer(allow_none=True)
    defaultSubtotal = Bool(allow_none=True)
    sumSubtotal = Bool(allow_none=True)
    countASubtotal = Bool(allow_none=True)
    avgSubtotal = Bool(allow_none=True)
    maxSubtotal = Bool(allow_none=True)
    minSubtotal = Bool(allow_none=True)
    productSubtotal = Bool(allow_none=True)
    countSubtotal = Bool(allow_none=True)
    stdDevSubtotal = Bool(allow_none=True)
    stdDevPSubtotal = Bool(allow_none=True)
    varSubtotal = Bool(allow_none=True)
    varPSubtotal = Bool(allow_none=True)
    showPropCell = Bool(allow_none=True)
    showPropTip = Bool(allow_none=True)
    showPropAsCaption = Bool(allow_none=True)
    defaultAttributeDrillState = Bool(allow_none=True)

    __elements__ = ('items', 'autoSortScope',)

    def __init__(self,
                 items=(),
                 autoSortScope=None,
                 name=None,
                 axis=None,
                 dataField=None,
                 subtotalCaption=None,
                 showDropDowns=True,
                 hiddenLevel=None,
                 uniqueMemberProperty=None,
                 compact=True,
                 allDrilled=None,
                 numFmtId=None,
                 outline=True,
                 subtotalTop=True,
                 dragToRow=True,
                 dragToCol=True,
                 multipleItemSelectionAllowed=None,
                 dragToPage=True,
                 dragToData=True,
                 dragOff=True,
                 showAll=True,
                 insertBlankRow=None,
                 serverField=None,
                 insertPageBreak=None,
                 autoShow=None,
                 topAutoShow=True,
                 hideNewItems=None,
                 measureFilter=None,
                 includeNewItemsInFilter=None,
                 itemPageCount=10,
                 sortType="manual",
                 dataSourceSort=None,
                 nonAutoSortDefault=None,
                 rankBy=None,
                 defaultSubtotal=True,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 showPropCell=None,
                 showPropTip=None,
                 showPropAsCaption=None,
                 defaultAttributeDrillState=None,
                 extLst=None,
                ):
        self.items = items
        self.autoSortScope = autoSortScope
        self.name = name
        self.axis = axis
        self.dataField = dataField
        self.subtotalCaption = subtotalCaption
        self.showDropDowns = showDropDowns
        self.hiddenLevel = hiddenLevel
        self.uniqueMemberProperty = uniqueMemberProperty
        self.compact = compact
        self.allDrilled = allDrilled
        self.numFmtId = numFmtId
        self.outline = outline
        self.subtotalTop = subtotalTop
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.showAll = showAll
        self.insertBlankRow = insertBlankRow
        self.serverField = serverField
        self.insertPageBreak = insertPageBreak
        self.autoShow = autoShow
        self.topAutoShow = topAutoShow
        self.hideNewItems = hideNewItems
        self.measureFilter = measureFilter
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.itemPageCount = itemPageCount
        self.sortType = sortType
        self.dataSourceSort = dataSourceSort
        self.nonAutoSortDefault = nonAutoSortDefault
        self.rankBy = rankBy
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.showPropCell = showPropCell
        self.showPropTip = showPropTip
        self.showPropAsCaption = showPropAsCaption
        self.defaultAttributeDrillState = defaultAttributeDrillState


class Location(Serialisable):

    tagname = "location"

    ref = String()
    firstHeaderRow = Integer()
    firstDataRow = Integer()
    firstDataCol = Integer()
    rowPageCount = Integer(allow_none=True)
    colPageCount = Integer(allow_none=True)

    def __init__(self,
                 ref=None,
                 firstHeaderRow=None,
                 firstDataRow=None,
                 firstDataCol=None,
                 rowPageCount=None,
                 colPageCount=None,
                ):
        self.ref = ref
        self.firstHeaderRow = firstHeaderRow
        self.firstDataRow = firstDataRow
        self.firstDataCol = firstDataCol
        self.rowPageCount = rowPageCount
        self.colPageCount = colPageCount


class TableDefinition(Serialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
    _id = 1
    _path = "/xl/pivotTables/pivotTable{0}.xml"

    tagname = "pivotTableDefinition"
    cache = None

    name = String()
    cacheId = Integer()
    dataOnRows = Bool()
    dataPosition = Integer(allow_none=True)
    dataCaption = String()
    grandTotalCaption = String(allow_none=True)
    errorCaption = String(allow_none=True)
    showError = Bool()
    missingCaption = String(allow_none=True)
    showMissing = Bool()
    pageStyle = String(allow_none=True)
    pivotTableStyle = String(allow_none=True)
    vacatedStyle = String(allow_none=True)
    tag = String(allow_none=True)
    updatedVersion = Integer()
    minRefreshableVersion = Integer()
    asteriskTotals = Bool()
    showItems = Bool()
    editData = Bool()
    disableFieldList = Bool()
    showCalcMbrs = Bool()
    visualTotals = Bool()
    showMultipleLabel = Bool()
    showDataDropDown = Bool()
    showDrill = Bool()
    printDrill = Bool()
    showMemberPropertyTips = Bool()
    showDataTips = Bool()
    enableWizard = Bool()
    enableDrill = Bool()
    enableFieldProperties = Bool()
    preserveFormatting = Bool()
    useAutoFormatting = Bool()
    pageWrap = Integer()
    pageOverThenDown = Bool()
    subtotalHiddenItems = Bool()
    rowGrandTotals = Bool()
    colGrandTotals = Bool()
    fieldPrintTitles = Bool()
    itemPrintTitles = Bool()
    mergeItem = Bool()
    showDropZones = Bool()
    createdVersion = Integer()
    indent = Integer()
    showEmptyRow = Bool()
    showEmptyCol = Bool()
    showHeaders = Bool()
    compact = Bool()
    outline = Bool()
    outlineData = Bool()
    compactData = Bool()
    published = Bool()
    gridDropZones = Bool()
    immersive = Bool()
    multipleFieldFilters = Bool()
    chartFormat = Integer()
    rowHeaderCaption = String(allow_none=True)
    colHeaderCaption = String(allow_none=True)
    fieldListSortAscending = Bool()
    mdxSubqueries = Bool()
    customListSort = Bool(allow_none=True)
    autoFormatId = Integer(allow_none=True)
    applyNumberFormats = Bool()
    applyBorderFormats = Bool()
    applyFontFormats = Bool()
    applyPatternFormats = Bool()
    applyAlignmentFormats = Bool()
    applyWidthHeightFormats = Bool()
    location = Typed(expected_type=Location, )
    pivotFields = NestedSequence(expected_type=PivotField, count=True)
    rowFields = NestedSequence(expected_type=RowColField, count=True)
    rowItems = NestedSequence(expected_type=RowColItem, count=True)
    colFields = NestedSequence(expected_type=RowColField, count=True)
    colItems = NestedSequence(expected_type=RowColItem, count=True)
    pageFields = NestedSequence(expected_type=PageField, count=True)
    dataFields = NestedSequence(expected_type=DataField, count=True)
    formats = NestedSequence(expected_type=Format, count=True)
    conditionalFormats = Typed(expected_type=ConditionalFormatList, allow_none=True)
    chartFormats = NestedSequence(expected_type=ChartFormat, count=True)
    pivotHierarchies = NestedSequence(expected_type=PivotHierarchy, count=True)
    pivotTableStyleInfo = Typed(expected_type=PivotTableStyle, allow_none=True)
    filters = NestedSequence(expected_type=PivotFilter, count=True)
    rowHierarchiesUsage = Typed(expected_type=RowHierarchiesUsage, allow_none=True)
    colHierarchiesUsage = Typed(expected_type=ColHierarchiesUsage, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    id = Relation()

    __elements__ = ('location', 'pivotFields', 'rowFields', 'rowItems',
                    'colFields', 'colItems', 'pageFields', 'dataFields', 'formats',
                    'conditionalFormats', 'chartFormats', 'pivotHierarchies',
                    'pivotTableStyleInfo', 'filters', 'rowHierarchiesUsage',
                    'colHierarchiesUsage',)

    def __init__(self,
                 name=None,
                 cacheId=None,
                 dataOnRows=False,
                 dataPosition=None,
                 dataCaption=None,
                 grandTotalCaption=None,
                 errorCaption=None,
                 showError=False,
                 missingCaption=None,
                 showMissing=True,
                 pageStyle=None,
                 pivotTableStyle=None,
                 vacatedStyle=None,
                 tag=None,
                 updatedVersion=0,
                 minRefreshableVersion=0,
                 asteriskTotals=False,
                 showItems=True,
                 editData=False,
                 disableFieldList=False,
                 showCalcMbrs=True,
                 visualTotals=True,
                 showMultipleLabel=True,
                 showDataDropDown=True,
                 showDrill=True,
                 printDrill=False,
                 showMemberPropertyTips=True,
                 showDataTips=True,
                 enableWizard=True,
                 enableDrill=True,
                 enableFieldProperties=True,
                 preserveFormatting=True,
                 useAutoFormatting=False,
                 pageWrap=0,
                 pageOverThenDown=False,
                 subtotalHiddenItems=False,
                 rowGrandTotals=True,
                 colGrandTotals=True,
                 fieldPrintTitles=False,
                 itemPrintTitles=False,
                 mergeItem=False,
                 showDropZones=True,
                 createdVersion=0,
                 indent=1,
                 showEmptyRow=False,
                 showEmptyCol=False,
                 showHeaders=True,
                 compact=True,
                 outline=False,
                 outlineData=False,
                 compactData=True,
                 published=False,
                 gridDropZones=False,
                 immersive=True,
                 multipleFieldFilters=None,
                 chartFormat=0,
                 rowHeaderCaption=None,
                 colHeaderCaption=None,
                 fieldListSortAscending=None,
                 mdxSubqueries=None,
                 customListSort=None,
                 autoFormatId=None,
                 applyNumberFormats=False,
                 applyBorderFormats=False,
                 applyFontFormats=False,
                 applyPatternFormats=False,
                 applyAlignmentFormats=False,
                 applyWidthHeightFormats=False,
                 location=None,
                 pivotFields=(),
                 rowFields=(),
                 rowItems=(),
                 colFields=(),
                 colItems=(),
                 pageFields=(),
                 dataFields=(),
                 formats=(),
                 conditionalFormats=None,
                 chartFormats=(),
                 pivotHierarchies=(),
                 pivotTableStyleInfo=None,
                 filters=(),
                 rowHierarchiesUsage=None,
                 colHierarchiesUsage=None,
                 extLst=None,
                 id=None,
                ):
        self.name = name
        self.cacheId = cacheId
        self.dataOnRows = dataOnRows
        self.dataPosition = dataPosition
        self.dataCaption = dataCaption
        self.grandTotalCaption = grandTotalCaption
        self.errorCaption = errorCaption
        self.showError = showError
        self.missingCaption = missingCaption
        self.showMissing = showMissing
        self.pageStyle = pageStyle
        self.pivotTableStyle = pivotTableStyle
        self.vacatedStyle = vacatedStyle
        self.tag = tag
        self.updatedVersion = updatedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.asteriskTotals = asteriskTotals
        self.showItems = showItems
        self.editData = editData
        self.disableFieldList = disableFieldList
        self.showCalcMbrs = showCalcMbrs
        self.visualTotals = visualTotals
        self.showMultipleLabel = showMultipleLabel
        self.showDataDropDown = showDataDropDown
        self.showDrill = showDrill
        self.printDrill = printDrill
        self.showMemberPropertyTips = showMemberPropertyTips
        self.showDataTips = showDataTips
        self.enableWizard = enableWizard
        self.enableDrill = enableDrill
        self.enableFieldProperties = enableFieldProperties
        self.preserveFormatting = preserveFormatting
        self.useAutoFormatting = useAutoFormatting
        self.pageWrap = pageWrap
        self.pageOverThenDown = pageOverThenDown
        self.subtotalHiddenItems = subtotalHiddenItems
        self.rowGrandTotals = rowGrandTotals
        self.colGrandTotals = colGrandTotals
        self.fieldPrintTitles = fieldPrintTitles
        self.itemPrintTitles = itemPrintTitles
        self.mergeItem = mergeItem
        self.showDropZones = showDropZones
        self.createdVersion = createdVersion
        self.indent = indent
        self.showEmptyRow = showEmptyRow
        self.showEmptyCol = showEmptyCol
        self.showHeaders = showHeaders
        self.compact = compact
        self.outline = outline
        self.outlineData = outlineData
        self.compactData = compactData
        self.published = published
        self.gridDropZones = gridDropZones
        self.immersive = immersive
        self.multipleFieldFilters = multipleFieldFilters
        self.chartFormat = chartFormat
        self.rowHeaderCaption = rowHeaderCaption
        self.colHeaderCaption = colHeaderCaption
        self.fieldListSortAscending = fieldListSortAscending
        self.mdxSubqueries = mdxSubqueries
        self.customListSort = customListSort
        self.autoFormatId = autoFormatId
        self.applyNumberFormats = applyNumberFormats
        self.applyBorderFormats = applyBorderFormats
        self.applyFontFormats = applyFontFormats
        self.applyPatternFormats = applyPatternFormats
        self.applyAlignmentFormats = applyAlignmentFormats
        self.applyWidthHeightFormats = applyWidthHeightFormats
        self.location = location
        self.pivotFields = pivotFields
        self.rowFields = rowFields
        self.rowItems = rowItems
        self.colFields = colFields
        self.colItems = colItems
        self.pageFields = pageFields
        self.dataFields = dataFields
        self.formats = formats
        self.conditionalFormats = conditionalFormats
        self.conditionalFormats = None
        self.chartFormats = chartFormats
        self.pivotHierarchies = pivotHierarchies
        self.pivotTableStyleInfo = pivotTableStyleInfo
        self.filters = filters
        self.rowHierarchiesUsage = rowHierarchiesUsage
        self.colHierarchiesUsage = colHierarchiesUsage
        self.extLst = extLst
        self.id = id


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Add to zipfile and update manifest
        """
        self._write_rels(archive, manifest)
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        """
        Write the relevant child objects and add links
        """
        if self.cache is None:
            return

        rels = RelationshipList()
        r = Relationship(Type=self.cache.rel_type, Target=self.cache.path)
        rels.append(r)
        self.id = r.id
        if self.cache.path[1:] not in archive.namelist():
            self.cache._write(archive, manifest)

        path = get_rels_path(self.path)
        xml = tostring(rels.to_tree())
        archive.writestr(path[1:], xml)


    def formatted_fields(self):
        """Map fields to associated conditional formats by priority"""
        if not self.conditionalFormats:
            return {}
        fields = defaultdict(list)
        for idx, prio in self.conditionalFormats.by_priority():
            name = self.dataFields[idx].name
            fields[name].append(prio)
        return fields


    @property
    def summary(self):
        """
        Provide a simplified summary of the table
        """

        return f"{self.name} {dict(self.location)}"

''',
    },
    'openpyxl.reader': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

""",
    },
    'openpyxl.reader.drawings': {
        'is_package': False,
        'source': r'''

# Copyright (c) 2010-2024 openpyxl


from io import BytesIO
from warnings import warn

from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import IMAGE_NS
from openpyxl.packaging.relationship import (
    get_rel,
    get_rels_path,
    get_dependents,
)
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.drawing.image import Image, PILImage
from openpyxl.chart.chartspace import ChartSpace
from openpyxl.chart.reader import read_chart


def find_images(archive, path):
    """
    Given the path to a drawing file extract charts and images

    Ignore errors due to unsupported parts of DrawingML
    """

    src = archive.read(path)
    tree = fromstring(src)
    try:
        drawing = SpreadsheetDrawing.from_tree(tree)
    except TypeError:
        warn("DrawingML support is incomplete and limited to charts and images only. Shapes and drawings will be lost.")
        return [], []

    rels_path = get_rels_path(path)
    deps = []
    if rels_path in archive.namelist():
        deps = get_dependents(archive, rels_path)

    charts = []
    for rel in drawing._chart_rels:
        try:
            cs = get_rel(archive, deps, rel.id, ChartSpace)
        except TypeError as e:
            warn(f"Unable to read chart {rel.id} from {path} {e}")
            continue
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        charts.append(chart)

    images = []
    if not PILImage: # Pillow not installed, drop images
        return charts, images

    for rel in drawing._blip_rels:
        dep = deps.get(rel.embed)
        if dep.Type == IMAGE_NS:
            try:
                image = Image(BytesIO(archive.read(dep.target)))
            except OSError:
                msg = "The image {0} will be removed because it cannot be read".format(dep.target)
                warn(msg)
                continue
            if image.format.upper() == "WMF": # cannot save
                msg = "{0} image format is not supported so the image is being dropped".format(image.format)
                warn(msg)
                continue
            image.anchor = rel.anchor
            images.append(image)
    return charts, images

''',
    },
    'openpyxl.reader.excel': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


"""Read an xlsx file into Python"""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED
from io import BytesIO
import os.path
import warnings

from openpyxl.pivot.table import TableDefinition

# Allow blanket setting of KEEP_VBA for testing
try:
    from ..tests import KEEP_VBA
except ImportError:
    KEEP_VBA = False

# package imports
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.xml.constants import (
    ARC_CORE,
    ARC_CUSTOM,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_THEME,
    COMMENTS_NS,
    SHARED_STRINGS,
    XLTM,
    XLTX,
    XLSM,
    XLSX,
)
from openpyxl.cell import MergedCell
from openpyxl.comments.comment_sheet import CommentSheet

from .strings import read_string_table, read_rich_text
from .workbook import WorkbookParser
from openpyxl.styles.stylesheet import apply_stylesheet

from openpyxl.packaging.core import DocumentProperties
from openpyxl.packaging.custom import CustomPropertyList
from openpyxl.packaging.manifest import Manifest, Override

from openpyxl.packaging.relationship import (
    RelationshipList,
    get_dependents,
    get_rels_path,
)

from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet._reader import WorksheetReader
from openpyxl.chartsheet import Chartsheet
from openpyxl.worksheet.table import Table
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing

from openpyxl.xml.functions import fromstring

from .drawings import find_images


SUPPORTED_FORMATS = ('.xlsx', '.xlsm', '.xltx', '.xltm')


def _validate_archive(filename):
    """
    Does a first check whether filename is a string or a file-like
    object. If it is a string representing a filename, a check is done
    for supported formats by checking the given file-extension. If the
    file-extension is not in SUPPORTED_FORMATS an InvalidFileException
    will raised. Otherwise the filename (resp. file-like object) will
    forwarded to zipfile.ZipFile returning a ZipFile-Instance.
    """
    is_file_like = hasattr(filename, 'read')
    if not is_file_like:
        file_format = os.path.splitext(filename)[-1].lower()
        if file_format not in SUPPORTED_FORMATS:
            if file_format == '.xls':
                msg = ('openpyxl does not support the old .xls file format, '
                       'please use xlrd to read this file, or convert it to '
                       'the more recent .xlsx file format.')
            elif file_format == '.xlsb':
                msg = ('openpyxl does not support binary format .xlsb, '
                       'please convert this file to .xlsx format if you want '
                       'to open it with openpyxl')
            else:
                msg = ('openpyxl does not support %s file format, '
                       'please check you can open '
                       'it with Excel first. '
                       'Supported formats are: %s') % (file_format,
                                                       ','.join(SUPPORTED_FORMATS))
            raise InvalidFileException(msg)

    archive = ZipFile(filename, 'r')
    return archive


def _find_workbook_part(package):
    workbook_types = [XLTM, XLTX, XLSM, XLSX]
    for ct in workbook_types:
        part = package.find(ct)
        if part:
            return part

    # some applications reassign the default for application/xml
    defaults = {p.ContentType for p in package.Default}
    workbook_type = defaults & set(workbook_types)
    if workbook_type:
        return Override("/" + ARC_WORKBOOK, workbook_type.pop())

    raise IOError("File contains no valid workbook part")


class ExcelReader:

    """
    Read an Excel package and dispatch the contents to the relevant modules
    """

    def __init__(self, fn, read_only=False, keep_vba=KEEP_VBA,
                 data_only=False, keep_links=True, rich_text=False):
        self.archive = _validate_archive(fn)
        self.valid_files = self.archive.namelist()
        self.read_only = read_only
        self.keep_vba = keep_vba
        self.data_only = data_only
        self.keep_links = keep_links
        self.rich_text = rich_text
        self.shared_strings = []


    def read_manifest(self):
        src = self.archive.read(ARC_CONTENT_TYPES)
        root = fromstring(src)
        self.package = Manifest.from_tree(root)


    def read_strings(self):
        ct = self.package.find(SHARED_STRINGS)
        reader = read_string_table
        if self.rich_text:
            reader = read_rich_text
        if ct is not None:
            strings_path = ct.PartName[1:]
            with self.archive.open(strings_path,) as src:
                self.shared_strings = reader(src)


    def read_workbook(self):
        wb_part = _find_workbook_part(self.package)
        self.parser = WorkbookParser(self.archive, wb_part.PartName[1:], keep_links=self.keep_links)
        self.parser.parse()
        wb = self.parser.wb
        wb._sheets = []
        wb._data_only = self.data_only
        wb._read_only = self.read_only
        wb.template = wb_part.ContentType in (XLTX, XLTM)

        # If are going to preserve the vba then attach a copy of the archive to the
        # workbook so that is available for the save.
        if self.keep_vba:
            wb.vba_archive = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
            for name in self.valid_files:
                wb.vba_archive.writestr(name, self.archive.read(name))

        if self.read_only:
            wb._archive = self.archive

        self.wb = wb


    def read_properties(self):
        if ARC_CORE in self.valid_files:
            src = fromstring(self.archive.read(ARC_CORE))
            self.wb.properties = DocumentProperties.from_tree(src)


    def read_custom(self):
        if ARC_CUSTOM in self.valid_files:
            src = fromstring(self.archive.read(ARC_CUSTOM))
            self.wb.custom_doc_props = CustomPropertyList.from_tree(src)


    def read_theme(self):
        if ARC_THEME in self.valid_files:
            self.wb.loaded_theme = self.archive.read(ARC_THEME)


    def read_chartsheet(self, sheet, rel):
        sheet_path = rel.target
        rels_path = get_rels_path(sheet_path)
        rels = []
        if rels_path in self.valid_files:
            rels = get_dependents(self.archive, rels_path)

        with self.archive.open(sheet_path, "r") as src:
            xml = src.read()
        node = fromstring(xml)
        cs = Chartsheet.from_tree(node)
        cs._parent = self.wb
        cs.title = sheet.name
        self.wb._add_sheet(cs)

        drawings = rels.find(SpreadsheetDrawing._rel_type)
        for rel in drawings:
            charts, images = find_images(self.archive, rel.target)
            for c in charts:
                cs.add_chart(c)


    def read_worksheets(self):
        comment_warning = """Cell '{0}':{1} is part of a merged range but has a comment which will be removed because merged cells cannot contain any data."""
        for sheet, rel in self.parser.find_sheets():
            if rel.target not in self.valid_files:
                continue

            if "chartsheet" in rel.Type:
                self.read_chartsheet(sheet, rel)
                continue

            rels_path = get_rels_path(rel.target)
            rels = RelationshipList()
            if rels_path in self.valid_files:
                rels = get_dependents(self.archive, rels_path)

            if self.read_only:
                ws = ReadOnlyWorksheet(self.wb, sheet.name, rel.target, self.shared_strings)
                ws.sheet_state = sheet.state
                self.wb._sheets.append(ws)
                continue
            else:
                fh = self.archive.open(rel.target)
                ws = self.wb.create_sheet(sheet.name)
                ws._rels = rels
                ws_parser = WorksheetReader(ws, fh, self.shared_strings, self.data_only, self.rich_text)
                ws_parser.bind_all()
                fh.close()

            # assign any comments to cells
            for r in rels.find(COMMENTS_NS):
                src = self.archive.read(r.target)
                comment_sheet = CommentSheet.from_tree(fromstring(src))
                for ref, comment in comment_sheet.comments:
                    try:
                        ws[ref].comment = comment
                    except AttributeError:
                        c = ws[ref]
                        if isinstance(c, MergedCell):
                            warnings.warn(comment_warning.format(ws.title, c.coordinate))
                            continue

            # preserve link to VML file if VBA
            if self.wb.vba_archive and ws.legacy_drawing:
                ws.legacy_drawing = rels.get(ws.legacy_drawing).target
            else:
                ws.legacy_drawing = None

            for t in ws_parser.tables:
                src = self.archive.read(t)
                xml = fromstring(src)
                table = Table.from_tree(xml)
                ws.add_table(table)

            drawings = rels.find(SpreadsheetDrawing._rel_type)
            for rel in drawings:
                charts, images = find_images(self.archive, rel.target)
                for c in charts:
                    ws.add_chart(c, c.anchor)
                for im in images:
                    ws.add_image(im, im.anchor)

            pivot_rel = rels.find(TableDefinition.rel_type)
            pivot_caches = self.parser.pivot_caches
            for r in pivot_rel:
                pivot_path = r.Target
                src = self.archive.read(pivot_path)
                tree = fromstring(src)
                pivot = TableDefinition.from_tree(tree)
                pivot.cache = pivot_caches[pivot.cacheId]
                ws.add_pivot(pivot)

            ws.sheet_state = sheet.state


    def read(self):
        action = "read manifest"
        try:
            self.read_manifest()
            action = "read strings"
            self.read_strings()
            action = "read workbook"
            self.read_workbook()
            action = "read properties"
            self.read_properties()
            action = "read custom properties"
            self.read_custom()
            action = "read theme"
            self.read_theme()
            action = "read stylesheet"
            apply_stylesheet(self.archive, self.wb)
            action = "read worksheets"
            self.read_worksheets()
            action = "assign names"
            self.parser.assign_names()
            if not self.read_only:
                self.archive.close()
        except ValueError as e:
            raise ValueError(
                f"Unable to read workbook: could not {action} from {self.archive.filename}.\n"
                "This is most probably because the workbook source files contain some invalid XML.\n"
                "Please see the exception for more details."
                ) from e


def load_workbook(filename, read_only=False, keep_vba=KEEP_VBA,
                  data_only=False, keep_links=True, rich_text=False):
    """Open the given filename and return the workbook

    :param filename: the path to open or a file-like object
    :type filename: string or a file-like object open in binary mode c.f., :class:`zipfile.ZipFile`

    :param read_only: optimised for reading, content cannot be edited
    :type read_only: bool

    :param keep_vba: preserve vba content (this does NOT mean you can use it)
    :type keep_vba: bool

    :param data_only: controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet
    :type data_only: bool

    :param keep_links: whether links to external workbooks should be preserved. The default is True
    :type keep_links: bool

    :param rich_text: if set to True openpyxl will preserve any rich text formatting in cells. The default is False
    :type rich_text: bool

    :rtype: :class:`openpyxl.workbook.Workbook`

    .. note::

        When using lazy load, all worksheets will be :class:`openpyxl.worksheet.iter_worksheet.IterableWorksheet`
        and the returned workbook will be read-only.

    """
    reader = ExcelReader(filename, read_only, keep_vba,
                         data_only, keep_links, rich_text)
    reader.read()
    return reader.wb

''',
    },
    'openpyxl.reader.strings': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.cell.text import Text

from openpyxl.xml.functions import iterparse
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.cell.rich_text import CellRichText


def read_string_table(xml_source):
    """Read in all shared strings in the table"""

    strings = []
    STRING_TAG = '{%s}si' % SHEET_MAIN_NS

    for _, node in iterparse(xml_source):
        if node.tag == STRING_TAG:
            text = Text.from_tree(node).content
            text = text.replace('x005F_', '')
            node.clear()

            strings.append(text)

    return strings


def read_rich_text(xml_source):
    """Read in all shared strings in the table"""

    strings = []
    STRING_TAG = '{%s}si' % SHEET_MAIN_NS

    for _, node in iterparse(xml_source):
        if node.tag == STRING_TAG:
            text = CellRichText.from_tree(node)
            if len(text) == 0:
                text = ''
            elif len(text) == 1 and isinstance(text[0], str):
                text = text[0]
            node.clear()

            strings.append(text)

    return strings

''',
    },
    'openpyxl.reader.workbook': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from warnings import warn

from openpyxl.xml.functions import fromstring

from openpyxl.packaging.relationship import (
    get_dependents,
    get_rels_path,
    get_rel,
)
from openpyxl.packaging.workbook import WorkbookPackage
from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedNameList
from openpyxl.workbook.external_link.external import read_external_link
from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.worksheet.print_settings import PrintTitles, PrintArea

from openpyxl.utils.datetime import CALENDAR_MAC_1904


class WorkbookParser:

    _rels = None

    def __init__(self, archive, workbook_part_name, keep_links=True):
        self.archive = archive
        self.workbook_part_name = workbook_part_name
        self.defined_names = DefinedNameList()
        self.wb = Workbook()
        self.keep_links = keep_links
        self.sheets = []


    @property
    def rels(self):
        if self._rels is None:
            self._rels = get_dependents(self.archive, get_rels_path(self.workbook_part_name)).to_dict()
        return self._rels


    def parse(self):
        src = self.archive.read(self.workbook_part_name)
        node = fromstring(src)
        package = WorkbookPackage.from_tree(node)
        if package.properties.date1904:
            self.wb.epoch = CALENDAR_MAC_1904

        self.wb.code_name = package.properties.codeName
        self.wb.active = package.active
        self.wb.views = package.bookViews
        self.sheets = package.sheets
        self.wb.calculation = package.calcPr
        self.caches = package.pivotCaches

        # external links contain cached worksheets and can be very big
        if not self.keep_links:
            package.externalReferences = []

        for ext_ref in package.externalReferences:
            rel = self.rels.get(ext_ref.id)
            self.wb._external_links.append(
                read_external_link(self.archive, rel.Target)
            )

        if package.definedNames:
            self.defined_names = package.definedNames

        self.wb.security = package.workbookProtection


    def find_sheets(self):
        """
        Find all sheets in the workbook and return the link to the source file.

        Older XLSM files sometimes contain invalid sheet elements.
        Warn user when these are removed.
        """

        for sheet in self.sheets:
            if not sheet.id:
                msg = f"File contains an invalid specification for {0}. This will be removed".format(sheet.name)
                warn(msg)
                continue
            yield sheet, self.rels[sheet.id]


    def assign_names(self):
        """
        Bind defined names and other definitions to worksheets or the workbook
        """

        for idx, names in self.defined_names.by_sheet().items():
            if idx == "global":
                self.wb.defined_names = names
                continue

            try:
                sheet = self.wb._sheets[idx]
            except IndexError:
                warn(f"Defined names for sheet index {idx} cannot be located")
                continue

            for name, defn in names.items():
                reserved = defn.is_reserved
                if reserved is None:
                    sheet.defined_names[name] = defn

                elif reserved == "Print_Titles":
                    titles = PrintTitles.from_string(defn.value)
                    sheet._print_rows = titles.rows
                    sheet._print_cols = titles.cols
                elif reserved == "Print_Area":
                    try:
                        sheet._print_area = PrintArea.from_string(defn.value)
                    except TypeError:
                        warn(f"Print area cannot be set to Defined name: {defn.value}.")
                        continue

    @property
    def pivot_caches(self):
        """
        Get PivotCache objects
        """
        d = {}
        for c in self.caches:
            cache = get_rel(self.archive, self.rels, id=c.id, cls=CacheDefinition)
            if cache.deps:
                records = get_rel(self.archive, cache.deps, cache.id, RecordList)
                cache.records = records
            d[c.cacheId] = cache
        return d

''',
    },
    'openpyxl.styles': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from .alignment import Alignment
from .borders import Border, Side
from .colors import Color
from .fills import PatternFill, GradientFill, Fill
from .fonts import Font, DEFAULT_FONT
from .numbers import NumberFormatDescriptor, is_date_format, is_builtin
from .protection import Protection
from .named_styles import NamedStyle

""",
    },
    'openpyxl.styles.alignment': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string

from openpyxl.descriptors import Bool, MinMax, Min, Alias, NoneSet
from openpyxl.descriptors.serialisable import Serialisable


horizontal_alignments = (
    "general", "left", "center", "right", "fill", "justify", "centerContinuous",
    "distributed", )
vertical_aligments = (
    "top", "center", "bottom", "justify", "distributed",
)

class Alignment(Serialisable):
    """Alignment options for use in styles."""

    tagname = "alignment"

    horizontal = NoneSet(values=horizontal_alignments)
    vertical = NoneSet(values=vertical_aligments)
    textRotation = NoneSet(values=range(181))
    textRotation.values.add(255)
    text_rotation = Alias('textRotation')
    wrapText = Bool(allow_none=True)
    wrap_text = Alias('wrapText')
    shrinkToFit = Bool(allow_none=True)
    shrink_to_fit = Alias('shrinkToFit')
    indent = MinMax(min=0, max=255)
    relativeIndent = MinMax(min=-255, max=255)
    justifyLastLine = Bool(allow_none=True)
    readingOrder = Min(min=0)

    def __init__(self, horizontal=None, vertical=None,
                 textRotation=0, wrapText=None, shrinkToFit=None, indent=0, relativeIndent=0,
                 justifyLastLine=None, readingOrder=0, text_rotation=None,
                 wrap_text=None, shrink_to_fit=None, mergeCell=None):
        self.horizontal = horizontal
        self.vertical = vertical
        self.indent = indent
        self.relativeIndent = relativeIndent
        self.justifyLastLine = justifyLastLine
        self.readingOrder = readingOrder
        if text_rotation is not None:
            textRotation = text_rotation
        if textRotation is not None:
            self.textRotation = int(textRotation)
        if wrap_text is not None:
            wrapText = wrap_text
        self.wrapText = wrapText
        if shrink_to_fit is not None:
            shrinkToFit = shrink_to_fit
        self.shrinkToFit = shrinkToFit
        # mergeCell is vestigial


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value is not None and value != 0:
                yield attr, safe_string(value)

''',
    },
    'openpyxl.styles.borders': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string
from openpyxl.descriptors import (
    NoneSet,
    Typed,
    Bool,
    Alias,
    Sequence,
    Integer,
)
from openpyxl.descriptors.serialisable import Serialisable

from .colors import ColorDescriptor


BORDER_NONE = None
BORDER_DASHDOT = 'dashDot'
BORDER_DASHDOTDOT = 'dashDotDot'
BORDER_DASHED = 'dashed'
BORDER_DOTTED = 'dotted'
BORDER_DOUBLE = 'double'
BORDER_HAIR = 'hair'
BORDER_MEDIUM = 'medium'
BORDER_MEDIUMDASHDOT = 'mediumDashDot'
BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
BORDER_MEDIUMDASHED = 'mediumDashed'
BORDER_SLANTDASHDOT = 'slantDashDot'
BORDER_THICK = 'thick'
BORDER_THIN = 'thin'


class Side(Serialisable):

    """Border options for use in styles.
    Caution: if you do not specify a border_style, other attributes will
    have no effect !"""


    color = ColorDescriptor(allow_none=True)
    style = NoneSet(values=('dashDot','dashDotDot', 'dashed','dotted',
                            'double','hair', 'medium', 'mediumDashDot', 'mediumDashDotDot',
                            'mediumDashed', 'slantDashDot', 'thick', 'thin')
                    )
    border_style = Alias('style')

    def __init__(self, style=None, color=None, border_style=None):
        if border_style is not None:
            style = border_style
        self.style = style
        self.color = color


class Border(Serialisable):
    """Border positioning for use in styles."""

    tagname = "border"

    __elements__ = ('start', 'end', 'left', 'right', 'top', 'bottom',
                    'diagonal', 'vertical', 'horizontal')

    # child elements
    start = Typed(expected_type=Side, allow_none=True)
    end = Typed(expected_type=Side, allow_none=True)
    left = Typed(expected_type=Side, allow_none=True)
    right = Typed(expected_type=Side, allow_none=True)
    top = Typed(expected_type=Side, allow_none=True)
    bottom = Typed(expected_type=Side, allow_none=True)
    diagonal = Typed(expected_type=Side, allow_none=True)
    vertical = Typed(expected_type=Side, allow_none=True)
    horizontal = Typed(expected_type=Side, allow_none=True)
    # attributes
    outline = Bool()
    diagonalUp = Bool()
    diagonalDown = Bool()

    def __init__(self, left=None, right=None, top=None,
                 bottom=None, diagonal=None, diagonal_direction=None,
                 vertical=None, horizontal=None, diagonalUp=False, diagonalDown=False,
                 outline=True, start=None, end=None):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal
        self.vertical = vertical
        self.horizontal = horizontal
        self.diagonal_direction = diagonal_direction
        self.diagonalUp = diagonalUp
        self.diagonalDown = diagonalDown
        self.outline = outline
        self.start = start
        self.end = end

    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value and attr != "outline":
                yield attr, safe_string(value)
            elif attr == "outline" and not value:
                yield attr, safe_string(value)

DEFAULT_BORDER = Border(left=Side(), right=Side(), top=Side(), bottom=Side(), diagonal=Side())

''',
    },
    'openpyxl.styles.builtins': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

# Builtins styles as defined in Part 4 Annex G.2

from .named_styles import NamedStyle
from openpyxl.xml.functions import fromstring


normal = """
  <namedStyle builtinId="0" name="Normal">
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

comma = """
  <namedStyle builtinId="3" name="Comma">
    <alignment/>
    <number_format>_-* #,##0.00\\ _$_-;\\-* #,##0.00\\ _$_-;_-* "-"??\\ _$_-;_-@_-</number_format>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

comma_0 = """
  <namedStyle builtinId="6" name="Comma [0]">
    <alignment/>
    <number_format>_-* #,##0\\ _$_-;\\-* #,##0\\ _$_-;_-* "-"\\ _$_-;_-@_-</number_format>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

currency = """
  <namedStyle builtinId="4" name="Currency">
    <alignment/>
    <number_format>_-* #,##0.00\\ "$"_-;\\-* #,##0.00\\ "$"_-;_-* "-"??\\ "$"_-;_-@_-</number_format>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

currency_0 = """
  <namedStyle builtinId="7" name="Currency [0]">
    <alignment/>
    <number_format>_-* #,##0\\ "$"_-;\\-* #,##0\\ "$"_-;_-* "-"\\ "$"_-;_-@_-</number_format>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

percent = """
  <namedStyle builtinId="5" name="Percent">
    <alignment/>
    <number_format>0%</number_format>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

hyperlink = """
  <namedStyle builtinId="8" name="Hyperlink" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="10"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

followed_hyperlink = """
  <namedStyle builtinId="9" name="Followed Hyperlink" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="11"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

title = """
  <namedStyle builtinId="15" name="Title">
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Cambria"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="3"/>
      <sz val="18"/>
      <scheme val="major"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

headline_1 = """
  <namedStyle builtinId="16" name="Headline 1" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style="thick">
        <color theme="4"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="3"/>
      <sz val="15"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

headline_2 = """
  <namedStyle builtinId="17" name="Headline 2" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style="thick">
        <color theme="4" tint="0.5"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="3"/>
      <sz val="13"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

headline_3 = """
   <namedStyle builtinId="18" name="Headline 3" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style="medium">
        <color theme="4" tint="0.4"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="3"/>
      <sz val="11"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>

"""

headline_4 = """
  <namedStyle builtinId="19" name="Headline 4">
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="3"/>
      <sz val="11"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

good = """
  <namedStyle builtinId="26" name="Good" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFC6EFCE"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color rgb="FF006100"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

bad = """
  <namedStyle builtinId="27" name="Bad" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFC7CE"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color rgb="FF9C0006"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

neutral = """
  <namedStyle builtinId="28" name="Neutral" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFEB9C"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color rgb="FF9C6500"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

input = """
  <namedStyle builtinId="20" name="Input" >
    <alignment/>
    <border>
      <left style="thin">
        <color rgb="FF7F7F7F"/>
      </left>
      <right style="thin">
        <color rgb="FF7F7F7F"/>
      </right>
      <top style="thin">
        <color rgb="FF7F7F7F"/>
      </top>
      <bottom style="thin">
        <color rgb="FF7F7F7F"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFCC99"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color rgb="FF3F3F76"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

output = """
  <namedStyle builtinId="21" name="Output" >
    <alignment/>
    <border>
      <left style="thin">
        <color rgb="FF3F3F3F"/>
      </left>
      <right style="thin">
        <color rgb="FF3F3F3F"/>
      </right>
      <top style="thin">
        <color rgb="FF3F3F3F"/>
      </top>
      <bottom style="thin">
        <color rgb="FF3F3F3F"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFF2F2F2"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color rgb="FF3F3F3F"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

calculation = """
  <namedStyle builtinId="22" name="Calculation" >
    <alignment/>
    <border>
      <left style="thin">
        <color rgb="FF7F7F7F"/>
      </left>
      <right style="thin">
        <color rgb="FF7F7F7F"/>
      </right>
      <top style="thin">
        <color rgb="FF7F7F7F"/>
      </top>
      <bottom style="thin">
        <color rgb="FF7F7F7F"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFF2F2F2"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color rgb="FFFA7D00"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

linked_cell = """
  <namedStyle builtinId="24" name="Linked Cell" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom style="double">
        <color rgb="FFFF8001"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color rgb="FFFA7D00"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

check_cell = """
  <namedStyle builtinId="23" name="Check Cell" >
    <alignment/>
    <border>
      <left style="double">
        <color rgb="FF3F3F3F"/>
      </left>
      <right style="double">
        <color rgb="FF3F3F3F"/>
      </right>
      <top style="double">
        <color rgb="FF3F3F3F"/>
      </top>
      <bottom style="double">
        <color rgb="FF3F3F3F"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFA5A5A5"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

warning = """
  <namedStyle builtinId="11" name="Warning Text" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color rgb="FFFF0000"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

note = """
  <namedStyle builtinId="10" name="Note" >
    <alignment/>
    <border>
      <left style="thin">
        <color rgb="FFB2B2B2"/>
      </left>
      <right style="thin">
        <color rgb="FFB2B2B2"/>
      </right>
      <top style="thin">
        <color rgb="FFB2B2B2"/>
      </top>
      <bottom style="thin">
        <color rgb="FFB2B2B2"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="FFFFFFCC"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

explanatory = """
  <namedStyle builtinId="53" name="Explanatory Text" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <i val="1"/>
      <color rgb="FF7F7F7F"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

total = """
  <namedStyle builtinId="25" name="Total" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top style="thin">
        <color theme="4"/>
      </top>
      <bottom style="double">
        <color theme="4"/>
      </bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <b val="1"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_1 = """
  <namedStyle builtinId="29" name="Accent1" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="4"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_1_20 = """
  <namedStyle builtinId="30" name="20 % - Accent1" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="4" tint="0.7999816888943144"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_1_40 = """
  <namedStyle builtinId="31" name="40 % - Accent1" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="4" tint="0.5999938962981048"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_1_60 = """
  <namedStyle builtinId="32" name="60 % - Accent1" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="4" tint="0.3999755851924192"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_2 = """<namedStyle builtinId="33" name="Accent2" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="5"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

accent_2_20 = """
  <namedStyle builtinId="34" name="20 % - Accent2" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="5" tint="0.7999816888943144"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

accent_2_40 = """
<namedStyle builtinId="35" name="40 % - Accent2" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="5" tint="0.5999938962981048"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

accent_2_60 = """
<namedStyle builtinId="36" name="60 % - Accent2" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="5" tint="0.3999755851924192"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

accent_3 = """
<namedStyle builtinId="37" name="Accent3" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="6"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

accent_3_20 = """
  <namedStyle builtinId="38" name="20 % - Accent3" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="6" tint="0.7999816888943144"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>"""

accent_3_40 = """
  <namedStyle builtinId="39" name="40 % - Accent3" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="6" tint="0.5999938962981048"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""
accent_3_60 = """
  <namedStyle builtinId="40" name="60 % - Accent3" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="6" tint="0.3999755851924192"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""
accent_4 = """
  <namedStyle builtinId="41" name="Accent4" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="7"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_4_20 = """
  <namedStyle builtinId="42" name="20 % - Accent4" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="7" tint="0.7999816888943144"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_4_40 = """
  <namedStyle builtinId="43" name="40 % - Accent4" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="7" tint="0.5999938962981048"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_4_60 = """
<namedStyle builtinId="44" name="60 % - Accent4" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="7" tint="0.3999755851924192"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_5 = """
  <namedStyle builtinId="45" name="Accent5" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="8"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_5_20 = """
  <namedStyle builtinId="46" name="20 % - Accent5" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="8" tint="0.7999816888943144"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_5_40 = """
  <namedStyle builtinId="47" name="40 % - Accent5" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="8" tint="0.5999938962981048"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_5_60 = """
  <namedStyle builtinId="48" name="60 % - Accent5" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="8" tint="0.3999755851924192"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_6 = """
  <namedStyle builtinId="49" name="Accent6" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="9"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_6_20 = """
  <namedStyle builtinId="50" name="20 % - Accent6" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="9" tint="0.7999816888943144"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_6_40 = """
  <namedStyle builtinId="51" name="40 % - Accent6" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="9" tint="0.5999938962981048"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="1"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

accent_6_60 = """
  <namedStyle builtinId="52" name="60 % - Accent6" >
    <alignment/>
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
    <fill>
      <patternFill patternType="solid">
        <fgColor theme="9" tint="0.3999755851924192"/>
        <bgColor indexed="65"/>
      </patternFill>
    </fill>
    <font>
      <name val="Calibri"/>
      <family val="2"/>
      <color theme="0"/>
      <sz val="12"/>
      <scheme val="minor"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

pandas_highlight = """
  <namedStyle hidden="0" name="Pandas">
    <alignment horizontal="center"/>
    <border>
      <left style="thin"><color rgb="00000000"/></left>
      <right style="thin"><color rgb="00000000"/></right>
      <top style="thin"><color rgb="00000000"/></top>
      <bottom style="thin"><color rgb="00000000"/></bottom>
      <diagonal/>
    </border>
    <fill>
      <patternFill/>
    </fill>
    <font>
      <b val="1"/>
    </font>
    <protection hidden="0" locked="1"/>
  </namedStyle>
"""

styles = dict(
    [
        ('Normal', NamedStyle.from_tree(fromstring(normal))),
        ('Comma', NamedStyle.from_tree(fromstring(comma))),
        ('Currency', NamedStyle.from_tree(fromstring(currency))),
        ('Percent', NamedStyle.from_tree(fromstring(percent))),
        ('Comma [0]', NamedStyle.from_tree(fromstring(comma_0))),
        ('Currency [0]', NamedStyle.from_tree(fromstring(currency_0))),
        ('Hyperlink', NamedStyle.from_tree(fromstring(hyperlink))),
        ('Followed Hyperlink', NamedStyle.from_tree(fromstring(followed_hyperlink))),
        ('Note', NamedStyle.from_tree(fromstring(note))),
        ('Warning Text', NamedStyle.from_tree(fromstring(warning))),
        ('Title', NamedStyle.from_tree(fromstring(title))),
        ('Headline 1', NamedStyle.from_tree(fromstring(headline_1))),
        ('Headline 2', NamedStyle.from_tree(fromstring(headline_2))),
        ('Headline 3', NamedStyle.from_tree(fromstring(headline_3))),
        ('Headline 4', NamedStyle.from_tree(fromstring(headline_4))),
        ('Input', NamedStyle.from_tree(fromstring(input))),
        ('Output', NamedStyle.from_tree(fromstring(output))),
        ('Calculation',NamedStyle.from_tree(fromstring(calculation))),
        ('Check Cell', NamedStyle.from_tree(fromstring(check_cell))),
        ('Linked Cell', NamedStyle.from_tree(fromstring(linked_cell))),
        ('Total', NamedStyle.from_tree(fromstring(total))),
        ('Good', NamedStyle.from_tree(fromstring(good))),
        ('Bad', NamedStyle.from_tree(fromstring(bad))),
        ('Neutral', NamedStyle.from_tree(fromstring(neutral))),
        ('Accent1', NamedStyle.from_tree(fromstring(accent_1))),
        ('20 % - Accent1', NamedStyle.from_tree(fromstring(accent_1_20))),
        ('40 % - Accent1', NamedStyle.from_tree(fromstring(accent_1_40))),
        ('60 % - Accent1', NamedStyle.from_tree(fromstring(accent_1_60))),
        ('Accent2', NamedStyle.from_tree(fromstring(accent_2))),
        ('20 % - Accent2', NamedStyle.from_tree(fromstring(accent_2_20))),
        ('40 % - Accent2', NamedStyle.from_tree(fromstring(accent_2_40))),
        ('60 % - Accent2', NamedStyle.from_tree(fromstring(accent_2_60))),
        ('Accent3', NamedStyle.from_tree(fromstring(accent_3))),
        ('20 % - Accent3', NamedStyle.from_tree(fromstring(accent_3_20))),
        ('40 % - Accent3', NamedStyle.from_tree(fromstring(accent_3_40))),
        ('60 % - Accent3', NamedStyle.from_tree(fromstring(accent_3_60))),
        ('Accent4', NamedStyle.from_tree(fromstring(accent_4))),
        ('20 % - Accent4', NamedStyle.from_tree(fromstring(accent_4_20))),
        ('40 % - Accent4', NamedStyle.from_tree(fromstring(accent_4_40))),
        ('60 % - Accent4', NamedStyle.from_tree(fromstring(accent_4_60))),
        ('Accent5', NamedStyle.from_tree(fromstring(accent_5))),
        ('20 % - Accent5', NamedStyle.from_tree(fromstring(accent_5_20))),
        ('40 % - Accent5', NamedStyle.from_tree(fromstring(accent_5_40))),
        ('60 % - Accent5', NamedStyle.from_tree(fromstring(accent_5_60))),
        ('Accent6', NamedStyle.from_tree(fromstring(accent_6))),
        ('20 % - Accent6', NamedStyle.from_tree(fromstring(accent_6_20))),
        ('40 % - Accent6', NamedStyle.from_tree(fromstring(accent_6_40))),
        ('60 % - Accent6', NamedStyle.from_tree(fromstring(accent_6_60))),
        ('Explanatory Text', NamedStyle.from_tree(fromstring(explanatory))),
        ('Pandas', NamedStyle.from_tree(fromstring(pandas_highlight)))
    ]
)

''',
    },
    'openpyxl.styles.cell_style': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from array import array

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Bool,
    Integer,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.utils.indexed_list import IndexedList


from .alignment import Alignment
from .protection import Protection


class ArrayDescriptor:

    def __init__(self, key):
        self.key = key

    def __get__(self, instance, cls):
        return instance[self.key]

    def __set__(self, instance, value):
        instance[self.key] = value


class StyleArray(array):
    """
    Simplified named tuple with an array
    """

    __slots__ = ()
    tagname = 'xf'

    fontId = ArrayDescriptor(0)
    fillId = ArrayDescriptor(1)
    borderId = ArrayDescriptor(2)
    numFmtId = ArrayDescriptor(3)
    protectionId = ArrayDescriptor(4)
    alignmentId = ArrayDescriptor(5)
    pivotButton = ArrayDescriptor(6)
    quotePrefix = ArrayDescriptor(7)
    xfId = ArrayDescriptor(8)


    def __new__(cls, args=[0]*9):
        return array.__new__(cls, 'i', args)


    def __hash__(self):
        return hash(tuple(self))


    def __copy__(self):
        return StyleArray((self))


    def __deepcopy__(self, memo):
        return StyleArray((self))


class CellStyle(Serialisable):

    tagname = "xf"

    numFmtId = Integer()
    fontId = Integer()
    fillId = Integer()
    borderId = Integer()
    xfId = Integer(allow_none=True)
    quotePrefix = Bool(allow_none=True)
    pivotButton = Bool(allow_none=True)
    applyNumberFormat = Bool(allow_none=True)
    applyFont = Bool(allow_none=True)
    applyFill = Bool(allow_none=True)
    applyBorder = Bool(allow_none=True)
    applyAlignment = Bool(allow_none=True)
    applyProtection = Bool(allow_none=True)
    alignment = Typed(expected_type=Alignment, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('alignment', 'protection')
    __attrs__ = ("numFmtId", "fontId", "fillId", "borderId",
                 "applyAlignment", "applyProtection", "pivotButton", "quotePrefix", "xfId")

    def __init__(self,
                 numFmtId=0,
                 fontId=0,
                 fillId=0,
                 borderId=0,
                 xfId=None,
                 quotePrefix=None,
                 pivotButton=None,
                 applyNumberFormat=None,
                 applyFont=None,
                 applyFill=None,
                 applyBorder=None,
                 applyAlignment=None,
                 applyProtection=None,
                 alignment=None,
                 protection=None,
                 extLst=None,
                ):
        self.numFmtId = numFmtId
        self.fontId = fontId
        self.fillId = fillId
        self.borderId = borderId
        self.xfId = xfId
        self.quotePrefix = quotePrefix
        self.pivotButton = pivotButton
        self.applyNumberFormat = applyNumberFormat
        self.applyFont = applyFont
        self.applyFill = applyFill
        self.applyBorder = applyBorder
        self.alignment = alignment
        self.protection = protection


    def to_array(self):
        """
        Convert to StyleArray
        """
        style = StyleArray()
        for k in ("fontId", "fillId", "borderId", "numFmtId", "pivotButton",
                  "quotePrefix", "xfId"):
            v = getattr(self, k, 0)
            if v is not None:
                setattr(style, k, v)
        return style


    @classmethod
    def from_array(cls, style):
        """
        Convert from StyleArray
        """
        return cls(numFmtId=style.numFmtId, fontId=style.fontId,
                   fillId=style.fillId, borderId=style.borderId, xfId=style.xfId,
                   quotePrefix=style.quotePrefix, pivotButton=style.pivotButton,)


    @property
    def applyProtection(self):
        return self.protection is not None or None


    @property
    def applyAlignment(self):
        return self.alignment is not None or None


class CellStyleList(Serialisable):

    tagname = "cellXfs"

    __attrs__ = ("count",)

    count = Integer(allow_none=True)
    xf = Sequence(expected_type=CellStyle)
    alignment = Sequence(expected_type=Alignment)
    protection = Sequence(expected_type=Protection)

    __elements__ = ('xf',)

    def __init__(self,
                 count=None,
                 xf=(),
                ):
        self.xf = xf


    @property
    def count(self):
        return len(self.xf)


    def __getitem__(self, idx):
        try:
            return self.xf[idx]
        except IndexError:
            print((f"{idx} is out of range"))
        return self.xf[idx]


    def _to_array(self):
        """
        Extract protection and alignments, convert to style array
        """
        self.prots = IndexedList([Protection()])
        self.alignments = IndexedList([Alignment()])
        styles = [] # allow duplicates
        for xf in self.xf:
            style = xf.to_array()
            if xf.alignment is not None:
                style.alignmentId = self.alignments.add(xf.alignment)
            if xf.protection is not None:
                style.protectionId = self.prots.add(xf.protection)
            styles.append(style)
        return IndexedList(styles)

''',
    },
    'openpyxl.styles.colors': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import re
from openpyxl.compat import safe_string
from openpyxl.descriptors import (
    String,
    Bool,
    MinMax,
    Integer,
    Typed,
)
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.descriptors.serialisable import Serialisable

# Default Color Index as per 18.8.27 of ECMA Part 4
COLOR_INDEX = (
    '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', #0-4
    '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF', #5-9
    '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', #10-14
    '0000FFFF', '00800000', '00008000', '00000080', '00808000', #15-19
    '00800080', '00008080', '00C0C0C0', '00808080', '009999FF', #20-24
    '00993366', '00FFFFCC', '00CCFFFF', '00660066', '00FF8080', #25-29
    '000066CC', '00CCCCFF', '00000080', '00FF00FF', '00FFFF00', #30-34
    '0000FFFF', '00800080', '00800000', '00008080', '000000FF', #35-39
    '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', #40-44
    '00FF99CC', '00CC99FF', '00FFCC99', '003366FF', '0033CCCC', #45-49
    '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699', #50-54
    '00969696', '00003366', '00339966', '00003300', '00333300', #55-59
    '00993300', '00993366', '00333399', '00333333',  #60-63
)
# indices 64 and 65 are reserved for the system foreground and background colours respectively

# Will remove these definitions in a future release
BLACK = COLOR_INDEX[0]
WHITE = COLOR_INDEX[1]
#RED = COLOR_INDEX[2]
#DARKRED = COLOR_INDEX[8]
BLUE = COLOR_INDEX[4]
#DARKBLUE = COLOR_INDEX[12]
#GREEN = COLOR_INDEX[3]
#DARKGREEN = COLOR_INDEX[9]
#YELLOW = COLOR_INDEX[5]
#DARKYELLOW = COLOR_INDEX[19]


aRGB_REGEX = re.compile("^([A-Fa-f0-9]{8}|[A-Fa-f0-9]{6})$")


class RGB(Typed):
    """
    Descriptor for aRGB values
    If not supplied alpha is 00
    """

    expected_type = str

    def __set__(self, instance, value):
        if not self.allow_none:
            m = aRGB_REGEX.match(value)
            if m is None:
                raise ValueError("Colors must be aRGB hex values")
            if len(value) == 6:
                value = "00" + value
        super().__set__(instance, value)


class Color(Serialisable):
    """Named colors for use in styles."""

    tagname = "color"

    rgb = RGB()
    indexed = Integer()
    auto = Bool()
    theme = Integer()
    tint = MinMax(min=-1, max=1, expected_type=float)
    type = String()


    def __init__(self, rgb=BLACK, indexed=None, auto=None, theme=None, tint=0.0, index=None, type='rgb'):
        if index is not None:
            indexed = index
        if indexed is not None:
            self.type = 'indexed'
            self.indexed = indexed
        elif theme is not None:
            self.type = 'theme'
            self.theme = theme
        elif auto is not None:
            self.type = 'auto'
            self.auto = auto
        else:
            self.rgb = rgb
            self.type = 'rgb'
        self.tint = tint

    @property
    def value(self):
        return getattr(self, self.type)

    @value.setter
    def value(self, value):
        setattr(self, self.type, value)

    def __iter__(self):
        attrs = [(self.type, self.value)]
        if self.tint != 0:
            attrs.append(('tint', self.tint))
        for k, v in attrs:
            yield k, safe_string(v)

    @property
    def index(self):
        # legacy
        return self.value


    def __add__(self, other):
        """
        Adding colours is undefined behaviour best do nothing
        """
        if not isinstance(other, Color):
            return super().__add__(other)
        return self


class ColorDescriptor(Typed):

    expected_type = Color

    def __set__(self, instance, value):
        if isinstance(value, str):
            value = Color(rgb=value)
        super().__set__(instance, value)


class RgbColor(Serialisable):

    tagname = "rgbColor"

    rgb = RGB()

    def __init__(self,
                 rgb=None,
                ):
        self.rgb = rgb


class ColorList(Serialisable):

    tagname = "colors"

    indexedColors = NestedSequence(expected_type=RgbColor)
    mruColors = NestedSequence(expected_type=Color)

    __elements__ = ('indexedColors', 'mruColors')

    def __init__(self,
                 indexedColors=(),
                 mruColors=(),
                ):
        self.indexedColors = indexedColors
        self.mruColors = mruColors


    def __bool__(self):
        return bool(self.indexedColors) or bool(self.mruColors)


    @property
    def index(self):
        return [val.rgb for val in self.indexedColors]

''',
    },
    'openpyxl.styles.differential': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Typed,
    Sequence,
    Alias,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles import (
    Font,
    Fill,
    Border,
    Alignment,
    Protection,
    )
from .numbers import NumberFormat


class DifferentialStyle(Serialisable):

    tagname = "dxf"

    __elements__ = ("font", "numFmt", "fill", "alignment", "border", "protection")

    font = Typed(expected_type=Font, allow_none=True)
    numFmt = Typed(expected_type=NumberFormat, allow_none=True)
    fill = Typed(expected_type=Fill, allow_none=True)
    alignment = Typed(expected_type=Alignment, allow_none=True)
    border = Typed(expected_type=Border, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)

    def __init__(self,
                 font=None,
                 numFmt=None,
                 fill=None,
                 alignment=None,
                 border=None,
                 protection=None,
                 extLst=None,
                ):
        self.font = font
        self.numFmt = numFmt
        self.fill = fill
        self.alignment = alignment
        self.border = border
        self.protection = protection
        self.extLst = extLst


class DifferentialStyleList(Serialisable):
    """
    Dedupable container for differential styles.
    """

    tagname = "dxfs"

    dxf = Sequence(expected_type=DifferentialStyle)
    styles = Alias("dxf")
    __attrs__ = ("count",)


    def __init__(self, dxf=(), count=None):
        self.dxf = dxf


    def append(self, dxf):
        """
        Check to see whether style already exists and append it if does not.
        """
        if not isinstance(dxf, DifferentialStyle):
            raise TypeError('expected ' + str(DifferentialStyle))
        if dxf in self.styles:
            return
        self.styles.append(dxf)


    def add(self, dxf):
        """
        Add a differential style and return its index
        """
        self.append(dxf)
        return self.styles.index(dxf)


    def __bool__(self):
        return bool(self.styles)


    def __getitem__(self, idx):
        return self.styles[idx]


    @property
    def count(self):
        return len(self.dxf)

''',
    },
    'openpyxl.styles.fills': {
        'is_package': False,
        'source': r'''

# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Float,
    Set,
    Alias,
    NoneSet,
    Sequence,
    Integer,
    MinMax,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.compat import safe_string

from .colors import ColorDescriptor, Color

from openpyxl.xml.functions import Element, localname
from openpyxl.xml.constants import SHEET_MAIN_NS


FILL_NONE = 'none'
FILL_SOLID = 'solid'
FILL_PATTERN_DARKDOWN = 'darkDown'
FILL_PATTERN_DARKGRAY = 'darkGray'
FILL_PATTERN_DARKGRID = 'darkGrid'
FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal'
FILL_PATTERN_DARKTRELLIS = 'darkTrellis'
FILL_PATTERN_DARKUP = 'darkUp'
FILL_PATTERN_DARKVERTICAL = 'darkVertical'
FILL_PATTERN_GRAY0625 = 'gray0625'
FILL_PATTERN_GRAY125 = 'gray125'
FILL_PATTERN_LIGHTDOWN = 'lightDown'
FILL_PATTERN_LIGHTGRAY = 'lightGray'
FILL_PATTERN_LIGHTGRID = 'lightGrid'
FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal'
FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis'
FILL_PATTERN_LIGHTUP = 'lightUp'
FILL_PATTERN_LIGHTVERTICAL = 'lightVertical'
FILL_PATTERN_MEDIUMGRAY = 'mediumGray'

fills = (FILL_SOLID, FILL_PATTERN_DARKDOWN, FILL_PATTERN_DARKGRAY,
         FILL_PATTERN_DARKGRID, FILL_PATTERN_DARKHORIZONTAL, FILL_PATTERN_DARKTRELLIS,
         FILL_PATTERN_DARKUP, FILL_PATTERN_DARKVERTICAL, FILL_PATTERN_GRAY0625,
         FILL_PATTERN_GRAY125, FILL_PATTERN_LIGHTDOWN, FILL_PATTERN_LIGHTGRAY,
         FILL_PATTERN_LIGHTGRID, FILL_PATTERN_LIGHTHORIZONTAL,
         FILL_PATTERN_LIGHTTRELLIS, FILL_PATTERN_LIGHTUP, FILL_PATTERN_LIGHTVERTICAL,
         FILL_PATTERN_MEDIUMGRAY)


class Fill(Serialisable):

    """Base class"""

    tagname = "fill"

    @classmethod
    def from_tree(cls, el):
        children = [c for c in el]
        if not children:
            return
        child = children[0]
        if "patternFill" in child.tag:
            return PatternFill._from_tree(child)
        return super(Fill, GradientFill).from_tree(child)


class PatternFill(Fill):
    """Area fill patterns for use in styles.
    Caution: if you do not specify a fill_type, other attributes will have
    no effect !"""

    tagname = "patternFill"

    __elements__ = ('fgColor', 'bgColor')

    patternType = NoneSet(values=fills)
    fill_type = Alias("patternType")
    fgColor = ColorDescriptor()
    start_color = Alias("fgColor")
    bgColor = ColorDescriptor()
    end_color = Alias("bgColor")

    def __init__(self, patternType=None, fgColor=Color(), bgColor=Color(),
                 fill_type=None, start_color=None, end_color=None):
        if fill_type is not None:
            patternType = fill_type
        self.patternType = patternType
        if start_color is not None:
            fgColor = start_color
        self.fgColor = fgColor
        if end_color is not None:
            bgColor = end_color
        self.bgColor = bgColor

    @classmethod
    def _from_tree(cls, el):
        attrib = dict(el.attrib)
        for child in el:
            desc = localname(child)
            attrib[desc] = Color.from_tree(child)
        return cls(**attrib)


    def to_tree(self, tagname=None, idx=None):
        parent = Element("fill")
        el = Element(self.tagname)
        if self.patternType is not None:
            el.set('patternType', self.patternType)
        for c in self.__elements__:
            value = getattr(self, c)
            if value != Color():
                el.append(value.to_tree(c))
        parent.append(el)
        return parent


DEFAULT_EMPTY_FILL = PatternFill()
DEFAULT_GRAY_FILL = PatternFill(patternType='gray125')


class Stop(Serialisable):

    tagname = "stop"

    position = MinMax(min=0, max=1)
    color = ColorDescriptor()

    def __init__(self, color, position):
        self.position = position
        self.color = color


def _assign_position(values):
    """
    Automatically assign positions if a list of colours is provided.

    It is not permitted to mix colours and stops
    """
    n_values = len(values)
    n_stops = sum(isinstance(value, Stop) for value in values)

    if n_stops == 0:
        interval = 1
        if n_values > 2:
            interval = 1 / (n_values - 1)
        values = [Stop(value, i * interval)
                  for i, value in enumerate(values)]

    elif n_stops < n_values:
        raise ValueError('Cannot interpret mix of Stops and Colors in GradientFill')

    pos = set()
    for stop in values:
        if stop.position in pos:
            raise ValueError("Duplicate position {0}".format(stop.position))
        pos.add(stop.position)

    return values


class StopList(Sequence):

    expected_type = Stop

    def __set__(self, obj, values):
        values = _assign_position(values)
        super().__set__(obj, values)


class GradientFill(Fill):
    """Fill areas with gradient

    Two types of gradient fill are supported:

        - A type='linear' gradient interpolates colours between
          a set of specified Stops, across the length of an area.
          The gradient is left-to-right by default, but this
          orientation can be modified with the degree
          attribute.  A list of Colors can be provided instead
          and they will be positioned with equal distance between them.

        - A type='path' gradient applies a linear gradient from each
          edge of the area. Attributes top, right, bottom, left specify
          the extent of fill from the respective borders. Thus top="0.2"
          will fill the top 20% of the cell.

    """

    tagname = "gradientFill"

    type = Set(values=('linear', 'path'))
    fill_type = Alias("type")
    degree = Float()
    left = Float()
    right = Float()
    top = Float()
    bottom = Float()
    stop = StopList()


    def __init__(self, type="linear", degree=0, left=0, right=0, top=0,
                 bottom=0, stop=()):
        self.degree = degree
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.stop = stop
        self.type = type


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value:
                yield attr, safe_string(value)


    def to_tree(self, tagname=None, namespace=None, idx=None):
        parent = Element("fill")
        el = super().to_tree()
        parent.append(el)
        return parent

''',
    },
    'openpyxl.styles.fonts': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors import (
    Alias,
    Sequence,
    Integer
)
from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.descriptors.nested import (
    NestedValue,
    NestedBool,
    NestedNoneSet,
    NestedMinMax,
    NestedString,
    NestedInteger,
    NestedFloat,
)
from .colors import ColorDescriptor, Color, BLACK

from openpyxl.compat import safe_string
from openpyxl.xml.functions import Element, SubElement
from openpyxl.xml.constants import SHEET_MAIN_NS


def _no_value(tagname, value, namespace=None):
    if value:
        return Element(tagname, val=safe_string(value))


class Font(Serialisable):
    """Font options used in styles."""

    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'

    name = NestedString(allow_none=True)
    charset = NestedInteger(allow_none=True)
    family = NestedMinMax(min=0, max=14, allow_none=True)
    sz = NestedFloat(allow_none=True)
    size = Alias("sz")
    b = NestedBool(to_tree=_no_value)
    bold = Alias("b")
    i = NestedBool(to_tree=_no_value)
    italic = Alias("i")
    strike = NestedBool(allow_none=True)
    strikethrough = Alias("strike")
    outline = NestedBool(allow_none=True)
    shadow = NestedBool(allow_none=True)
    condense = NestedBool(allow_none=True)
    extend = NestedBool(allow_none=True)
    u = NestedNoneSet(values=('single', 'double', 'singleAccounting',
                             'doubleAccounting'))
    underline = Alias("u")
    vertAlign = NestedNoneSet(values=('superscript', 'subscript', 'baseline'))
    color = ColorDescriptor(allow_none=True)
    scheme = NestedNoneSet(values=("major", "minor"))

    tagname = "font"

    __elements__ = ('name', 'charset', 'family', 'b', 'i', 'strike', 'outline',
                  'shadow', 'condense', 'color', 'extend', 'sz', 'u', 'vertAlign',
                  'scheme')


    def __init__(self, name=None, sz=None, b=None, i=None, charset=None,
                 u=None, strike=None, color=None, scheme=None, family=None, size=None,
                 bold=None, italic=None, strikethrough=None, underline=None,
                 vertAlign=None, outline=None, shadow=None, condense=None,
                 extend=None):
        self.name = name
        self.family = family
        if size is not None:
            sz = size
        self.sz = sz
        if bold is not None:
            b = bold
        self.b = b
        if italic is not None:
            i = italic
        self.i = i
        if underline is not None:
            u = underline
        self.u = u
        if strikethrough is not None:
            strike = strikethrough
        self.strike = strike
        self.color = color
        self.vertAlign = vertAlign
        self.charset = charset
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.scheme = scheme


    @classmethod
    def from_tree(cls, node):
        """
        Set default value for underline if child element is present
        """
        underline = node.find("{%s}u" % SHEET_MAIN_NS)
        if underline is not None and underline.get('val') is None:
            underline.set("val", "single")
        return super().from_tree(node)


DEFAULT_FONT = Font(name="Calibri", sz=11, family=2, b=False, i=False,
                    color=Color(theme=1), scheme="minor")

''',
    },
    'openpyxl.styles.named_styles': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string

from openpyxl.descriptors import (
    Typed,
    Integer,
    Bool,
    String,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable

from .fills import PatternFill, Fill
from .fonts import Font
from .borders import Border
from .alignment import Alignment
from .protection import Protection
from .numbers import (
    NumberFormatDescriptor,
    BUILTIN_FORMATS_MAX_SIZE,
    BUILTIN_FORMATS_REVERSE,
)
from .cell_style import (
    StyleArray,
    CellStyle,
)


class NamedStyle(Serialisable):

    """
    Named and editable styles
    """

    font = Typed(expected_type=Font)
    fill = Typed(expected_type=Fill)
    border = Typed(expected_type=Border)
    alignment = Typed(expected_type=Alignment)
    number_format = NumberFormatDescriptor()
    protection = Typed(expected_type=Protection)
    builtinId = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    name = String()
    _wb = None
    _style = StyleArray()


    def __init__(self,
                 name="Normal",
                 font=None,
                 fill=None,
                 border=None,
                 alignment=None,
                 number_format=None,
                 protection=None,
                 builtinId=None,
                 hidden=False,
                 ):
        self.name = name
        self.font = font or Font()
        self.fill = fill or PatternFill()
        self.border = border or Border()
        self.alignment = alignment or Alignment()
        self.number_format = number_format
        self.protection = protection or Protection()
        self.builtinId = builtinId
        self.hidden = hidden
        self._wb = None
        self._style = StyleArray()


    def __setattr__(self, attr, value):
        super().__setattr__(attr, value)
        if getattr(self, '_wb', None) and attr in (
           'font', 'fill', 'border', 'alignment', 'number_format', 'protection',
            ):
            self._recalculate()


    def __iter__(self):
        for key in ('name', 'builtinId', 'hidden', 'xfId'):
            value = getattr(self, key, None)
            if value is not None:
                yield key, safe_string(value)


    def bind(self, wb):
        """
        Bind a named style to a workbook
        """
        self._wb = wb
        self._recalculate()


    def _recalculate(self):
        self._style.fontId =  self._wb._fonts.add(self.font)
        self._style.borderId = self._wb._borders.add(self.border)
        self._style.fillId =  self._wb._fills.add(self.fill)
        self._style.protectionId = self._wb._protections.add(self.protection)
        self._style.alignmentId = self._wb._alignments.add(self.alignment)
        fmt = self.number_format
        if fmt in BUILTIN_FORMATS_REVERSE:
            fmt = BUILTIN_FORMATS_REVERSE[fmt]
        else:
            fmt = self._wb._number_formats.add(self.number_format) + (
                  BUILTIN_FORMATS_MAX_SIZE)
        self._style.numFmtId = fmt


    def as_tuple(self):
        """Return a style array representing the current style"""
        return self._style


    def as_xf(self):
        """
        Return equivalent XfStyle
        """
        xf = CellStyle.from_array(self._style)
        xf.xfId = None
        xf.pivotButton = None
        xf.quotePrefix = None
        if self.alignment != Alignment():
            xf.alignment = self.alignment
        if self.protection != Protection():
            xf.protection = self.protection
        return xf


    def as_name(self):
        """
        Return relevant named style

        """
        named = _NamedCellStyle(
            name=self.name,
            builtinId=self.builtinId,
            hidden=self.hidden,
            xfId=self._style.xfId
        )
        return named


class NamedStyleList(list):
    """
    Named styles are editable and can be applied to multiple objects

    As only the index is stored in referencing objects the order mus
    be preserved.

    Returns a list of NamedStyles
    """

    def __init__(self, iterable=()):
        """
        Allow a list of named styles to be passed in and index them.
        """

        for idx, s in enumerate(iterable, len(self)):
            s._style.xfId = idx
        super().__init__(iterable)


    @property
    def names(self):
        return [s.name for s in self]


    def __getitem__(self, key):
        if isinstance(key, int):
            return super().__getitem__(key)


        for idx, name in enumerate(self.names):
            if name == key:
                return self[idx]

        raise KeyError("No named style with the name{0} exists".format(key))

    def append(self, style):
        if not isinstance(style, NamedStyle):
            raise TypeError("""Only NamedStyle instances can be added""")
        elif style.name in self.names: # hotspot
            raise ValueError("""Style {0} exists already""".format(style.name))
        style._style.xfId = (len(self))
        super().append(style)


class _NamedCellStyle(Serialisable):

    """
    Pointer-based representation of named styles in XML
    xfId refers to the corresponding CellStyleXfs

    Not used in client code.
    """

    tagname = "cellStyle"

    name = String()
    xfId = Integer()
    builtinId = Integer(allow_none=True)
    iLevel = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    customBuiltin = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()


    def __init__(self,
                 name=None,
                 xfId=None,
                 builtinId=None,
                 iLevel=None,
                 hidden=None,
                 customBuiltin=None,
                 extLst=None,
                ):
        self.name = name
        self.xfId = xfId
        self.builtinId = builtinId
        self.iLevel = iLevel
        self.hidden = hidden
        self.customBuiltin = customBuiltin


class _NamedCellStyleList(Serialisable):
    """
    Container for named cell style objects

    Not used in client code
    """

    tagname = "cellStyles"

    count = Integer(allow_none=True)
    cellStyle = Sequence(expected_type=_NamedCellStyle)

    __attrs__ = ("count",)

    def __init__(self,
                 count=None,
                 cellStyle=(),
                ):
        self.cellStyle = cellStyle


    @property
    def count(self):
        return len(self.cellStyle)


    def remove_duplicates(self):
        """
        Some applications contain duplicate definitions either by name or
        referenced style.

        As the references are 0-based indices, styles are sorted by
        index.

        Returns a list of style references with duplicates removed
        """

        def sort_fn(v):
            return v.xfId

        styles = []
        names = set()
        ids = set()

        for ns in sorted(self.cellStyle, key=sort_fn):
            if ns.xfId in ids or ns.name in names: # skip duplicates
                continue
            ids.add(ns.xfId)
            names.add(ns.name)

            styles.append(ns)

        return styles

''',
    },
    'openpyxl.styles.numbers': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import re

from openpyxl.descriptors import (
    String,
    Sequence,
    Integer,
)
from openpyxl.descriptors.serialisable import Serialisable


BUILTIN_FORMATS = {
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    5: '"$"#,##0_);("$"#,##0)',
    6: '"$"#,##0_);[Red]("$"#,##0)',
    7: '"$"#,##0.00_);("$"#,##0.00)',
    8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'mm-dd-yy',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'm/d/yy h:mm',

    37: '#,##0_);(#,##0)',
    38: '#,##0_);[Red](#,##0)',
    39: '#,##0.00_);(#,##0.00)',
    40: '#,##0.00_);[Red](#,##0.00)',

    41: r'_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)',
    42: r'_("$"* #,##0_);_("$"* \(#,##0\);_("$"* "-"_);_(@_)',
    43: r'_(* #,##0.00_);_(* \(#,##0.00\);_(* "-"??_);_(@_)',

    44: r'_("$"* #,##0.00_)_("$"* \(#,##0.00\)_("$"* "-"??_)_(@_)',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mmss.0',
    48: '##0.0E+0',
    49: '@', }

BUILTIN_FORMATS_MAX_SIZE = 164
BUILTIN_FORMATS_REVERSE = dict(
        [(value, key) for key, value in BUILTIN_FORMATS.items()])

FORMAT_GENERAL = BUILTIN_FORMATS[0]
FORMAT_TEXT = BUILTIN_FORMATS[49]
FORMAT_NUMBER = BUILTIN_FORMATS[1]
FORMAT_NUMBER_00 = BUILTIN_FORMATS[2]
FORMAT_NUMBER_COMMA_SEPARATED1 = BUILTIN_FORMATS[4]
FORMAT_NUMBER_COMMA_SEPARATED2 = '#,##0.00_-'
FORMAT_PERCENTAGE = BUILTIN_FORMATS[9]
FORMAT_PERCENTAGE_00 = BUILTIN_FORMATS[10]
FORMAT_DATE_YYYYMMDD2 = 'yyyy-mm-dd'
FORMAT_DATE_YYMMDD = 'yy-mm-dd'
FORMAT_DATE_DDMMYY = 'dd/mm/yy'
FORMAT_DATE_DMYSLASH = 'd/m/y'
FORMAT_DATE_DMYMINUS = 'd-m-y'
FORMAT_DATE_DMMINUS = 'd-m'
FORMAT_DATE_MYMINUS = 'm-y'
FORMAT_DATE_XLSX14 = BUILTIN_FORMATS[14]
FORMAT_DATE_XLSX15 = BUILTIN_FORMATS[15]
FORMAT_DATE_XLSX16 = BUILTIN_FORMATS[16]
FORMAT_DATE_XLSX17 = BUILTIN_FORMATS[17]
FORMAT_DATE_XLSX22 = BUILTIN_FORMATS[22]
FORMAT_DATE_DATETIME = 'yyyy-mm-dd h:mm:ss'
FORMAT_DATE_TIME1 = BUILTIN_FORMATS[18]
FORMAT_DATE_TIME2 = BUILTIN_FORMATS[19]
FORMAT_DATE_TIME3 = BUILTIN_FORMATS[20]
FORMAT_DATE_TIME4 = BUILTIN_FORMATS[21]
FORMAT_DATE_TIME5 = BUILTIN_FORMATS[45]
FORMAT_DATE_TIME6 = BUILTIN_FORMATS[21]
FORMAT_DATE_TIME7 = 'i:s.S'
FORMAT_DATE_TIME8 = 'h:mm:ss@'
FORMAT_DATE_TIMEDELTA = '[hh]:mm:ss'
FORMAT_DATE_YYMMDDSLASH = 'yy/mm/dd@'
FORMAT_CURRENCY_USD_SIMPLE = '"$"#,##0.00_-'
FORMAT_CURRENCY_USD = '$#,##0_-'
FORMAT_CURRENCY_EUR_SIMPLE = '[$EUR ]#,##0.00_-'


COLORS = r"\[(BLACK|BLUE|CYAN|GREEN|MAGENTA|RED|WHITE|YELLOW)\]"
LITERAL_GROUP = r'".*?"' # anything in quotes
LOCALE_GROUP = r'\[(?!hh?\]|mm?\]|ss?\])[^\]]*\]' # anything in square brackets, except hours or minutes or seconds
STRIP_RE = re.compile(f"{LITERAL_GROUP}|{LOCALE_GROUP}")
TIMEDELTA_RE = re.compile(r'\[hh?\](:mm(:ss(\.0*)?)?)?|\[mm?\](:ss(\.0*)?)?|\[ss?\](\.0*)?', re.I)


# Spec 18.8.31 numFmts
# +ve;-ve;zero;text

def is_date_format(fmt):
    if fmt is None:
        return False
    fmt = fmt.split(";")[0] # only look at the first format
    fmt = STRIP_RE.sub("", fmt) # ignore some formats
    return re.search(r"(?<![_\\])[dmhysDMHYS]", fmt) is not None


def is_timedelta_format(fmt):
    if fmt is None:
        return False
    fmt = fmt.split(";")[0] # only look at the first format
    return TIMEDELTA_RE.search(fmt) is not None


def is_datetime(fmt):
    """
    Return date, time or datetime
    """
    if not is_date_format(fmt):
        return

    DATE = TIME = False

    if any((x in fmt for x in 'dy')):
        DATE = True
    if any((x in fmt for x in 'hs')):
        TIME = True

    if DATE and TIME:
        return "datetime"
    if DATE:
        return "date"
    return "time"


def is_builtin(fmt):
    return fmt in BUILTIN_FORMATS.values()


def builtin_format_code(index):
    """Return one of the standard format codes by index."""
    try:
        fmt = BUILTIN_FORMATS[index]
    except KeyError:
        fmt = None
    return fmt


def builtin_format_id(fmt):
    """Return the id of a standard style."""
    return BUILTIN_FORMATS_REVERSE.get(fmt)


class NumberFormatDescriptor(String):

    def __set__(self, instance, value):
        if value is None:
            value = FORMAT_GENERAL
        super().__set__(instance, value)


class NumberFormat(Serialisable):

    numFmtId = Integer()
    formatCode = String()

    def __init__(self,
                 numFmtId=None,
                 formatCode=None,
                ):
        self.numFmtId = numFmtId
        self.formatCode = formatCode


class NumberFormatList(Serialisable):

    count = Integer(allow_none=True)
    numFmt = Sequence(expected_type=NumberFormat)

    __elements__ = ('numFmt',)
    __attrs__ = ("count",)

    def __init__(self,
                 count=None,
                 numFmt=(),
                ):
        self.numFmt = numFmt


    @property
    def count(self):
        return len(self.numFmt)


    def __getitem__(self, idx):
        return self.numFmt[idx]

''',
    },
    'openpyxl.styles.protection': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import Bool
from openpyxl.descriptors.serialisable import Serialisable


class Protection(Serialisable):
    """Protection options for use in styles."""

    tagname = "protection"

    locked = Bool()
    hidden = Bool()

    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden

''',
    },
    'openpyxl.styles.proxy': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from copy import copy

from openpyxl.compat import deprecated


class StyleProxy:
    """
    Proxy formatting objects so that they cannot be altered
    """

    __slots__ = ('__target')

    def __init__(self, target):
        self.__target = target


    def __repr__(self):
        return repr(self.__target)


    def __getattr__(self, attr):
        return getattr(self.__target, attr)


    def __setattr__(self, attr, value):
        if attr != "_StyleProxy__target":
            raise AttributeError("Style objects are immutable and cannot be changed."
                                 "Reassign the style with a copy")
        super().__setattr__(attr, value)


    def __copy__(self):
        """
        Return a copy of the proxied object.
        """
        return copy(self.__target)


    def __add__(self, other):
        """
        Add proxied object to another instance and return the combined object
        """
        return self.__target + other


    @deprecated("Use copy(obj) or cell.obj = cell.obj + other")
    def copy(self, **kw):
        """Return a copy of the proxied object. Keyword args will be passed through"""
        cp = copy(self.__target)
        for k, v in kw.items():
            setattr(cp, k, v)
        return cp


    def __eq__(self, other):
        return self.__target == other


    def __ne__(self, other):
        return not self == other

''',
    },
    'openpyxl.styles.styleable': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from copy import copy

from .numbers import (
    BUILTIN_FORMATS,
    BUILTIN_FORMATS_MAX_SIZE,
    BUILTIN_FORMATS_REVERSE,
)
from .proxy import StyleProxy
from .cell_style import StyleArray
from .named_styles import NamedStyle
from .builtins import styles


class StyleDescriptor:

    def __init__(self, collection, key):
        self.collection = collection
        self.key = key

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, coll.add(value))


    def __get__(self, instance, cls):
        coll = getattr(instance.parent.parent, self.collection)
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx =  getattr(instance._style, self.key)
        return StyleProxy(coll[idx])


class NumberFormatDescriptor:

    key = "numFmtId"
    collection = '_number_formats'

    def __set__(self, instance, value):
        coll = getattr(instance.parent.parent, self.collection)
        if value in BUILTIN_FORMATS_REVERSE:
            idx = BUILTIN_FORMATS_REVERSE[value]
        else:
            idx = coll.add(value) + BUILTIN_FORMATS_MAX_SIZE

        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        setattr(instance._style, self.key, idx)


    def __get__(self, instance, cls):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        if idx < BUILTIN_FORMATS_MAX_SIZE:
            return BUILTIN_FORMATS.get(idx, "General")
        coll = getattr(instance.parent.parent, self.collection)
        return coll[idx - BUILTIN_FORMATS_MAX_SIZE]


class NamedStyleDescriptor:

    key = "xfId"
    collection = "_named_styles"


    def __set__(self, instance, value):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        coll = getattr(instance.parent.parent, self.collection)
        if isinstance(value, NamedStyle):
            style = value
            if style not in coll:
                instance.parent.parent.add_named_style(style)
        elif value not in coll.names:
            if value in styles: # is it builtin?
                style = styles[value]
                if style not in coll:
                    instance.parent.parent.add_named_style(style)
            else:
                raise ValueError("{0} is not a known style".format(value))
        else:
            style = coll[value]
        instance._style = copy(style.as_tuple())


    def __get__(self, instance, cls):
        if not getattr(instance, "_style"):
            instance._style = StyleArray()
        idx = getattr(instance._style, self.key)
        coll = getattr(instance.parent.parent, self.collection)
        return coll.names[idx]


class StyleArrayDescriptor:

    def __init__(self, key):
        self.key = key

    def __set__(self, instance, value):
        if instance._style is None:
            instance._style = StyleArray()
        setattr(instance._style, self.key, value)


    def __get__(self, instance, cls):
        if instance._style is None:
            return False
        return bool(getattr(instance._style, self.key))


class StyleableObject:
    """
    Base class for styleble objects implementing proxy and lookup functions
    """

    font = StyleDescriptor('_fonts', "fontId")
    fill = StyleDescriptor('_fills', "fillId")
    border = StyleDescriptor('_borders', "borderId")
    number_format = NumberFormatDescriptor()
    protection = StyleDescriptor('_protections', "protectionId")
    alignment = StyleDescriptor('_alignments', "alignmentId")
    style = NamedStyleDescriptor()
    quotePrefix = StyleArrayDescriptor('quotePrefix')
    pivotButton = StyleArrayDescriptor('pivotButton')

    __slots__ = ('parent', '_style')

    def __init__(self, sheet, style_array=None):
        self.parent = sheet
        if style_array is not None:
            style_array = StyleArray(style_array)
        self._style = style_array


    @property
    def style_id(self):
        if self._style is None:
            self._style = StyleArray()
        return self.parent.parent._cell_styles.add(self._style)


    @property
    def has_style(self):
        if self._style is None:
            return False
        return any(self._style)


''',
    },
    'openpyxl.styles.stylesheet': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from warnings import warn

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
)
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.xml.constants import ARC_STYLE, SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring

from .builtins import styles
from .colors import ColorList
from .differential import DifferentialStyle
from .table import TableStyleList
from .borders import Border
from .fills import Fill
from .fonts import Font
from .numbers import (
    NumberFormatList,
    BUILTIN_FORMATS,
    BUILTIN_FORMATS_MAX_SIZE,
    BUILTIN_FORMATS_REVERSE,
    is_date_format,
    is_timedelta_format,
    builtin_format_code
)
from .named_styles import (
    _NamedCellStyleList,
    NamedStyleList,
    NamedStyle,
)
from .cell_style import CellStyle, CellStyleList


class Stylesheet(Serialisable):

    tagname = "styleSheet"

    numFmts = Typed(expected_type=NumberFormatList)
    fonts = NestedSequence(expected_type=Font, count=True)
    fills = NestedSequence(expected_type=Fill, count=True)
    borders = NestedSequence(expected_type=Border, count=True)
    cellStyleXfs = Typed(expected_type=CellStyleList)
    cellXfs = Typed(expected_type=CellStyleList)
    cellStyles = Typed(expected_type=_NamedCellStyleList)
    dxfs = NestedSequence(expected_type=DifferentialStyle, count=True)
    tableStyles = Typed(expected_type=TableStyleList, allow_none=True)
    colors = Typed(expected_type=ColorList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs',
                    'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors')

    def __init__(self,
                 numFmts=None,
                 fonts=(),
                 fills=(),
                 borders=(),
                 cellStyleXfs=None,
                 cellXfs=None,
                 cellStyles=None,
                 dxfs=(),
                 tableStyles=None,
                 colors=None,
                 extLst=None,
                ):
        if numFmts is None:
            numFmts = NumberFormatList()
        self.numFmts = numFmts
        self.number_formats = IndexedList()
        self.fonts = fonts
        self.fills = fills
        self.borders = borders
        if cellStyleXfs is None:
            cellStyleXfs = CellStyleList()
        self.cellStyleXfs = cellStyleXfs
        if cellXfs is None:
            cellXfs = CellStyleList()
        self.cellXfs = cellXfs
        if cellStyles is None:
            cellStyles = _NamedCellStyleList()
        self.cellStyles = cellStyles

        self.dxfs = dxfs
        self.tableStyles = tableStyles
        self.colors = colors

        self.cell_styles = self.cellXfs._to_array()
        self.alignments = self.cellXfs.alignments
        self.protections = self.cellXfs.prots
        self._normalise_numbers()
        self.named_styles = self._merge_named_styles()


    @classmethod
    def from_tree(cls, node):
        # strip all attribs
        attrs = dict(node.attrib)
        for k in attrs:
            del node.attrib[k]
        return super().from_tree(node)


    def _merge_named_styles(self):
        """
        Merge named style names "cellStyles" with their associated styles
        "cellStyleXfs"
        """
        style_refs = self.cellStyles.remove_duplicates()
        from_ref = [self._expand_named_style(style_ref) for style_ref in style_refs]

        return NamedStyleList(from_ref)


    def _expand_named_style(self, style_ref):
        """
        Expand a named style reference element to a
        named style object by binding the relevant
        objects from the stylesheet
        """
        xf = self.cellStyleXfs[style_ref.xfId]
        named_style = NamedStyle(
            name=style_ref.name,
            hidden=style_ref.hidden,
            builtinId=style_ref.builtinId,
        )

        named_style.font = self.fonts[xf.fontId]
        named_style.fill = self.fills[xf.fillId]
        named_style.border = self.borders[xf.borderId]
        if xf.numFmtId < BUILTIN_FORMATS_MAX_SIZE:
            formats = BUILTIN_FORMATS
        else:
            formats = self.custom_formats

        if xf.numFmtId in formats:
            named_style.number_format = formats[xf.numFmtId]
        if xf.alignment:
            named_style.alignment = xf.alignment
        if xf.protection:
            named_style.protection = xf.protection

        return named_style


    def _split_named_styles(self, wb):
        """
        Convert NamedStyle into separate CellStyle and Xf objects

        """
        for  style in wb._named_styles:
            self.cellStyles.cellStyle.append(style.as_name())
            self.cellStyleXfs.xf.append(style.as_xf())


    @property
    def custom_formats(self):
        return dict([(n.numFmtId, n.formatCode) for n in self.numFmts.numFmt])


    def _normalise_numbers(self):
        """
        Rebase custom numFmtIds with a floor of 164 when reading stylesheet
        And index datetime formats
        """
        date_formats = set()
        timedelta_formats = set()
        custom = self.custom_formats
        formats = self.number_formats
        for idx, style in enumerate(self.cell_styles):
            if style.numFmtId in custom:
                fmt = custom[style.numFmtId]
                if fmt in BUILTIN_FORMATS_REVERSE: # remove builtins
                    style.numFmtId = BUILTIN_FORMATS_REVERSE[fmt]
                else:
                    style.numFmtId = formats.add(fmt) + BUILTIN_FORMATS_MAX_SIZE
            else:
                fmt = builtin_format_code(style.numFmtId)
            if is_date_format(fmt):
                # Create an index of which styles refer to datetimes
                date_formats.add(idx)
            if is_timedelta_format(fmt):
                # Create an index of which styles refer to timedeltas
                timedelta_formats.add(idx)
        self.date_formats = date_formats
        self.timedelta_formats = timedelta_formats


    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


def apply_stylesheet(archive, wb):
    """
    Add styles to workbook if present
    """
    try:
        src = archive.read(ARC_STYLE)
    except KeyError:
        return wb

    node = fromstring(src)
    stylesheet = Stylesheet.from_tree(node)

    if stylesheet.cell_styles:

        wb._borders = IndexedList(stylesheet.borders)
        wb._fonts = IndexedList(stylesheet.fonts)
        wb._fills = IndexedList(stylesheet.fills)
        wb._differential_styles.styles = stylesheet.dxfs
        wb._number_formats = stylesheet.number_formats
        wb._protections = stylesheet.protections
        wb._alignments = stylesheet.alignments
        wb._table_styles = stylesheet.tableStyles

        # need to overwrite openpyxl defaults in case workbook has different ones
        wb._cell_styles = stylesheet.cell_styles
        wb._named_styles = stylesheet.named_styles
        wb._date_formats = stylesheet.date_formats
        wb._timedelta_formats = stylesheet.timedelta_formats

        for ns in wb._named_styles:
            ns.bind(wb)

    else:
        warn("Workbook contains no stylesheet, using openpyxl's defaults")

    if not wb._named_styles:
        normal = styles['Normal']
        wb.add_named_style(normal)
        warn("Workbook contains no default style, apply openpyxl's default")

    if stylesheet.colors is not None:
        wb._colors = stylesheet.colors.index


def write_stylesheet(wb):
    stylesheet = Stylesheet()
    stylesheet.fonts = wb._fonts
    stylesheet.fills = wb._fills
    stylesheet.borders = wb._borders
    stylesheet.dxfs = wb._differential_styles.styles
    stylesheet.colors = ColorList(indexedColors=wb._colors)

    from .numbers import NumberFormat
    fmts = []
    for idx, code in enumerate(wb._number_formats, BUILTIN_FORMATS_MAX_SIZE):
        fmt = NumberFormat(idx, code)
        fmts.append(fmt)

    stylesheet.numFmts.numFmt = fmts

    xfs = []
    for style in wb._cell_styles:
        xf = CellStyle.from_array(style)

        if style.alignmentId:
            xf.alignment = wb._alignments[style.alignmentId]

        if style.protectionId:
            xf.protection = wb._protections[style.protectionId]
        xfs.append(xf)
    stylesheet.cellXfs = CellStyleList(xf=xfs)

    stylesheet._split_named_styles(wb)
    stylesheet.tableStyles = wb._table_styles

    return stylesheet.to_tree()

''',
    },
    'openpyxl.styles.table': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Bool,
    Set,
    Integer,
    NoneSet,
    String,
    Sequence
)

from .colors import Color


class TableStyleElement(Serialisable):

    tagname = "tableStyleElement"

    type = Set(values=(['wholeTable', 'headerRow', 'totalRow', 'firstColumn',
                        'lastColumn', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe',
                        'secondColumnStripe', 'firstHeaderCell', 'lastHeaderCell',
                        'firstTotalCell', 'lastTotalCell', 'firstSubtotalColumn',
                        'secondSubtotalColumn', 'thirdSubtotalColumn', 'firstSubtotalRow',
                        'secondSubtotalRow', 'thirdSubtotalRow', 'blankRow',
                        'firstColumnSubheading', 'secondColumnSubheading',
                        'thirdColumnSubheading', 'firstRowSubheading', 'secondRowSubheading',
                        'thirdRowSubheading', 'pageFieldLabels', 'pageFieldValues']))
    size = Integer(allow_none=True)
    dxfId = Integer(allow_none=True)

    def __init__(self,
                 type=None,
                 size=None,
                 dxfId=None,
                ):
        self.type = type
        self.size = size
        self.dxfId = dxfId


class TableStyle(Serialisable):

    tagname = "tableStyle"

    name = String()
    pivot = Bool(allow_none=True)
    table = Bool(allow_none=True)
    count = Integer(allow_none=True)
    tableStyleElement = Sequence(expected_type=TableStyleElement, allow_none=True)

    __elements__ = ('tableStyleElement',)

    def __init__(self,
                 name=None,
                 pivot=None,
                 table=None,
                 count=None,
                 tableStyleElement=(),
                ):
        self.name = name
        self.pivot = pivot
        self.table = table
        self.count = count
        self.tableStyleElement = tableStyleElement


class TableStyleList(Serialisable):

    tagname = "tableStyles"

    defaultTableStyle = String(allow_none=True)
    defaultPivotStyle = String(allow_none=True)
    tableStyle = Sequence(expected_type=TableStyle, allow_none=True)

    __elements__ = ('tableStyle',)
    __attrs__ = ("count", "defaultTableStyle", "defaultPivotStyle")

    def __init__(self,
                 count=None,
                 defaultTableStyle="TableStyleMedium9",
                 defaultPivotStyle="PivotStyleLight16",
                 tableStyle=(),
                ):
        self.defaultTableStyle = defaultTableStyle
        self.defaultPivotStyle = defaultPivotStyle
        self.tableStyle = tableStyle


    @property
    def count(self):
        return len(self.tableStyle)

""",
    },
    'openpyxl.utils': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from .cell import (
    absolute_coordinate,
    cols_from_range,
    column_index_from_string,
    coordinate_to_tuple,
    get_column_letter,
    get_column_interval,
    quote_sheetname,
    range_boundaries,
    range_to_tuple,
    rows_from_range,
)

from .formulas import FORMULAE

""",
    },
    'openpyxl.utils.bound_dictionary': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from collections import defaultdict


class BoundDictionary(defaultdict):
    """
    A default dictionary where elements are tightly coupled.

    The factory method is responsible for binding the parent object to the child.

    If a reference attribute is assigned then child objects will have the key assigned to this.

    Otherwise it's just a defaultdict.
    """

    def __init__(self, reference=None, *args, **kw):
        self.reference = reference
        super().__init__(*args, **kw)


    def __getitem__(self, key):
        value = super().__getitem__(key)
        if self.reference is not None:
            setattr(value, self.reference, key)
        return value

''',
    },
    'openpyxl.utils.cell': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Collection of utilities used within the package and also available for client code
"""
from functools import lru_cache
from itertools import chain, product
from string import ascii_uppercase, digits
import re

from .exceptions import CellCoordinatesException

# constants
COORD_RE = re.compile(r'^[$]?([A-Za-z]{1,3})[$]?(\d+)$')
COL_RANGE = """[A-Z]{1,3}:[A-Z]{1,3}:"""
ROW_RANGE = r"""\d+:\d+:"""
RANGE_EXPR = r"""
[$]?(?P<min_col>[A-Za-z]{1,3})?
[$]?(?P<min_row>\d+)?
(:[$]?(?P<max_col>[A-Za-z]{1,3})?
[$]?(?P<max_row>\d+)?)?
"""
ABSOLUTE_RE = re.compile('^' + RANGE_EXPR +'$', re.VERBOSE)
SHEET_TITLE = r"""
(('(?P<quoted>([^']|'')*)')|(?P<notquoted>[^'^ ^!]*))!"""
SHEETRANGE_RE = re.compile("""{0}(?P<cells>{1})(?=,?)""".format(
    SHEET_TITLE, RANGE_EXPR), re.VERBOSE)


def get_column_interval(start, end):
    """
    Given the start and end columns, return all the columns in the series.

    The start and end columns can be either column letters or 1-based
    indexes.
    """
    if isinstance(start, str):
        start = column_index_from_string(start)
    if isinstance(end, str):
        end = column_index_from_string(end)
    return [get_column_letter(x) for x in range(start, end + 1)]


def coordinate_from_string(coord_string):
    """Convert a coordinate string like 'B12' to a tuple ('B', 12)"""
    match = COORD_RE.match(coord_string)
    if not match:
        msg = f"Invalid cell coordinates ({coord_string})"
        raise CellCoordinatesException(msg)
    column, row = match.groups()
    row = int(row)
    if not row:
        msg = f"There is no row 0 ({coord_string})"
        raise CellCoordinatesException(msg)
    return column, row


def absolute_coordinate(coord_string):
    """Convert a coordinate to an absolute coordinate string (B12 -> $B$12)"""
    m = ABSOLUTE_RE.match(coord_string)
    if not m:
        raise ValueError(f"{coord_string} is not a valid coordinate range")

    d = m.groupdict('')
    for k, v in d.items():
        if v:
            d[k] = f"${v}"

    if d['max_col'] or d['max_row']:
        fmt = "{min_col}{min_row}:{max_col}{max_row}"
    else:
        fmt = "{min_col}{min_row}"
    return fmt.format(**d)


__decimal_to_alpha = [""] + list(ascii_uppercase)

@lru_cache(maxsize=None)
def get_column_letter(col_idx):
    """
    Convert decimal column position to its ASCII (base 26) form.

    Because column indices are 1-based, strides are actually pow(26, n) + 26
    Hence, a correction is applied between pow(26, n) and pow(26, 2) + 26 to
    prevent and additional column letter being prepended

    "A" == 1 == pow(26, 0)
    "Z" == 26 == pow(26, 0) + 26 // decimal equivalent 10
    "AA" == 27 == pow(26, 1) + 1
    "ZZ" == 702 == pow(26, 2) + 26 // decimal equivalent 100
    """

    if not 1 <= col_idx <= 18278:
        raise ValueError("Invalid column index {0}".format(col_idx))

    result = []

    if col_idx < 26:
        return __decimal_to_alpha[col_idx]

    while col_idx:
        col_idx, remainder = divmod(col_idx, 26)
        result.insert(0, __decimal_to_alpha[remainder])
        if not remainder:
            col_idx -= 1
            result.insert(0, "Z")

    return "".join(result)


__alpha_to_decimal = {letter:pos for pos, letter in enumerate(ascii_uppercase, 1)}
__powers = (1, 26, 676)

@lru_cache(maxsize=None)
def column_index_from_string(col):
    """
    Convert ASCII column name (base 26) to decimal with 1-based index

    Characters represent descending multiples of powers of 26

    "AFZ" == 26 * pow(26, 0) + 6 * pow(26, 1) + 1 * pow(26, 2)
    """
    error_msg = f"'{col}' is not a valid column name. Column names are from A to ZZZ"
    if len(col) > 3:
        raise ValueError(error_msg)
    idx = 0
    col = reversed(col.upper())
    for letter, power in zip(col, __powers):
        try:
            pos = __alpha_to_decimal[letter]
        except KeyError:
            raise ValueError(error_msg)
        idx += pos * power
    if not 0 < idx < 18279:
        raise ValueError(error_msg)
    return idx


def range_boundaries(range_string):
    """
    Convert a range string into a tuple of boundaries:
    (min_col, min_row, max_col, max_row)
    Cell coordinates will be converted into a range with the cell at both end
    """
    msg = "{0} is not a valid coordinate or range".format(range_string)
    m = ABSOLUTE_RE.match(range_string)
    if not m:
        raise ValueError(msg)

    min_col, min_row, sep, max_col, max_row = m.groups()

    if sep:
        cols = min_col, max_col
        rows = min_row, max_row

        if not (
            all(cols + rows) or
            all(cols) and not any(rows) or
            all(rows) and not any(cols)
        ):
            raise ValueError(msg)

    if min_col is not None:
        min_col = column_index_from_string(min_col)

    if min_row is not None:
        min_row = int(min_row)

    if max_col is not None:
        max_col = column_index_from_string(max_col)
    else:
        max_col = min_col

    if max_row is not None:
        max_row = int(max_row)
    else:
        max_row = min_row

    return min_col, min_row, max_col, max_row


def rows_from_range(range_string):
    """
    Get individual addresses for every cell in a range.
    Yields one row at a time.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    rows = range(min_row, max_row + 1)
    cols = [get_column_letter(col) for col in range(min_col, max_col + 1)]
    for row in rows:
        yield tuple('{0}{1}'.format(col, row) for col in cols)


def cols_from_range(range_string):
    """
    Get individual addresses for every cell in a range.
    Yields one row at a time.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_string)
    rows = range(min_row, max_row+1)
    cols = (get_column_letter(col) for col in range(min_col, max_col+1))
    for col in cols:
        yield tuple('{0}{1}'.format(col, row) for row in rows)


def coordinate_to_tuple(coordinate):
    """
    Convert an Excel style coordinate to (row, column) tuple
    """
    for idx, c in enumerate(coordinate):
        if c in digits:
            break
    col = coordinate[:idx]
    row = coordinate[idx:]
    return int(row), column_index_from_string(col)


def range_to_tuple(range_string):
    """
    Convert a worksheet range to the sheetname and maximum and minimum
    coordinate indices
    """
    m = SHEETRANGE_RE.match(range_string)
    if m is None:
        raise ValueError("Value must be of the form sheetname!A1:E4")
    sheetname = m.group("quoted") or m.group("notquoted")
    cells = m.group("cells")
    boundaries = range_boundaries(cells)
    return sheetname, boundaries


def quote_sheetname(sheetname):
    """
    Add quotes around sheetnames if they contain spaces.
    """
    if "'" in sheetname:
        sheetname = sheetname.replace("'", "''")

    sheetname = u"'{0}'".format(sheetname)
    return sheetname

''',
    },
    'openpyxl.utils.dataframe': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from itertools import accumulate
import operator
import numpy
from openpyxl.compat.product import prod


def dataframe_to_rows(df, index=True, header=True):
    """
    Convert a Pandas dataframe into something suitable for passing into a worksheet.
    If index is True then the index will be included, starting one row below the header.
    If header is True then column headers will be included starting one column to the right.
    Formatting should be done by client code.
    """
    from pandas import Timestamp

    if header:
        if df.columns.nlevels > 1:
            rows = expand_index(df.columns, header)
        else:
            rows = [list(df.columns.values)]
        for row in rows:
            n = []
            for v in row:
                if isinstance(v, numpy.datetime64):
                    v = Timestamp(v)
                n.append(v)
            row = n
            if index:
                row = [None]*df.index.nlevels + row
            yield row

    if index:
        yield df.index.names

    expanded = ([v] for v in df.index)
    if df.index.nlevels > 1:
        expanded = expand_index(df.index)

    # Using the expanded index is preferable to df.itertuples(index=True) so that we have 'None' inserted where applicable
    for (df_index, row) in zip(expanded, df.itertuples(index=False)):
        row = list(row)
        if index:
            row = df_index + row
        yield row


def expand_index(index, header=False):
    """
    Expand axis or column Multiindex
    For columns use header = True
    For axes use header = False (default)
    """

    # For each element of the index, zip the members with the previous row
    # If the 2 elements of the zipped list do not match, we can insert the new value into the row
    # or if an earlier member was different, all later members should be added to the row
    values = list(index.values)
    previous_value = [None] * len(values[0])
    result = []

    for value in values:
        row = [None] * len(value)

        # Once there's a difference in member of an index with the prior index, we need to store all subsequent members in the row
        prior_change = False
        for idx, (current_index_member, previous_index_member) in enumerate(zip(value, previous_value)):

            if current_index_member != previous_index_member or prior_change:
                row[idx] = current_index_member
                prior_change = True

        previous_value = value

        # If this is for a row index, we're already returning a row so just yield
        if not header:
            yield row
        else:
            result.append(row)

    # If it's for a header, we need to transpose to get it in row order
    # Example: result = [['A', 'A'], [None, 'B']] -> [['A', None], ['A', 'B']]
    if header:
        result = numpy.array(result).transpose().tolist()
        for row in result:
            yield row

''',
    },
    'openpyxl.utils.datetime': {
        'is_package': False,
        'source': '# Copyright (c) 2010-2024 openpyxl\n\n"""Manage Excel date weirdness."""\n\n# Python stdlib imports\nimport datetime\nfrom math import isnan\nimport re\n\n\n# constants\nMAC_EPOCH = datetime.datetime(1904, 1, 1)\nWINDOWS_EPOCH = datetime.datetime(1899, 12, 30)\nCALENDAR_WINDOWS_1900 = 2415018.5   # Julian date of WINDOWS_EPOCH\nCALENDAR_MAC_1904 = 2416480.5       # Julian date of MAC_EPOCH\nCALENDAR_WINDOWS_1900 = WINDOWS_EPOCH\nCALENDAR_MAC_1904 = MAC_EPOCH\nSECS_PER_DAY = 86400\n\nISO_FORMAT = \'%Y-%m-%dT%H:%M:%SZ\'\nISO_REGEX = re.compile(r\'\'\'\n(?P<date>(?P<year>\\d{4})-(?P<month>\\d{2})-(?P<day>\\d{2}))?T?\n(?P<time>(?P<hour>\\d{2}):(?P<minute>\\d{2})(:(?P<second>\\d{2})(?P<microsecond>\\.\\d{1,3})?)?)?Z?\'\'\',\n                                       re.VERBOSE)\nISO_DURATION = re.compile(r\'PT((?P<hours>\\d+)H)?((?P<minutes>\\d+)M)?((?P<seconds>\\d+(\\.\\d{1,3})?)S)?\')\n\n\ndef to_ISO8601(dt):\n    """Convert from a datetime to a timestamp string."""\n    if hasattr(dt, "microsecond") and dt.microsecond:\n        return dt.isoformat(timespec="milliseconds")\n    return dt.isoformat()\n\n\ndef from_ISO8601(formatted_string):\n    """Convert from a timestamp string to a datetime object. According to\n    18.17.4 in the specification the following ISO 8601 formats are\n    supported.\n\n    Dates B.1.1 and B.2.1\n    Times B.1.2 and B.2.2\n    Datetimes B.1.3 and B.2.3\n\n    There is no concept of timedeltas in the specification, but Excel\n    writes them (in strict OOXML mode), so these are also understood.\n    """\n    if not formatted_string:\n        return None\n\n    match = ISO_REGEX.match(formatted_string)\n    if match and any(match.groups()):\n        parts = match.groupdict(0)\n        for key in ["year", "month", "day", "hour", "minute", "second"]:\n            if parts[key]:\n                parts[key] = int(parts[key])\n\n        if parts["microsecond"]:\n            parts["microsecond"] = int(float(parts[\'microsecond\']) * 1_000_000)\n\n        if not parts["date"]:\n            dt = datetime.time(parts[\'hour\'], parts[\'minute\'], parts[\'second\'], parts["microsecond"])\n        elif not parts["time"]:\n            dt = datetime.date(parts[\'year\'], parts[\'month\'], parts[\'day\'])\n        else:\n            del parts["time"]\n            del parts["date"]\n            dt = datetime.datetime(**parts)\n        return dt\n\n    match = ISO_DURATION.match(formatted_string)\n    if match and any(match.groups()):\n        parts = match.groupdict(0)\n        for key, val in parts.items():\n            if val:\n                parts[key] = float(val)\n        return datetime.timedelta(**parts)\n\n    raise ValueError("Invalid datetime value {}".format(formatted_string))\n\n\ndef to_excel(dt, epoch=WINDOWS_EPOCH):\n    """Convert Python datetime to Excel serial"""\n    if isinstance(dt, datetime.time):\n        return time_to_days(dt)\n    if isinstance(dt, datetime.timedelta):\n        return timedelta_to_days(dt)\n    if isnan(dt.year):  # Pandas supports Not a Date\n        return\n\n    if not hasattr(dt, "date"):\n        dt = datetime.datetime.combine(dt, datetime.time())\n\n    # rebase on epoch and adjust for < 1900-03-01\n    days = (dt - epoch).days\n    if 0 < days <= 60 and epoch == WINDOWS_EPOCH:\n        days -= 1\n    return days + time_to_days(dt)\n\n\ndef from_excel(value, epoch=WINDOWS_EPOCH, timedelta=False):\n    """Convert Excel serial to Python datetime"""\n    if value is None:\n        return\n\n    if timedelta:\n        td = datetime.timedelta(days=value)\n        if td.microseconds:\n            # round to millisecond precision\n            td = datetime.timedelta(seconds=td.total_seconds() // 1,\n                                    microseconds=round(td.microseconds, -3))\n        return td\n\n    day, fraction = divmod(value, 1)\n    diff = datetime.timedelta(milliseconds=round(fraction * SECS_PER_DAY * 1000))\n    if 0 <= value < 1 and diff.days == 0:\n        return days_to_time(diff)\n    if 0 < value < 60 and epoch == WINDOWS_EPOCH:\n        day += 1\n    return epoch + datetime.timedelta(days=day) + diff\n\n\ndef time_to_days(value):\n    """Convert a time value to fractions of day"""\n    return (\n        (value.hour * 3600)\n        + (value.minute * 60)\n        + value.second\n        + value.microsecond / 10**6\n        ) / SECS_PER_DAY\n\n\ndef timedelta_to_days(value):\n    """Convert a timedelta value to fractions of a day"""\n    return value.total_seconds() / SECS_PER_DAY\n\n\ndef days_to_time(value):\n    mins, seconds = divmod(value.seconds, 60)\n    hours, mins = divmod(mins, 60)\n    return datetime.time(hours, mins, seconds, value.microseconds)\n',
    },
    'openpyxl.utils.escape': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
OOXML has non-standard escaping for characters < \031
"""

import re


def escape(value):
    r"""
    Convert ASCII < 31 to OOXML: \n == _x + hex(ord(\n)) + _
    """

    CHAR_REGEX = re.compile(r"[\001-\031]")

    def _sub(match):
        """
        Callback to escape chars
        """
        return "_x{:0>4x}_".format(ord(match.group(0)))

    return CHAR_REGEX.sub(_sub, value)


def unescape(value):
    r"""
    Convert escaped strings to ASCIII: _x000a_ == \n
    """


    ESCAPED_REGEX = re.compile("_x([0-9A-Fa-f]{4})_")

    def _sub(match):
        """
        Callback to unescape chars
        """
        return chr(int(match.group(1), 16))

    if "_x" in value:
        value = ESCAPED_REGEX.sub(_sub, value)

    return value

''',
    },
    'openpyxl.utils.exceptions': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


"""Definitions for openpyxl shared exception classes."""


class CellCoordinatesException(Exception):
    """Error for converting between numeric and A1-style cell references."""


class IllegalCharacterError(Exception):
    """The data submitted which cannot be used directly in Excel files. It
    must be removed or escaped."""


class NamedRangeException(Exception):
    """Error for badly formatted named ranges."""


class SheetTitleException(Exception):
    """Error for bad sheet names."""


class InvalidFileException(Exception):
    """Error for trying to open a non-ooxml file."""


class ReadOnlyWorkbookException(Exception):
    """Error for trying to modify a read-only workbook"""


class WorkbookAlreadySaved(Exception):
    """Error when attempting to perform operations on a dump workbook
    while it has already been dumped once"""

''',
    },
    'openpyxl.utils.formulas': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
List of builtin formulae
"""

FORMULAE = ("CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY", "CUBERANKEDMEMBER", "CUBESET", "CUBESETCOUNT", "CUBEVALUE", "DAVERAGE", "DCOUNT", "DCOUNTA", "DGET", "DMAX", "DMIN", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DVAR", "DVARP", "DATE", "DATEDIF", "DATEVALUE", "DAY", "DAYS360", "EDATE", "EOMONTH", "HOUR", "MINUTE", "MONTH", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOW", "SECOND", "TIME", "TIMEVALUE", "TODAY", "WEEKDAY", "WEEKNUM", "WORKDAY", "WORKDAY.INTL", "YEAR", "YEARFRAC", "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BIN2DEC", "BIN2HEX", "BIN2OCT", "COMPLEX", "CONVERT", "DEC2BIN", "DEC2HEX", "DEC2OCT", "DELTA", "ERF", "ERFC", "GESTEP", "HEX2BIN", "HEX2DEC", "HEX2OCT", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS", "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSIN", "IMSQRT", "IMSUB", "IMSUM", "OCT2BIN", "OCT2DEC", "OCT2HEX", "ACCRINT", "ACCRINTM", "AMORDEGRC", "AMORLINC", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM", "COUPPCD", "CUMIPMT", "CUMPRINC", "DB", "DDB", "DISC", "DOLLARDE", "DOLLARFR", "DURATION", "EFFECT", "FV", "FVSCHEDULE", "INTRATE", "IPMT", "IRR", "ISPMT", "MDURATION", "MIRR", "NOMINAL", "NPER", "NPV", "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD", "PMT", "PPMT", "PRICE", "PRICEDISC", "PRICEMAT", "PV", "RATE", "RECEIVED", "SLN", "SYD", "TBILLEQ", "TBILLPRICE", "TBILLYIELD", "VDB", "XIRR", "XNPV", "YIELD", "YIELDDISC", "YIELDMAT", "CELL", "ERROR.TYPE", "INFO", "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER", "ISODD", "ISREF", "ISTEXT", "N", "NA", "TYPE", "AND", "FALSE", "IF", "IFERROR", "NOT", "OR", "TRUE", "ADDRESS", "AREAS", "CHOOSE", "COLUMN", "COLUMNS", "GETPIVOTDATA", "HLOOKUP", "HYPERLINK", "INDEX", "INDIRECT", "LOOKUP", "MATCH", "OFFSET", "ROW", "ROWS", "RTD", "TRANSPOSE", "VLOOKUP", "ABS", "ACOS", "ACOSH", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "CEILING", "COMBIN", "COS", "COSH", "DEGREES", "ECMA.CEILING", "EVEN", "EXP", "FACT", "FACTDOUBLE", "FLOOR", "GCD", "INT", "ISO.CEILING", "LCM", "LN", "LOG", "LOG10", "MDETERM", "MINVERSE", "MMULT", "MOD", "MROUND", "MULTINOMIAL", "ODD", "PI", "POWER", "PRODUCT", "QUOTIENT", "RADIANS", "RAND", "RANDBETWEEN", "ROMAN", "ROUND", "ROUNDDOWN", "ROUNDUP", "SERIESSUM", "SIGN", "SIN", "SINH", "SQRT", "SQRTPI", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ", "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "TAN", "TANH", "TRUNC", "AVEDEV", "AVERAGE", "AVERAGEA", "AVERAGEIF", "AVERAGEIFS", "BETADIST", "BETAINV", "BINOMDIST", "CHIDIST", "CHIINV", "CHITEST", "CONFIDENCE", "CORREL", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "COVAR", "CRITBINOM", "DEVSQ", "EXPONDIST", "FDIST", "FINV", "FISHER", "FISHERINV", "FORECAST", "FREQUENCY", "FTEST", "GAMMADIST", "GAMMAINV", "GAMMALN", "GEOMEAN", "GROWTH", "HARMEAN", "HYPGEOMDIST", "INTERCEPT", "KURT", "LARGE", "LINEST", "LOGEST", "LOGINV", "LOGNORMDIST", "MAX", "MAXA", "MEDIAN", "MIN", "MINA", "MODE", "NEGBINOMDIST", "NORMDIST", "NORMINV", "NORMSDIST", "NORMSINV", "PEARSON", "PERCENTILE", "PERCENTRANK", "PERMUT", "POISSON", "PROB", "QUARTILE", "RANK", "RSQ", "SKEW", "SLOPE", "SMALL", "STANDARDIZE", "STDEV", "STDEVA", "STDEVP", "STDEVPA", "STEYX", "TDIST", "TINV", "TREND", "TRIMMEAN", "TTEST", "VAR", "VARA", "VARP", "VARPA", "WEIBULL", "ZTEST", "ASC", "BAHTTEXT", "CHAR", "CLEAN", "CODE", "CONCATENATE", "DOLLAR", "EXACT", "FIND", "FINDB", "FIXED", "JIS", "LEFT", "LEFTB", "LEN", "LENB", "LOWER", "MID", "MIDB", "PHONETIC", "PROPER", "REPLACE", "REPLACEB", "REPT", "RIGHT", "RIGHTB", "SEARCH", "SEARCHB", "SUBSTITUTE", "T", "TEXT", "TRIM", "UPPER", "VALUE")

FORMULAE = frozenset(FORMULAE)


from openpyxl.formula import Tokenizer


def validate(formula):
    """
    Utility function for checking whether a formula is syntactically correct
    """
    assert formula.startswith("=")
    formula = Tokenizer(formula)
    for t in formula.items:
        if t.type == "FUNC" and t.subtype == "OPEN":
            if not t.value.startswith("_xlfn.") and t.value[:-1] not in FORMULAE:
                raise ValueError(f"Unknown function {t.value} in {formula.formula}. The function may need a prefix")

''',
    },
    'openpyxl.utils.indexed_list': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


class IndexedList(list):
    """
    List with optimised access by value
    Based on Alex Martelli's recipe

    http://code.activestate.com/recipes/52303-the-auxiliary-dictionary-idiom-for-sequences-with-/
    """

    _dict = {}

    def __init__(self, iterable=None):
        self.clean = True
        self._dict = {}
        if iterable is not None:
            self.clean = False
            for idx, val in enumerate(iterable):
                self._dict[val] = idx
                list.append(self, val)

    def _rebuild_dict(self):
        self._dict = {}
        idx = 0
        for value in self:
            if value not in self._dict:
                self._dict[value] = idx
                idx += 1
        self.clean = True

    def __contains__(self, value):
        if not self.clean:
            self._rebuild_dict()
        return value in self._dict

    def index(self, value):
        if value in self:
            return self._dict[value]
        raise ValueError

    def append(self, value):
        if value not in self._dict:
            self._dict[value] = len(self)
            list.append(self, value)

    def add(self, value):
        self.append(value)
        return self._dict[value]

''',
    },
    'openpyxl.utils.inference': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
Type inference functions
"""
import datetime
import re

from openpyxl.styles import numbers

PERCENT_REGEX = re.compile(r'^(?P<number>\-?[0-9]*\.?[0-9]*\s?)\%$')
TIME_REGEX = re.compile(r"""
^(?: # HH:MM and HH:MM:SS
(?P<hour>[0-1]{0,1}[0-9]{2}):
(?P<minute>[0-5][0-9]):?
(?P<second>[0-5][0-9])?$)
|
^(?: # MM:SS.
([0-5][0-9]):
([0-5][0-9])?\.
(?P<microsecond>\d{1,6}))
""", re.VERBOSE)
NUMBER_REGEX = re.compile(r'^-?([\d]|[\d]+\.[\d]*|\.[\d]+|[1-9][\d]+\.?[\d]*)((E|e)[-+]?[\d]+)?$')


def cast_numeric(value):
    """Explicitly convert a string to a numeric value"""
    if NUMBER_REGEX.match(value):
        try:
            return int(value)
        except ValueError:
            return float(value)


def cast_percentage(value):
    """Explicitly convert a string to numeric value and format as a
    percentage"""
    match = PERCENT_REGEX.match(value)
    if match:
        return float(match.group('number')) / 100



def cast_time(value):
    """Explicitly convert a string to a number and format as datetime or
    time"""
    match = TIME_REGEX.match(value)
    if match:
        if match.group("microsecond") is not None:
            value = value[:12]
            pattern = "%M:%S.%f"
            #fmt = numbers.FORMAT_DATE_TIME5
        elif match.group('second') is None:
            #fmt = numbers.FORMAT_DATE_TIME3
            pattern = "%H:%M"
        else:
            pattern = "%H:%M:%S"
            #fmt = numbers.FORMAT_DATE_TIME6
        value = datetime.datetime.strptime(value, pattern)
        return value.time()

''',
    },
    'openpyxl.utils.protection': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


def hash_password(plaintext_password=''):
    """
    Create a password hash from a given string for protecting a worksheet
    only. This will not work for encrypting a workbook.

    This method is based on the algorithm provided by
    Daniel Rentz of OpenOffice and the PEAR package
    Spreadsheet_Excel_Writer by Xavier Noguer <xnoguer@rezebra.com>.
    See also http://blogs.msdn.com/b/ericwhite/archive/2008/02/23/the-legacy-hashing-algorithm-in-open-xml.aspx
    """
    password = 0x0000
    for idx, char in enumerate(plaintext_password, 1):
        value = ord(char) << idx
        rotated_bits = value >> 15
        value &= 0x7fff
        password ^= (value | rotated_bits)
    password ^= len(plaintext_password)
    password ^= 0xCE4B
    return str(hex(password)).upper()[2:]

''',
    },
    'openpyxl.utils.units': {
        'is_package': False,
        'source': r'''

# Copyright (c) 2010-2024 openpyxl

import math


#constants

DEFAULT_ROW_HEIGHT = 15.  # Default row height measured in point size.
BASE_COL_WIDTH = 8 # in characters
DEFAULT_COLUMN_WIDTH = BASE_COL_WIDTH + 5
#  = baseColumnWidth + {margin padding (2 pixels on each side, totalling 4 pixels)} + {gridline (1pixel)}


DEFAULT_LEFT_MARGIN = 0.7 # in inches, = right margin
DEFAULT_TOP_MARGIN = 0.7874 # in inches = bottom margin
DEFAULT_HEADER = 0.3 # in inches


# Conversion functions
"""
From the ECMA Spec (4th Edition part 1)
Page setup: "Left Page Margin in inches" p. 1647

Docs from
http://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/

See also http://msdn.microsoft.com/en-us/library/dd560821(v=office.12).aspx

dxa: The main unit in OOXML is a twentieth of a point. Also called twips.
pt: point. In Excel there are 72 points to an inch
hp: half-points are used to specify font sizes. A font-size of 12pt equals 24 half points
pct: Half-points are used to specify font sizes. A font-size of 12pt equals 24 half points

EMU: English Metric Unit, EMUs are used for coordinates in vector-based
drawings and embedded pictures. One inch equates to 914400 EMUs and a
centimeter is 360000. For bitmaps the default resolution is 96 dpi (known as
PixelsPerInch in Excel). Spec p. 1122

For radial geometry Excel uses integer units of 1/60000th of a degree.
"""



def inch_to_dxa(value):
    """1 inch = 72 * 20 dxa"""
    return int(value * 20 * 72)

def dxa_to_inch(value):
    return value / 72 / 20


def dxa_to_cm(value):
    return 2.54 * dxa_to_inch(value)

def cm_to_dxa(value):
    emu = cm_to_EMU(value)
    inch = EMU_to_inch(emu)
    return inch_to_dxa(inch)


def pixels_to_EMU(value):
    """1 pixel = 9525 EMUs"""
    return int(value * 9525)

def EMU_to_pixels(value):
    return round(value / 9525)


def cm_to_EMU(value):
    """1 cm = 360000 EMUs"""
    return int(value * 360000)

def EMU_to_cm(value):
    return round(value / 360000, 4)


def inch_to_EMU(value):
    """1 inch = 914400 EMUs"""
    return int(value * 914400)

def EMU_to_inch(value):
    return round(value / 914400, 4)


def pixels_to_points(value, dpi=96):
    """96 dpi, 72i"""
    return value * 72 / dpi


def points_to_pixels(value, dpi=96):
    return int(math.ceil(value * dpi / 72))


def degrees_to_angle(value):
    """1 degree = 60000 angles"""
    return int(round(value * 60000))


def angle_to_degrees(value):
    return round(value / 60000, 2)


def short_color(color):
    """ format a color to its short size """
    if len(color) > 6:
        return color[2:]
    return color

''',
    },
    'openpyxl.workbook': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl


from .workbook import Workbook

""",
    },
    'openpyxl.workbook._writer': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Write the workbook global settings to the archive."""

from openpyxl.utils import quote_sheetname
from openpyxl.xml.constants import (
    ARC_APP,
    ARC_CORE,
    ARC_CUSTOM,
    ARC_WORKBOOK,
    PKG_REL_NS,
    CUSTOMUI_NS,
    ARC_ROOT_RELS,
)
from openpyxl.xml.functions import tostring, fromstring

from openpyxl.packaging.relationship import Relationship, RelationshipList
from openpyxl.workbook.defined_name import (
    DefinedName,
    DefinedNameList,
)
from openpyxl.workbook.external_reference import ExternalReference
from openpyxl.packaging.workbook import ChildSheet, WorkbookPackage, PivotCache
from openpyxl.workbook.properties import WorkbookProperties
from openpyxl.utils.datetime import CALENDAR_MAC_1904


def get_active_sheet(wb):
    """
    Return the index of the active sheet.
    If the sheet set to active is hidden return the next visible sheet or None
    """
    visible_sheets = [idx for idx, sheet in enumerate(wb._sheets) if sheet.sheet_state == "visible"]
    if not visible_sheets:
        raise IndexError("At least one sheet must be visible")

    idx = wb._active_sheet_index
    sheet = wb.active
    if sheet and sheet.sheet_state == "visible":
        return idx

    for idx in visible_sheets[idx:]:
        wb.active = idx
        return idx

    return None


class WorkbookWriter:

    def __init__(self, wb):
        self.wb = wb
        self.rels = RelationshipList()
        self.package = WorkbookPackage()
        self.package.workbookProtection = wb.security
        self.package.calcPr = wb.calculation


    def write_properties(self):

        props = WorkbookProperties() # needs a mapping to the workbook for preservation
        if self.wb.code_name is not None:
            props.codeName = self.wb.code_name
        if self.wb.excel_base_date == CALENDAR_MAC_1904:
            props.date1904 = True
        self.package.workbookPr = props


    def write_worksheets(self):
        for idx, sheet in enumerate(self.wb._sheets, 1):
            sheet_node = ChildSheet(name=sheet.title, sheetId=idx, id="rId{0}".format(idx))
            rel = Relationship(type=sheet._rel_type, Target=sheet.path)
            self.rels.append(rel)

            if not sheet.sheet_state == 'visible':
                if len(self.wb._sheets) == 1:
                    raise ValueError("The only worksheet of a workbook cannot be hidden")
                sheet_node.state = sheet.sheet_state
            self.package.sheets.append(sheet_node)


    def write_refs(self):
        for link in self.wb._external_links:
            # need to match a counter with a workbook's relations
            rId = len(self.wb.rels) + 1
            rel = Relationship(type=link._rel_type, Target=link.path)
            self.rels.append(rel)
            ext = ExternalReference(id=rel.id)
            self.package.externalReferences.append(ext)


    def write_names(self):
        defined_names = list(self.wb.defined_names.values())

        for idx, sheet in enumerate(self.wb.worksheets):
            quoted = quote_sheetname(sheet.title)

            # local names
            if sheet.defined_names:
                names = sheet.defined_names.values()
                for n in names:
                    n.localSheetId = idx
                defined_names.extend(names)

            if sheet.auto_filter:
                name = DefinedName(name='_FilterDatabase', localSheetId=idx, hidden=True)
                name.value = f"{quoted}!{sheet.auto_filter}"
                defined_names.append(name)

            if sheet.print_titles:
                name = DefinedName(name="Print_Titles", localSheetId=idx)
                name.value = sheet.print_titles
                defined_names.append(name)

            if sheet.print_area:
                name = DefinedName(name="Print_Area", localSheetId=idx)
                name.value = sheet.print_area
                defined_names.append(name)

        self.package.definedNames = DefinedNameList(definedName=defined_names)


    def write_pivots(self):
        pivot_caches = set()
        for pivot in self.wb._pivots:
            if pivot.cache not in pivot_caches:
                pivot_caches.add(pivot.cache)
                c = PivotCache(cacheId=pivot.cacheId)
                self.package.pivotCaches.append(c)
                rel = Relationship(Type=pivot.cache.rel_type, Target=pivot.cache.path)
                self.rels.append(rel)
                c.id = rel.id
        #self.wb._pivots = [] # reset


    def write_views(self):
        active = get_active_sheet(self.wb)
        if self.wb.views:
            self.wb.views[0].activeTab = active
        self.package.bookViews = self.wb.views


    def write(self):
        """Write the core workbook xml."""

        self.write_properties()
        self.write_worksheets()
        self.write_names()
        self.write_pivots()
        self.write_views()
        self.write_refs()

        return tostring(self.package.to_tree())


    def write_rels(self):
        """Write the workbook relationships xml."""

        styles =  Relationship(type='styles', Target='styles.xml')
        self.rels.append(styles)

        theme =  Relationship(type='theme', Target='theme/theme1.xml')
        self.rels.append(theme)

        if self.wb.vba_archive:
            vba =  Relationship(type='', Target='vbaProject.bin')
            vba.Type ='http://schemas.microsoft.com/office/2006/relationships/vbaProject'
            self.rels.append(vba)

        return tostring(self.rels.to_tree())


    def write_root_rels(self):
        """Write the package relationships"""

        rels = RelationshipList()

        rel = Relationship(type="officeDocument", Target=ARC_WORKBOOK)
        rels.append(rel)
        rel = Relationship(Type=f"{PKG_REL_NS}/metadata/core-properties", Target=ARC_CORE)
        rels.append(rel)

        rel = Relationship(type="extended-properties", Target=ARC_APP)
        rels.append(rel)

        if len(self.wb.custom_doc_props) >= 1:
            rel = Relationship(type="custom-properties", Target=ARC_CUSTOM)
            rels.append(rel)

        if self.wb.vba_archive is not None:
            # See if there was a customUI relation and reuse it
            xml = fromstring(self.wb.vba_archive.read(ARC_ROOT_RELS))
            root_rels = RelationshipList.from_tree(xml)
            for rel in root_rels.find(CUSTOMUI_NS):
                rels.append(rel)

        return tostring(rels.to_tree())

''',
    },
    'openpyxl.workbook.child': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import re
import warnings

from openpyxl.worksheet.header_footer import HeaderFooter

"""
Base class for worksheets, chartsheets, etc. that can be added to workbooks
"""

INVALID_TITLE_REGEX = re.compile(r'[\\*?:/\[\]]')


def avoid_duplicate_name(names, value):
    """
    Naive check to see whether name already exists.
    If name does exist suggest a name using an incrementer
    Duplicates are case insensitive
    """
    # Check for an absolute match in which case we need to find an alternative
    match = [n for n in names if n.lower() == value.lower()]
    if match:
        names = u",".join(names)
        sheet_title_regex = re.compile(f'(?P<title>{re.escape(value)})(?P<count>\\d*),?', re.I)
        matches = sheet_title_regex.findall(names)
        if matches:
            # use name, but append with the next highest integer
            counts = [int(idx) for (t, idx) in matches if idx.isdigit()]
            highest = 0
            if counts:
                highest = max(counts)
            value = u"{0}{1}".format(value, highest + 1)
    return value


class _WorkbookChild:

    __title = ""
    _id = None
    _path = "{0}"
    _parent = None
    _default_title = "Sheet"

    def __init__(self, parent=None, title=None):
        self._parent = parent
        self.title = title or self._default_title
        self.HeaderFooter = HeaderFooter()


    def __repr__(self):
        return '<{0} "{1}">'.format(self.__class__.__name__, self.title)


    @property
    def parent(self):
        return self._parent


    @property
    def encoding(self):
        return self._parent.encoding


    @property
    def title(self):
        return self.__title


    @title.setter
    def title(self, value):
        """
        Set a sheet title, ensuring it is valid.
        Limited to 31 characters, no special characters.
        Duplicate titles will be incremented numerically
        """
        if not self._parent:
            return

        if not value:
            raise ValueError("Title must have at least one character")

        if hasattr(value, "decode"):
            if not isinstance(value, str):
                try:
                    value = value.decode("ascii")
                except UnicodeDecodeError:
                    raise ValueError("Worksheet titles must be str")

        m = INVALID_TITLE_REGEX.search(value)
        if m:
            msg = "Invalid character {0} found in sheet title".format(m.group(0))
            raise ValueError(msg)

        if self.title is not None and self.title != value:
            value = avoid_duplicate_name(self.parent.sheetnames, value)

        if len(value) > 31:
            warnings.warn("Title is more than 31 characters. Some applications may not be able to read the file")

        self.__title = value


    @property
    def oddHeader(self):
        return self.HeaderFooter.oddHeader


    @oddHeader.setter
    def oddHeader(self, value):
        self.HeaderFooter.oddHeader = value


    @property
    def oddFooter(self):
        return self.HeaderFooter.oddFooter


    @oddFooter.setter
    def oddFooter(self, value):
        self.HeaderFooter.oddFooter = value


    @property
    def evenHeader(self):
        return self.HeaderFooter.evenHeader


    @evenHeader.setter
    def evenHeader(self, value):
        self.HeaderFooter.evenHeader = value


    @property
    def evenFooter(self):
        return self.HeaderFooter.evenFooter


    @evenFooter.setter
    def evenFooter(self, value):
        self.HeaderFooter.evenFooter = value


    @property
    def firstHeader(self):
        return self.HeaderFooter.firstHeader


    @firstHeader.setter
    def firstHeader(self, value):
        self.HeaderFooter.firstHeader = value


    @property
    def firstFooter(self):
        return self.HeaderFooter.firstFooter


    @firstFooter.setter
    def firstFooter(self, value):
        self.HeaderFooter.firstFooter = value


    @property
    def path(self):
        return self._path.format(self._id)

''',
    },
    'openpyxl.workbook.defined_name': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from collections import defaultdict
import re

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    String,
    Integer,
    Bool,
    Sequence,
    Descriptor,
)
from openpyxl.compat import safe_string
from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import SHEETRANGE_RE

RESERVED = frozenset(["Print_Area", "Print_Titles", "Criteria",
                      "_FilterDatabase", "Extract", "Consolidate_Area",
                      "Sheet_Title"])

_names = "|".join(RESERVED)
RESERVED_REGEX = re.compile(r"^_xlnm\.(?P<name>{0})".format(_names))


class DefinedName(Serialisable):

    tagname = "definedName"

    name = String() # unique per workbook/worksheet
    comment = String(allow_none=True)
    customMenu = String(allow_none=True)
    description = String(allow_none=True)
    help = String(allow_none=True)
    statusBar = String(allow_none=True)
    localSheetId = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    function = Bool(allow_none=True)
    vbProcedure = Bool(allow_none=True)
    xlm = Bool(allow_none=True)
    functionGroupId = Integer(allow_none=True)
    shortcutKey = String(allow_none=True)
    publishToServer = Bool(allow_none=True)
    workbookParameter = Bool(allow_none=True)
    attr_text = Descriptor()
    value = Alias("attr_text")


    def __init__(self,
                 name=None,
                 comment=None,
                 customMenu=None,
                 description=None,
                 help=None,
                 statusBar=None,
                 localSheetId=None,
                 hidden=None,
                 function=None,
                 vbProcedure=None,
                 xlm=None,
                 functionGroupId=None,
                 shortcutKey=None,
                 publishToServer=None,
                 workbookParameter=None,
                 attr_text=None
                ):
        self.name = name
        self.comment = comment
        self.customMenu = customMenu
        self.description = description
        self.help = help
        self.statusBar = statusBar
        self.localSheetId = localSheetId
        self.hidden = hidden
        self.function = function
        self.vbProcedure = vbProcedure
        self.xlm = xlm
        self.functionGroupId = functionGroupId
        self.shortcutKey = shortcutKey
        self.publishToServer = publishToServer
        self.workbookParameter = workbookParameter
        self.attr_text = attr_text


    @property
    def type(self):
        tok = Tokenizer("=" + self.value)
        parsed = tok.items[0]
        if parsed.type == "OPERAND":
            return parsed.subtype
        return parsed.type


    @property
    def destinations(self):
        if self.type == "RANGE":
            tok = Tokenizer("=" + self.value)
            for part in tok.items:
                if part.subtype == "RANGE":
                    m = SHEETRANGE_RE.match(part.value)
                    sheetname = m.group('notquoted') or m.group('quoted')
                    yield sheetname, m.group('cells')


    @property
    def is_reserved(self):
        m = RESERVED_REGEX.match(self.name)
        if m:
            return m.group("name")


    @property
    def is_external(self):
        return re.compile(r"^\[\d+\].*").match(self.value) is not None


    def __iter__(self):
        for key in self.__attrs__:
            if key == "attr_text":
                continue
            v = getattr(self, key)
            if v is not None:
                if v in RESERVED:
                    v = "_xlnm." + v
                yield key, safe_string(v)


class DefinedNameDict(dict):

    """
    Utility class for storing defined names.
    Allows access by name and separation of global and scoped names
    """

    def __setitem__(self, key, value):
        if not isinstance(value, DefinedName):
            raise TypeError("Value must be a an instance of DefinedName")
        elif value.name != key:
            raise ValueError("Key must be the same as the name")
        super().__setitem__(key, value)


    def add(self, value):
        """
        Add names without worrying about key and name matching.
        """
        self[value.name] = value


class DefinedNameList(Serialisable):

    tagname = "definedNames"

    definedName = Sequence(expected_type=DefinedName)


    def __init__(self, definedName=()):
        self.definedName = definedName


    def by_sheet(self):
        """
        Break names down into sheet locals and globals
        """
        names = defaultdict(DefinedNameDict)
        for defn in self.definedName:
            if defn.localSheetId is None:
                if defn.name in ("_xlnm.Print_Titles", "_xlnm.Print_Area", "_xlnm._FilterDatabase"):
                    continue
                names["global"][defn.name] = defn
            else:
                sheet = int(defn.localSheetId)
                names[sheet][defn.name] = defn
        return names


    def _duplicate(self, defn):
        """
        Check for whether DefinedName with the same name and scope already
        exists
        """
        for d in self.definedName:
            if d.name == defn.name and d.localSheetId == defn.localSheetId:
                return True


    def __len__(self):
        return len(self.definedName)

''',
    },
    'openpyxl.workbook.external_link': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from .external import ExternalLink

""",
    },
    'openpyxl.workbook.external_link.external': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Bool,
    Integer,
    NoneSet,
    Sequence,
)
from openpyxl.descriptors.excel import Relation
from openpyxl.descriptors.nested import NestedText
from openpyxl.descriptors.sequence import NestedSequence, ValueSequence

from openpyxl.packaging.relationship import (
    Relationship,
    get_rels_path,
    get_dependents
    )
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring


"""Manage links to external Workbooks"""


class ExternalCell(Serialisable):

    r = String()
    t = NoneSet(values=(['b', 'd', 'n', 'e', 's', 'str', 'inlineStr']))
    vm = Integer(allow_none=True)
    v = NestedText(allow_none=True, expected_type=str)

    def __init__(self,
                 r=None,
                 t=None,
                 vm=None,
                 v=None,
                ):
        self.r = r
        self.t = t
        self.vm = vm
        self.v = v


class ExternalRow(Serialisable):

    r = Integer()
    cell = Sequence(expected_type=ExternalCell)

    __elements__ = ('cell',)

    def __init__(self,
                 r=(),
                 cell=None,
                ):
        self.r = r
        self.cell = cell


class ExternalSheetData(Serialisable):

    sheetId = Integer()
    refreshError = Bool(allow_none=True)
    row = Sequence(expected_type=ExternalRow)

    __elements__ = ('row',)

    def __init__(self,
                 sheetId=None,
                 refreshError=None,
                 row=(),
                ):
        self.sheetId = sheetId
        self.refreshError = refreshError
        self.row = row


class ExternalSheetDataSet(Serialisable):

    sheetData = Sequence(expected_type=ExternalSheetData, )

    __elements__ = ('sheetData',)

    def __init__(self,
                 sheetData=None,
                ):
        self.sheetData = sheetData


class ExternalSheetNames(Serialisable):

    sheetName = ValueSequence(expected_type=str)

    __elements__ = ('sheetName',)

    def __init__(self,
                 sheetName=(),
                ):
        self.sheetName = sheetName


class ExternalDefinedName(Serialisable):

    tagname = "definedName"

    name = String()
    refersTo = String(allow_none=True)
    sheetId = Integer(allow_none=True)

    def __init__(self,
                 name=None,
                 refersTo=None,
                 sheetId=None,
                ):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


class ExternalBook(Serialisable):

    tagname = "externalBook"

    sheetNames = Typed(expected_type=ExternalSheetNames, allow_none=True)
    definedNames = NestedSequence(expected_type=ExternalDefinedName)
    sheetDataSet = Typed(expected_type=ExternalSheetDataSet, allow_none=True)
    id = Relation()

    __elements__ = ('sheetNames', 'definedNames', 'sheetDataSet')

    def __init__(self,
                 sheetNames=None,
                 definedNames=(),
                 sheetDataSet=None,
                 id=None,
                ):
        self.sheetNames = sheetNames
        self.definedNames = definedNames
        self.sheetDataSet = sheetDataSet
        self.id = id


class ExternalLink(Serialisable):

    tagname = "externalLink"

    _id = None
    _path = "/xl/externalLinks/externalLink{0}.xml"
    _rel_type = "externalLink"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"

    externalBook = Typed(expected_type=ExternalBook, allow_none=True)
    file_link = Typed(expected_type=Relationship, allow_none=True) # link to external file

    __elements__ = ('externalBook', )

    def __init__(self,
                 externalBook=None,
                 ddeLink=None,
                 oleLink=None,
                 extLst=None,
                ):
        self.externalBook = externalBook
        # ignore other items for the moment.


    def to_tree(self):
        node = super().to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node


    @property
    def path(self):
        return self._path.format(self._id)


def read_external_link(archive, book_path):
    src = archive.read(book_path)
    node = fromstring(src)
    book = ExternalLink.from_tree(node)

    link_path = get_rels_path(book_path)
    deps = get_dependents(archive, link_path)
    book.file_link = deps[0]

    return book

''',
    },
    'openpyxl.workbook.external_reference': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence
)
from openpyxl.descriptors.excel import (
    Relation,
)

class ExternalReference(Serialisable):

    tagname = "externalReference"

    id = Relation()

    def __init__(self, id):
        self.id = id

""",
    },
    'openpyxl.workbook.function_group': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    String,
    Integer,
)

class FunctionGroup(Serialisable):

    tagname = "functionGroup"

    name = String()

    def __init__(self,
                 name=None,
                ):
        self.name = name


class FunctionGroupList(Serialisable):

    tagname = "functionGroups"

    builtInGroupCount = Integer(allow_none=True)
    functionGroup = Sequence(expected_type=FunctionGroup, allow_none=True)

    __elements__ = ('functionGroup',)

    def __init__(self,
                 builtInGroupCount=16,
                 functionGroup=(),
                ):
        self.builtInGroupCount = builtInGroupCount
        self.functionGroup = functionGroup

""",
    },
    'openpyxl.workbook.properties': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    String,
    Float,
    Integer,
    Bool,
    NoneSet,
    Set,
)

from openpyxl.descriptors.excel import Guid


class WorkbookProperties(Serialisable):

    tagname = "workbookPr"

    date1904 = Bool(allow_none=True)
    dateCompatibility = Bool(allow_none=True)
    showObjects = NoneSet(values=(['all', 'placeholders']))
    showBorderUnselectedTables = Bool(allow_none=True)
    filterPrivacy = Bool(allow_none=True)
    promptedSolutions = Bool(allow_none=True)
    showInkAnnotation = Bool(allow_none=True)
    backupFile = Bool(allow_none=True)
    saveExternalLinkValues = Bool(allow_none=True)
    updateLinks = NoneSet(values=(['userSet', 'never', 'always']))
    codeName = String(allow_none=True)
    hidePivotFieldList = Bool(allow_none=True)
    showPivotChartFilter = Bool(allow_none=True)
    allowRefreshQuery = Bool(allow_none=True)
    publishItems = Bool(allow_none=True)
    checkCompatibility = Bool(allow_none=True)
    autoCompressPictures = Bool(allow_none=True)
    refreshAllConnections = Bool(allow_none=True)
    defaultThemeVersion = Integer(allow_none=True)

    def __init__(self,
                 date1904=None,
                 dateCompatibility=None,
                 showObjects=None,
                 showBorderUnselectedTables=None,
                 filterPrivacy=None,
                 promptedSolutions=None,
                 showInkAnnotation=None,
                 backupFile=None,
                 saveExternalLinkValues=None,
                 updateLinks=None,
                 codeName=None,
                 hidePivotFieldList=None,
                 showPivotChartFilter=None,
                 allowRefreshQuery=None,
                 publishItems=None,
                 checkCompatibility=None,
                 autoCompressPictures=None,
                 refreshAllConnections=None,
                 defaultThemeVersion=None,
                ):
        self.date1904 = date1904
        self.dateCompatibility = dateCompatibility
        self.showObjects = showObjects
        self.showBorderUnselectedTables = showBorderUnselectedTables
        self.filterPrivacy = filterPrivacy
        self.promptedSolutions = promptedSolutions
        self.showInkAnnotation = showInkAnnotation
        self.backupFile = backupFile
        self.saveExternalLinkValues = saveExternalLinkValues
        self.updateLinks = updateLinks
        self.codeName = codeName
        self.hidePivotFieldList = hidePivotFieldList
        self.showPivotChartFilter = showPivotChartFilter
        self.allowRefreshQuery = allowRefreshQuery
        self.publishItems = publishItems
        self.checkCompatibility = checkCompatibility
        self.autoCompressPictures = autoCompressPictures
        self.refreshAllConnections = refreshAllConnections
        self.defaultThemeVersion = defaultThemeVersion


class CalcProperties(Serialisable):

    tagname = "calcPr"

    calcId = Integer()
    calcMode = NoneSet(values=(['manual', 'auto', 'autoNoTable']))
    fullCalcOnLoad = Bool(allow_none=True)
    refMode = NoneSet(values=(['A1', 'R1C1']))
    iterate = Bool(allow_none=True)
    iterateCount = Integer(allow_none=True)
    iterateDelta = Float(allow_none=True)
    fullPrecision = Bool(allow_none=True)
    calcCompleted = Bool(allow_none=True)
    calcOnSave = Bool(allow_none=True)
    concurrentCalc = Bool(allow_none=True)
    concurrentManualCount = Integer(allow_none=True)
    forceFullCalc = Bool(allow_none=True)

    def __init__(self,
                 calcId=124519,
                 calcMode=None,
                 fullCalcOnLoad=True,
                 refMode=None,
                 iterate=None,
                 iterateCount=None,
                 iterateDelta=None,
                 fullPrecision=None,
                 calcCompleted=None,
                 calcOnSave=None,
                 concurrentCalc=None,
                 concurrentManualCount=None,
                 forceFullCalc=None,
                ):
        self.calcId = calcId
        self.calcMode = calcMode
        self.fullCalcOnLoad = fullCalcOnLoad
        self.refMode = refMode
        self.iterate = iterate
        self.iterateCount = iterateCount
        self.iterateDelta = iterateDelta
        self.fullPrecision = fullPrecision
        self.calcCompleted = calcCompleted
        self.calcOnSave = calcOnSave
        self.concurrentCalc = concurrentCalc
        self.concurrentManualCount = concurrentManualCount
        self.forceFullCalc = forceFullCalc


class FileVersion(Serialisable):

    tagname = "fileVersion"

    appName = String(allow_none=True)
    lastEdited = String(allow_none=True)
    lowestEdited = String(allow_none=True)
    rupBuild = String(allow_none=True)
    codeName = Guid(allow_none=True)

    def __init__(self,
                 appName=None,
                 lastEdited=None,
                 lowestEdited=None,
                 rupBuild=None,
                 codeName=None,
                ):
        self.appName = appName
        self.lastEdited = lastEdited
        self.lowestEdited = lowestEdited
        self.rupBuild = rupBuild
        self.codeName = codeName

""",
    },
    'openpyxl.workbook.protection': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    String,
    Float,
    Integer,
    Bool,
    NoneSet,
    Set,
)
from openpyxl.descriptors.excel import (
    ExtensionList,
    HexBinary,
    Guid,
    Relation,
    Base64Binary,
)
from openpyxl.utils.protection import hash_password


class WorkbookProtection(Serialisable):

    _workbook_password, _revisions_password = None, None

    tagname = "workbookPr"

    workbook_password = Alias("workbookPassword")
    workbookPasswordCharacterSet = String(allow_none=True)
    revision_password = Alias("revisionsPassword")
    revisionsPasswordCharacterSet = String(allow_none=True)
    lockStructure = Bool(allow_none=True)
    lock_structure = Alias("lockStructure")
    lockWindows = Bool(allow_none=True)
    lock_windows = Alias("lockWindows")
    lockRevision = Bool(allow_none=True)
    lock_revision = Alias("lockRevision")
    revisionsAlgorithmName = String(allow_none=True)
    revisionsHashValue = Base64Binary(allow_none=True)
    revisionsSaltValue = Base64Binary(allow_none=True)
    revisionsSpinCount = Integer(allow_none=True)
    workbookAlgorithmName = String(allow_none=True)
    workbookHashValue = Base64Binary(allow_none=True)
    workbookSaltValue = Base64Binary(allow_none=True)
    workbookSpinCount = Integer(allow_none=True)

    __attrs__ = ('workbookPassword', 'workbookPasswordCharacterSet', 'revisionsPassword',
                 'revisionsPasswordCharacterSet', 'lockStructure', 'lockWindows', 'lockRevision',
                 'revisionsAlgorithmName', 'revisionsHashValue', 'revisionsSaltValue',
                 'revisionsSpinCount', 'workbookAlgorithmName', 'workbookHashValue',
                 'workbookSaltValue', 'workbookSpinCount')

    def __init__(self,
                 workbookPassword=None,
                 workbookPasswordCharacterSet=None,
                 revisionsPassword=None,
                 revisionsPasswordCharacterSet=None,
                 lockStructure=None,
                 lockWindows=None,
                 lockRevision=None,
                 revisionsAlgorithmName=None,
                 revisionsHashValue=None,
                 revisionsSaltValue=None,
                 revisionsSpinCount=None,
                 workbookAlgorithmName=None,
                 workbookHashValue=None,
                 workbookSaltValue=None,
                 workbookSpinCount=None,
                ):
        if workbookPassword is not None:
            self.workbookPassword = workbookPassword
        self.workbookPasswordCharacterSet = workbookPasswordCharacterSet
        if revisionsPassword is not None:
            self.revisionsPassword = revisionsPassword
        self.revisionsPasswordCharacterSet = revisionsPasswordCharacterSet
        self.lockStructure = lockStructure
        self.lockWindows = lockWindows
        self.lockRevision = lockRevision
        self.revisionsAlgorithmName = revisionsAlgorithmName
        self.revisionsHashValue = revisionsHashValue
        self.revisionsSaltValue = revisionsSaltValue
        self.revisionsSpinCount = revisionsSpinCount
        self.workbookAlgorithmName = workbookAlgorithmName
        self.workbookHashValue = workbookHashValue
        self.workbookSaltValue = workbookSaltValue
        self.workbookSpinCount = workbookSpinCount

    def set_workbook_password(self, value='', already_hashed=False):
        """Set a password on this workbook."""
        if not already_hashed:
            value = hash_password(value)
        self._workbook_password = value

    @property
    def workbookPassword(self):
        """Return the workbook password value, regardless of hash."""
        return self._workbook_password

    @workbookPassword.setter
    def workbookPassword(self, value):
        """Set a workbook password directly, forcing a hash step."""
        self.set_workbook_password(value)

    def set_revisions_password(self, value='', already_hashed=False):
        """Set a revision password on this workbook."""
        if not already_hashed:
            value = hash_password(value)
        self._revisions_password = value

    @property
    def revisionsPassword(self):
        """Return the revisions password value, regardless of hash."""
        return self._revisions_password

    @revisionsPassword.setter
    def revisionsPassword(self, value):
        """Set a revisions password directly, forcing a hash step."""
        self.set_revisions_password(value)

    @classmethod
    def from_tree(cls, node):
        """Don't hash passwords when deserialising from XML"""
        self = super().from_tree(node)
        if self.workbookPassword:
            self.set_workbook_password(node.get('workbookPassword'), already_hashed=True)
        if self.revisionsPassword:
            self.set_revisions_password(node.get('revisionsPassword'), already_hashed=True)
        return self

# Backwards compatibility
DocumentSecurity = WorkbookProtection


class FileSharing(Serialisable):

    tagname = "fileSharing"

    readOnlyRecommended = Bool(allow_none=True)
    userName = String(allow_none=True)
    reservationPassword = HexBinary(allow_none=True)
    algorithmName = String(allow_none=True)
    hashValue = Base64Binary(allow_none=True)
    saltValue = Base64Binary(allow_none=True)
    spinCount = Integer(allow_none=True)

    def __init__(self,
                 readOnlyRecommended=None,
                 userName=None,
                 reservationPassword=None,
                 algorithmName=None,
                 hashValue=None,
                 saltValue=None,
                 spinCount=None,
                ):
        self.readOnlyRecommended = readOnlyRecommended
        self.userName = userName
        self.reservationPassword = reservationPassword
        self.algorithmName = algorithmName
        self.hashValue = hashValue
        self.saltValue = saltValue
        self.spinCount = spinCount

''',
    },
    'openpyxl.workbook.smart_tags': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    String,
    Bool,
    NoneSet,

)

class SmartTag(Serialisable):

    tagname = "smartTagType"

    namespaceUri = String(allow_none=True)
    name = String(allow_none=True)
    url = String(allow_none=True)

    def __init__(self,
                 namespaceUri=None,
                 name=None,
                 url=None,
                ):
        self.namespaceUri = namespaceUri
        self.name = name
        self.url = url


class SmartTagList(Serialisable):

    tagname = "smartTagTypes"

    smartTagType = Sequence(expected_type=SmartTag, allow_none=True)

    __elements__ = ('smartTagType',)

    def __init__(self,
                 smartTagType=(),
                ):
        self.smartTagType = smartTagType


class SmartTagProperties(Serialisable):

    tagname = "smartTagPr"

    embed = Bool(allow_none=True)
    show = NoneSet(values=(['all', 'noIndicator']))

    def __init__(self,
                 embed=None,
                 show=None,
                ):
        self.embed = embed
        self.show = show

""",
    },
    'openpyxl.workbook.views': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Sequence,
    String,
    Float,
    Integer,
    Bool,
    NoneSet,
    Set,
)
from openpyxl.descriptors.excel import (
    ExtensionList,
    Guid,
)


class BookView(Serialisable):

    tagname = "workbookView"

    visibility = NoneSet(values=(['visible', 'hidden', 'veryHidden']))
    minimized = Bool(allow_none=True)
    showHorizontalScroll = Bool(allow_none=True)
    showVerticalScroll = Bool(allow_none=True)
    showSheetTabs = Bool(allow_none=True)
    xWindow = Integer(allow_none=True)
    yWindow = Integer(allow_none=True)
    windowWidth = Integer(allow_none=True)
    windowHeight = Integer(allow_none=True)
    tabRatio = Integer(allow_none=True)
    firstSheet = Integer(allow_none=True)
    activeTab = Integer(allow_none=True)
    autoFilterDateGrouping = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 visibility="visible",
                 minimized=False,
                 showHorizontalScroll=True,
                 showVerticalScroll=True,
                 showSheetTabs=True,
                 xWindow=None,
                 yWindow=None,
                 windowWidth=None,
                 windowHeight=None,
                 tabRatio=600,
                 firstSheet=0,
                 activeTab=0,
                 autoFilterDateGrouping=True,
                 extLst=None,
                ):
        self.visibility = visibility
        self.minimized = minimized
        self.showHorizontalScroll = showHorizontalScroll
        self.showVerticalScroll = showVerticalScroll
        self.showSheetTabs = showSheetTabs
        self.xWindow = xWindow
        self.yWindow = yWindow
        self.windowWidth = windowWidth
        self.windowHeight = windowHeight
        self.tabRatio = tabRatio
        self.firstSheet = firstSheet
        self.activeTab = activeTab
        self.autoFilterDateGrouping = autoFilterDateGrouping


class CustomWorkbookView(Serialisable):

    tagname = "customWorkbookView"

    name = String()
    guid = Guid()
    autoUpdate = Bool(allow_none=True)
    mergeInterval = Integer(allow_none=True)
    changesSavedWin = Bool(allow_none=True)
    onlySync = Bool(allow_none=True)
    personalView = Bool(allow_none=True)
    includePrintSettings = Bool(allow_none=True)
    includeHiddenRowCol = Bool(allow_none=True)
    maximized = Bool(allow_none=True)
    minimized = Bool(allow_none=True)
    showHorizontalScroll = Bool(allow_none=True)
    showVerticalScroll = Bool(allow_none=True)
    showSheetTabs = Bool(allow_none=True)
    xWindow = Integer(allow_none=True)
    yWindow = Integer(allow_none=True)
    windowWidth = Integer()
    windowHeight = Integer()
    tabRatio = Integer(allow_none=True)
    activeSheetId = Integer()
    showFormulaBar = Bool(allow_none=True)
    showStatusbar = Bool(allow_none=True)
    showComments = NoneSet(values=(['commNone', 'commIndicator',
                                'commIndAndComment']))
    showObjects = NoneSet(values=(['all', 'placeholders']))
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 name=None,
                 guid=None,
                 autoUpdate=None,
                 mergeInterval=None,
                 changesSavedWin=None,
                 onlySync=None,
                 personalView=None,
                 includePrintSettings=None,
                 includeHiddenRowCol=None,
                 maximized=None,
                 minimized=None,
                 showHorizontalScroll=None,
                 showVerticalScroll=None,
                 showSheetTabs=None,
                 xWindow=None,
                 yWindow=None,
                 windowWidth=None,
                 windowHeight=None,
                 tabRatio=None,
                 activeSheetId=None,
                 showFormulaBar=None,
                 showStatusbar=None,
                 showComments="commIndicator",
                 showObjects="all",
                 extLst=None,
                ):
        self.name = name
        self.guid = guid
        self.autoUpdate = autoUpdate
        self.mergeInterval = mergeInterval
        self.changesSavedWin = changesSavedWin
        self.onlySync = onlySync
        self.personalView = personalView
        self.includePrintSettings = includePrintSettings
        self.includeHiddenRowCol = includeHiddenRowCol
        self.maximized = maximized
        self.minimized = minimized
        self.showHorizontalScroll = showHorizontalScroll
        self.showVerticalScroll = showVerticalScroll
        self.showSheetTabs = showSheetTabs
        self.xWindow = xWindow
        self.yWindow = yWindow
        self.windowWidth = windowWidth
        self.windowHeight = windowHeight
        self.tabRatio = tabRatio
        self.activeSheetId = activeSheetId
        self.showFormulaBar = showFormulaBar
        self.showStatusbar = showStatusbar
        self.showComments = showComments
        self.showObjects = showObjects

""",
    },
    'openpyxl.workbook.web': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Sequence,
    String,
    Float,
    Integer,
    Bool,
    NoneSet,
)


class WebPublishObject(Serialisable):

    tagname = "webPublishingObject"

    id = Integer()
    divId = String()
    sourceObject = String(allow_none=True)
    destinationFile = String()
    title = String(allow_none=True)
    autoRepublish = Bool(allow_none=True)

    def __init__(self,
                 id=None,
                 divId=None,
                 sourceObject=None,
                 destinationFile=None,
                 title=None,
                 autoRepublish=None,
                ):
        self.id = id
        self.divId = divId
        self.sourceObject = sourceObject
        self.destinationFile = destinationFile
        self.title = title
        self.autoRepublish = autoRepublish


class WebPublishObjectList(Serialisable):

    tagname ="webPublishingObjects"

    count = Integer(allow_none=True)
    webPublishObject = Sequence(expected_type=WebPublishObject)

    __elements__ = ('webPublishObject',)

    def __init__(self,
                 count=None,
                 webPublishObject=(),
                ):
        self.webPublishObject = webPublishObject


    @property
    def count(self):
        return len(self.webPublishObject)


class WebPublishing(Serialisable):

    tagname = "webPublishing"

    css = Bool(allow_none=True)
    thicket = Bool(allow_none=True)
    longFileNames = Bool(allow_none=True)
    vml = Bool(allow_none=True)
    allowPng = Bool(allow_none=True)
    targetScreenSize = NoneSet(values=(['544x376', '640x480', '720x512', '800x600',
                                    '1024x768', '1152x882', '1152x900', '1280x1024', '1600x1200',
                                    '1800x1440', '1920x1200']))
    dpi = Integer(allow_none=True)
    codePage = Integer(allow_none=True)
    characterSet = String(allow_none=True)

    def __init__(self,
                 css=None,
                 thicket=None,
                 longFileNames=None,
                 vml=None,
                 allowPng=None,
                 targetScreenSize='800x600',
                 dpi=None,
                 codePage=None,
                 characterSet=None,
                ):
        self.css = css
        self.thicket = thicket
        self.longFileNames = longFileNames
        self.vml = vml
        self.allowPng = allowPng
        self.targetScreenSize = targetScreenSize
        self.dpi = dpi
        self.codePage = codePage
        self.characterSet = characterSet

""",
    },
    'openpyxl.workbook.workbook': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Workbook is the top-level container for all document information."""
from copy import copy

from openpyxl.compat import deprecated
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
from openpyxl.worksheet.copier import WorksheetCopy

from openpyxl.utils import quote_sheetname
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.datetime  import WINDOWS_EPOCH, MAC_EPOCH
from openpyxl.utils.exceptions import ReadOnlyWorkbookException

from openpyxl.writer.excel import save_workbook

from openpyxl.styles.cell_style import StyleArray
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.styles.differential import DifferentialStyleList
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import DEFAULT_BORDER
from openpyxl.styles.fills import DEFAULT_EMPTY_FILL, DEFAULT_GRAY_FILL
from openpyxl.styles.fonts import DEFAULT_FONT
from openpyxl.styles.protection import Protection
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.styles.named_styles import NamedStyleList
from openpyxl.styles.table import TableStyleList

from openpyxl.chartsheet import Chartsheet
from .defined_name import DefinedName, DefinedNameDict
from openpyxl.packaging.core import DocumentProperties
from openpyxl.packaging.custom import CustomPropertyList
from openpyxl.packaging.relationship import RelationshipList
from .child import _WorkbookChild
from .protection import DocumentSecurity
from .properties import CalcProperties
from .views import BookView


from openpyxl.xml.constants import (
    XLSM,
    XLSX,
    XLTM,
    XLTX
)

INTEGER_TYPES = (int,)

class Workbook:
    """Workbook is the container for all other parts of the document."""

    _read_only = False
    _data_only = False
    template = False
    path = "/xl/workbook.xml"

    def __init__(self,
                 write_only=False,
                 iso_dates=False,
                 ):
        self._sheets = []
        self._pivots = []
        self._active_sheet_index = 0
        self.defined_names = DefinedNameDict()
        self._external_links = []
        self.properties = DocumentProperties()
        self.custom_doc_props = CustomPropertyList()
        self.security = DocumentSecurity()
        self.__write_only = write_only
        self.shared_strings = IndexedList()

        self._setup_styles()

        self.loaded_theme = None
        self.vba_archive = None
        self.is_template = False
        self.code_name = None
        self.epoch = WINDOWS_EPOCH
        self.encoding = "utf-8"
        self.iso_dates = iso_dates

        if not self.write_only:
            self._sheets.append(Worksheet(self))

        self.rels = RelationshipList()
        self.calculation = CalcProperties()
        self.views = [BookView()]


    def _setup_styles(self):
        """Bootstrap styles"""

        self._fonts = IndexedList()
        self._fonts.add(DEFAULT_FONT)

        self._alignments = IndexedList([Alignment()])

        self._borders = IndexedList()
        self._borders.add(DEFAULT_BORDER)

        self._fills = IndexedList()
        self._fills.add(DEFAULT_EMPTY_FILL)
        self._fills.add(DEFAULT_GRAY_FILL)

        self._number_formats = IndexedList()
        self._date_formats = {}
        self._timedelta_formats = {}

        self._protections = IndexedList([Protection()])

        self._colors = COLOR_INDEX
        self._cell_styles = IndexedList([StyleArray()])
        self._named_styles = NamedStyleList()
        self.add_named_style(NamedStyle(font=copy(DEFAULT_FONT), border=copy(DEFAULT_BORDER), builtinId=0))
        self._table_styles = TableStyleList()
        self._differential_styles = DifferentialStyleList()


    @property
    def epoch(self):
        if self._epoch == WINDOWS_EPOCH:
            return WINDOWS_EPOCH
        return MAC_EPOCH


    @epoch.setter
    def epoch(self, value):
        if value not in (WINDOWS_EPOCH, MAC_EPOCH):
            raise ValueError("The epoch must be either 1900 or 1904")
        self._epoch = value


    @property
    def read_only(self):
        return self._read_only

    @property
    def data_only(self):
        return self._data_only

    @property
    def write_only(self):
        return self.__write_only


    @property
    def excel_base_date(self):
        return self.epoch

    @property
    def active(self):
        """Get the currently active sheet or None

        :type: :class:`openpyxl.worksheet.worksheet.Worksheet`
        """
        try:
            return self._sheets[self._active_sheet_index]
        except IndexError:
            pass

    @active.setter
    def active(self, value):
        """Set the active sheet"""
        if not isinstance(value, (_WorkbookChild, INTEGER_TYPES)):
            raise TypeError("Value must be either a worksheet, chartsheet or numerical index")
        if isinstance(value, INTEGER_TYPES):
            self._active_sheet_index = value
            return
            #if self._sheets and 0 <= value < len(self._sheets):
                #value = self._sheets[value]
            #else:
                #raise ValueError("Sheet index is outside the range of possible values", value)
        if value not in self._sheets:
            raise ValueError("Worksheet is not in the workbook")
        if value.sheet_state != "visible":
            raise ValueError("Only visible sheets can be made active")

        idx = self._sheets.index(value)
        self._active_sheet_index = idx


    def create_sheet(self, title=None, index=None):
        """Create a worksheet (at an optional index).

        :param title: optional title of the sheet
        :type title: str
        :param index: optional position at which the sheet will be inserted
        :type index: int

        """
        if self.read_only:
            raise ReadOnlyWorkbookException('Cannot create new sheet in a read-only workbook')

        if self.write_only :
            new_ws = WriteOnlyWorksheet(parent=self, title=title)
        else:
            new_ws = Worksheet(parent=self, title=title)

        self._add_sheet(sheet=new_ws, index=index)
        return new_ws


    def _add_sheet(self, sheet, index=None):
        """Add an worksheet (at an optional index)."""

        if not isinstance(sheet, (Worksheet, WriteOnlyWorksheet, Chartsheet)):
            raise TypeError("Cannot be added to a workbook")

        if sheet.parent != self:
            raise ValueError("You cannot add worksheets from another workbook.")

        if index is None:
            self._sheets.append(sheet)
        else:
            self._sheets.insert(index, sheet)


    def move_sheet(self, sheet, offset=0):
        """
        Move a sheet or sheetname
        """
        if not isinstance(sheet, Worksheet):
            sheet = self[sheet]
        idx = self._sheets.index(sheet)
        del self._sheets[idx]
        new_pos = idx + offset
        self._sheets.insert(new_pos, sheet)


    def remove(self, worksheet):
        """Remove `worksheet` from this workbook."""
        idx = self._sheets.index(worksheet)
        self._sheets.remove(worksheet)


    @deprecated("Use wb.remove(worksheet) or del wb[sheetname]")
    def remove_sheet(self, worksheet):
        """Remove `worksheet` from this workbook."""
        self.remove(worksheet)


    def create_chartsheet(self, title=None, index=None):
        if self.read_only:
            raise ReadOnlyWorkbookException("Cannot create new sheet in a read-only workbook")
        cs = Chartsheet(parent=self, title=title)

        self._add_sheet(cs, index)
        return cs


    @deprecated("Use wb[sheetname]")
    def get_sheet_by_name(self, name):
        """Returns a worksheet by its name.

        :param name: the name of the worksheet to look for
        :type name: string

        """
        return self[name]

    def __contains__(self, key):
        return key in self.sheetnames


    def index(self, worksheet):
        """Return the index of a worksheet."""
        return self.worksheets.index(worksheet)


    @deprecated("Use wb.index(worksheet)")
    def get_index(self, worksheet):
        """Return the index of the worksheet."""
        return self.index(worksheet)

    def __getitem__(self, key):
        """Returns a worksheet by its name.

        :param name: the name of the worksheet to look for
        :type name: string

        """
        for sheet in self.worksheets + self.chartsheets:
            if sheet.title == key:
                return sheet
        raise KeyError("Worksheet {0} does not exist.".format(key))

    def __delitem__(self, key):
        sheet = self[key]
        self.remove(sheet)

    def __iter__(self):
        return iter(self.worksheets)


    @deprecated("Use wb.sheetnames")
    def get_sheet_names(self):
        return self.sheetnames

    @property
    def worksheets(self):
        """A list of sheets in this workbook

        :type: list of :class:`openpyxl.worksheet.worksheet.Worksheet`
        """
        return [s for s in self._sheets if isinstance(s, (Worksheet, ReadOnlyWorksheet, WriteOnlyWorksheet))]

    @property
    def chartsheets(self):
        """A list of Chartsheets in this workbook

        :type: list of :class:`openpyxl.chartsheet.chartsheet.Chartsheet`
        """
        return [s for s in self._sheets if isinstance(s, Chartsheet)]

    @property
    def sheetnames(self):
        """Returns the list of the names of worksheets in this workbook.

        Names are returned in the worksheets order.

        :type: list of strings

        """
        return [s.title for s in self._sheets]


    @deprecated("Assign scoped named ranges directly to worksheets or global ones to the workbook. Deprecated in 3.1")
    def create_named_range(self, name, worksheet=None, value=None, scope=None):
        """Create a new named_range on a worksheet

        """
        defn = DefinedName(name=name)
        if worksheet is not None:
            defn.value = "{0}!{1}".format(quote_sheetname(worksheet.title), value)
        else:
            defn.value = value

        self.defined_names[name] = defn


    def add_named_style(self, style):
        """
        Add a named style
        """
        self._named_styles.append(style)
        style.bind(self)


    @property
    def named_styles(self):
        """
        List available named styles
        """
        return self._named_styles.names


    @property
    def mime_type(self):
        """
        The mime type is determined by whether a workbook is a template or
        not and whether it contains macros or not. Excel requires the file
        extension to match but openpyxl does not enforce this.

        """
        ct = self.template and XLTX or XLSX
        if self.vba_archive:
            ct = self.template and XLTM or XLSM
        return ct


    def save(self, filename):
        """Save the current workbook under the given `filename`.
        Use this function instead of using an `ExcelWriter`.

        .. warning::
            When creating your workbook using `write_only` set to True,
            you will only be able to call this function once. Subsequent attempts to
            modify or save the file will raise an :class:`openpyxl.shared.exc.WorkbookAlreadySaved` exception.
        """
        if self.read_only:
            raise TypeError("""Workbook is read-only""")
        if self.write_only and not self.worksheets:
            self.create_sheet()
        save_workbook(self, filename)


    @property
    def style_names(self):
        """
        List of named styles
        """
        return [s.name for s in self._named_styles]


    def copy_worksheet(self, from_worksheet):
        """Copy an existing worksheet in the current workbook

        .. warning::
            This function cannot copy worksheets between workbooks.
            worksheets can only be copied within the workbook that they belong

        :param from_worksheet: the worksheet to be copied from
        :return: copy of the initial worksheet
        """
        if self.__write_only or self._read_only:
            raise ValueError("Cannot copy worksheets in read-only or write-only mode")

        new_title = u"{0} Copy".format(from_worksheet.title)
        to_worksheet = self.create_sheet(title=new_title)
        cp = WorksheetCopy(source_worksheet=from_worksheet, target_worksheet=to_worksheet)
        cp.copy_worksheet()
        return to_worksheet


    def close(self):
        """
        Close workbook file if open. Only affects read-only and write-only modes.
        """
        if hasattr(self, '_archive'):
            self._archive.close()


    def _duplicate_name(self, name):
        """
        Check for duplicate name in defined name list and table list of each worksheet.
        Names are not case sensitive.
        """
        name = name.lower()
        for sheet in self.worksheets:
            for t in sheet.tables:
                if name == t.lower():
                    return True

        if name in self.defined_names:
            return True


''',
    },
    'openpyxl.worksheet': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

""",
    },
    'openpyxl.worksheet._read_only': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

""" Read worksheets on-demand
"""

from .worksheet import Worksheet
from openpyxl.cell.read_only import ReadOnlyCell, EMPTY_CELL
from openpyxl.utils import get_column_letter

from ._reader import WorkSheetParser
from openpyxl.workbook.defined_name import DefinedNameDict


def read_dimension(source):
    parser = WorkSheetParser(source, [])
    return parser.parse_dimensions()


class ReadOnlyWorksheet:

    _min_column = 1
    _min_row = 1
    _max_column = _max_row = None

    # from Standard Worksheet
    # Methods from Worksheet
    cell = Worksheet.cell
    iter_rows = Worksheet.iter_rows
    values = Worksheet.values
    rows = Worksheet.rows
    __getitem__ = Worksheet.__getitem__
    __iter__ = Worksheet.__iter__


    def __init__(self, parent_workbook, title, worksheet_path, shared_strings):
        self.parent = parent_workbook
        self.title = title
        self.sheet_state = 'visible'
        self._current_row = None
        self._worksheet_path = worksheet_path
        self._shared_strings = shared_strings
        self._get_size()
        self.defined_names = DefinedNameDict()


    def _get_size(self):
        src = self._get_source()
        parser = WorkSheetParser(src, [])
        dimensions = parser.parse_dimensions()
        src.close()
        if dimensions is not None:
            self._min_column, self._min_row, self._max_column, self._max_row = dimensions


    def _get_source(self):
        """Parse xml source on demand, must close after use"""
        return self.parent._archive.open(self._worksheet_path)


    def _cells_by_row(self, min_col, min_row, max_col, max_row, values_only=False):
        """
        The source worksheet file may have columns or rows missing.
        Missing cells will be created.
        """
        filler = EMPTY_CELL
        if values_only:
            filler = None

        max_col = max_col or self.max_column
        max_row = max_row or self.max_row
        empty_row = []
        if max_col is not None:
            empty_row = (filler,) * (max_col + 1 - min_col)

        counter = min_row
        idx = 1
        with self._get_source() as src:
            parser = WorkSheetParser(src,
                                     self._shared_strings,
                                     data_only=self.parent.data_only,
                                     epoch=self.parent.epoch,
                                     date_formats=self.parent._date_formats,
                                     timedelta_formats=self.parent._timedelta_formats)

            for idx, row in parser.parse():
                if max_row is not None and idx > max_row:
                    break

                # some rows are missing
                for _ in range(counter, idx):
                    counter += 1
                    yield empty_row

                # return cells from a row
                if counter <= idx:
                    row = self._get_row(row, min_col, max_col, values_only)
                    counter += 1
                    yield row

        if max_row is not None and max_row < idx:
            for _ in range(counter, max_row+1):
                yield empty_row


    def _get_row(self, row, min_col=1, max_col=None, values_only=False):
        """
        Make sure a row contains always the same number of cells or values
        """
        if not row and not max_col: # in case someone wants to force rows where there aren't any
            return ()

        max_col = max_col or  row[-1]['column']
        row_width = max_col + 1 - min_col

        new_row = [EMPTY_CELL] * row_width
        if values_only:
            new_row = [None] * row_width

        for cell in row:
            counter = cell['column']
            if min_col <= counter <= max_col:
                idx = counter - min_col # position in list of cells returned
                new_row[idx] = cell['value']
                if not values_only:
                    new_row[idx] = ReadOnlyCell(self, **cell)

        return tuple(new_row)


    def _get_cell(self, row, column):
        """Cells are returned by a generator which can be empty"""
        for row in self._cells_by_row(column, row, column, row):
            if row:
                return row[0]
        return EMPTY_CELL


    def calculate_dimension(self, force=False):
        if not all([self.max_column, self.max_row]):
            if force:
                self._calculate_dimension()
            else:
                raise ValueError("Worksheet is unsized, use calculate_dimension(force=True)")
        return f"{get_column_letter(self.min_column)}{self.min_row}:{get_column_letter(self.max_column)}{self.max_row}"


    def _calculate_dimension(self):
        """
        Loop through all the cells to get the size of a worksheet.
        Do this only if it is explicitly requested.
        """

        max_col = 0
        for r in self.rows:
            if not r:
                continue
            cell = r[-1]
            max_col = max(max_col, cell.column)

        self._max_row = cell.row
        self._max_column = max_col


    def reset_dimensions(self):
        """
        Remove worksheet dimensions if these are incorrect in the worksheet source.
        NB. This probably indicates a bug in the library or application that created
        the workbook.
        """
        self._max_row = self._max_column = None


    @property
    def min_row(self):
        return self._min_row


    @property
    def max_row(self):
        return self._max_row


    @property
    def min_column(self):
        return self._min_column


    @property
    def max_column(self):
        return self._max_column

''',
    },
    'openpyxl.worksheet._reader': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Reader for a single worksheet."""
from copy import copy
from warnings import warn

# compatibility imports
from openpyxl.xml.functions import iterparse

# package imports
from openpyxl.cell import Cell, MergedCell
from openpyxl.cell.text import Text
from openpyxl.worksheet.dimensions import (
    ColumnDimension,
    RowDimension,
    SheetFormatProperties,
)

from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    EXT_TYPES,
)
from openpyxl.formatting.formatting import ConditionalFormatting
from openpyxl.formula.translate import Translator
from openpyxl.utils import (
    get_column_letter,
    coordinate_to_tuple,
    )
from openpyxl.utils.datetime import from_excel, from_ISO8601, WINDOWS_EPOCH
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.cell.rich_text import CellRichText

from .formula import DataTableFormula, ArrayFormula
from .filters import AutoFilter
from .header_footer import HeaderFooter
from .hyperlink import HyperlinkList
from .merge import MergeCells
from .page import PageMargins, PrintOptions, PrintPageSetup
from .pagebreak import RowBreak, ColBreak
from .protection import SheetProtection
from .scenario import ScenarioList
from .views import SheetViewList
from .datavalidation import DataValidationList
from .table import TablePartList
from .properties import WorksheetProperties
from .dimensions import SheetDimension
from .related import Related


CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
MERGE_TAG = '{%s}mergeCells' % SHEET_MAIN_NS
INLINE_STRING = "{%s}is" % SHEET_MAIN_NS
COL_TAG = '{%s}col' % SHEET_MAIN_NS
ROW_TAG = '{%s}row' % SHEET_MAIN_NS
CF_TAG = '{%s}conditionalFormatting' % SHEET_MAIN_NS
LEGACY_TAG = '{%s}legacyDrawing' % SHEET_MAIN_NS
PROT_TAG = '{%s}sheetProtection' % SHEET_MAIN_NS
EXT_TAG = "{%s}extLst" % SHEET_MAIN_NS
HYPERLINK_TAG = "{%s}hyperlinks" % SHEET_MAIN_NS
TABLE_TAG = "{%s}tableParts" % SHEET_MAIN_NS
PRINT_TAG = '{%s}printOptions' % SHEET_MAIN_NS
MARGINS_TAG = '{%s}pageMargins' % SHEET_MAIN_NS
PAGE_TAG = '{%s}pageSetup' % SHEET_MAIN_NS
HEADER_TAG = '{%s}headerFooter' % SHEET_MAIN_NS
FILTER_TAG = '{%s}autoFilter' % SHEET_MAIN_NS
VALIDATION_TAG = '{%s}dataValidations' % SHEET_MAIN_NS
PROPERTIES_TAG = '{%s}sheetPr' % SHEET_MAIN_NS
VIEWS_TAG = '{%s}sheetViews' % SHEET_MAIN_NS
FORMAT_TAG = '{%s}sheetFormatPr' % SHEET_MAIN_NS
ROW_BREAK_TAG = '{%s}rowBreaks' % SHEET_MAIN_NS
COL_BREAK_TAG = '{%s}colBreaks' % SHEET_MAIN_NS
SCENARIOS_TAG = '{%s}scenarios' % SHEET_MAIN_NS
DATA_TAG = '{%s}sheetData' % SHEET_MAIN_NS
DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS
CUSTOM_VIEWS_TAG = '{%s}customSheetViews' % SHEET_MAIN_NS


def _cast_number(value):
    "Convert numbers as string to an int or float"
    if "." in value or "E" in value or "e" in value:
        return float(value)
    return int(value)


def parse_richtext_string(element):
    """
    Parse inline string and preserve rich text formatting
    """
    value = CellRichText.from_tree(element) or ""
    if len(value) == 1 and isinstance(value[0], str):
        value = value[0]
    return value


class WorkSheetParser:

    def __init__(self, src, shared_strings, data_only=False,
                 epoch=WINDOWS_EPOCH, date_formats=set(),
                 timedelta_formats=set(), rich_text=False):
        self.min_row = self.min_col = None
        self.epoch = epoch
        self.source = src
        self.shared_strings = shared_strings
        self.data_only = data_only
        self.shared_formulae = {}
        self.row_counter = self.col_counter = 0
        self.tables = TablePartList()
        self.date_formats = date_formats
        self.timedelta_formats = timedelta_formats
        self.row_dimensions = {}
        self.column_dimensions = {}
        self.number_formats = []
        self.keep_vba = False
        self.hyperlinks = HyperlinkList()
        self.formatting = []
        self.legacy_drawing = None
        self.merged_cells = None
        self.row_breaks = RowBreak()
        self.col_breaks = ColBreak()
        self.rich_text = rich_text


    def parse(self):
        dispatcher = {
            COL_TAG: self.parse_column_dimensions,
            PROT_TAG: self.parse_sheet_protection,
            EXT_TAG: self.parse_extensions,
            CF_TAG: self.parse_formatting,
            LEGACY_TAG: self.parse_legacy,
            ROW_BREAK_TAG: self.parse_row_breaks,
            COL_BREAK_TAG: self.parse_col_breaks,
            CUSTOM_VIEWS_TAG: self.parse_custom_views,
                      }

        properties = {
            PRINT_TAG: ('print_options', PrintOptions),
            MARGINS_TAG: ('page_margins', PageMargins),
            PAGE_TAG: ('page_setup', PrintPageSetup),
            HEADER_TAG: ('HeaderFooter', HeaderFooter),
            FILTER_TAG: ('auto_filter', AutoFilter),
            VALIDATION_TAG: ('data_validations', DataValidationList),
            PROPERTIES_TAG: ('sheet_properties', WorksheetProperties),
            VIEWS_TAG: ('views', SheetViewList),
            FORMAT_TAG: ('sheet_format', SheetFormatProperties),
            SCENARIOS_TAG: ('scenarios', ScenarioList),
            TABLE_TAG: ('tables', TablePartList),
            HYPERLINK_TAG: ('hyperlinks', HyperlinkList),
            MERGE_TAG: ('merged_cells', MergeCells),

        }

        it = iterparse(self.source) # add a finaliser to close the source when this becomes possible

        for _, element in it:
            tag_name = element.tag
            if tag_name in dispatcher:
                dispatcher[tag_name](element)
                element.clear()
            elif tag_name in properties:
                prop = properties[tag_name]
                obj = prop[1].from_tree(element)
                setattr(self, prop[0], obj)
                element.clear()
            elif tag_name == ROW_TAG:
                row = self.parse_row(element)
                element.clear()
                yield row


    def parse_dimensions(self):
        """
        Get worksheet dimensions if they are provided.
        """
        it = iterparse(self.source)

        for _event, element in it:
            if element.tag == DIMENSION_TAG:
                dim = SheetDimension.from_tree(element)
                return dim.boundaries

            elif element.tag == DATA_TAG:
                # Dimensions missing
                break
            element.clear()


    def parse_cell(self, element):
        data_type = element.get('t', 'n')
        coordinate = element.get('r')
        style_id = element.get('s', 0)
        if style_id:
            style_id = int(style_id)

        if data_type == "inlineStr":
            value = None
        else:
            value = element.findtext(VALUE_TAG, None) or None

        if coordinate:
            row, column = coordinate_to_tuple(coordinate)
            self.col_counter = column
        else:
            self.col_counter += 1
            row, column = self.row_counter, self.col_counter

        if not self.data_only and element.find(FORMULA_TAG) is not None:
            data_type = 'f'
            value = self.parse_formula(element)

        elif value is not None:
            if data_type == 'n':
                value = _cast_number(value)
                if style_id in self.date_formats:
                    data_type = 'd'
                    try:
                        value = from_excel(
                            value, self.epoch, timedelta=style_id in self.timedelta_formats
                        )
                    except (OverflowError, ValueError):
                        msg = f"""Cell {coordinate} is marked as a date but the serial value {value} is outside the limits for dates. The cell will be treated as an error."""
                        warn(msg)
                        data_type = "e"
                        value = "#VALUE!"
            elif data_type == 's':
                value = self.shared_strings[int(value)]
            elif data_type == 'b':
                value = bool(int(value))
            elif data_type == "str":
                data_type = "s"
            elif data_type == 'd':
                value = from_ISO8601(value)

        elif data_type == 'inlineStr':
                child = element.find(INLINE_STRING)
                if child is not None:
                    data_type = 's'
                    if self.rich_text:
                        value = parse_richtext_string(child)
                    else:
                        value = Text.from_tree(child).content

        return {'row':row, 'column':column, 'value':value, 'data_type':data_type, 'style_id':style_id}


    def parse_formula(self, element):
        """
        possible formulae types: shared, array, datatable
        """
        formula = element.find(FORMULA_TAG)
        formula_type = formula.get('t')
        coordinate = element.get('r')
        value = "="
        if formula.text is not None:
            value += formula.text

        if formula_type == "array":
            value = ArrayFormula(ref=formula.get('ref'), text=value)

        elif formula_type == "shared":
            idx = formula.get('si')
            if idx in self.shared_formulae:
                trans = self.shared_formulae[idx]
                value = trans.translate_formula(coordinate)
            elif value != "=":
                self.shared_formulae[idx] = Translator(value, coordinate)

        elif formula_type == "dataTable":
            value = DataTableFormula(**formula.attrib)

        return value


    def parse_column_dimensions(self, col):
        attrs = dict(col.attrib)
        column = get_column_letter(int(attrs['min']))
        attrs['index'] = column
        self.column_dimensions[column] = attrs


    def parse_row(self, row):
        attrs = dict(row.attrib)

        if "r" in attrs:
            try:
                self.row_counter = int(attrs['r'])
            except ValueError:
                val = float(attrs['r'])
                if val.is_integer():
                    self.row_counter = int(val)
                else:
                    raise ValueError(f"{attrs['r']} is not a valid row number")
        else:
            self.row_counter += 1
        self.col_counter = 0

        keys = {k for k in attrs if not k.startswith('{')}
        if keys - {'r', 'spans'}:
            # don't create dimension objects unless they have relevant information
            self.row_dimensions[str(self.row_counter)] = attrs

        cells = [self.parse_cell(el) for el in row]
        return self.row_counter, cells


    def parse_formatting(self, element):
        try:
            cf = ConditionalFormatting.from_tree(element)
            self.formatting.append(cf)
        except TypeError as e:
            msg = f"Failed to load a conditional formatting rule. It will be discarded. Cause: {e}"
            warn(msg)


    def parse_sheet_protection(self, element):
        protection = SheetProtection.from_tree(element)
        password = element.get("password")
        if password is not None:
            protection.set_password(password, True)
        self.protection = protection


    def parse_extensions(self, element):
        extLst = ExtensionList.from_tree(element)
        for e in extLst.ext:
            ext_type = EXT_TYPES.get(e.uri.upper(), "Unknown")
            msg = "{0} extension is not supported and will be removed".format(ext_type)
            warn(msg)


    def parse_legacy(self, element):
        obj = Related.from_tree(element)
        self.legacy_drawing = obj.id


    def parse_row_breaks(self, element):
        brk = RowBreak.from_tree(element)
        self.row_breaks = brk


    def parse_col_breaks(self, element):
        brk = ColBreak.from_tree(element)
        self.col_breaks = brk


    def parse_custom_views(self, element):
        # clear page_breaks to avoid duplication which Excel doesn't like
        # basically they're ignored in custom views
        self.row_breaks = RowBreak()
        self.col_breaks = ColBreak()


class WorksheetReader:
    """
    Create a parser and apply it to a workbook
    """

    def __init__(self, ws, xml_source, shared_strings, data_only, rich_text):
        self.ws = ws
        self.parser = WorkSheetParser(xml_source, shared_strings,
                data_only, ws.parent.epoch, ws.parent._date_formats,
                ws.parent._timedelta_formats, rich_text)
        self.tables = []


    def bind_cells(self):
        for idx, row in self.parser.parse():
            for cell in row:
                style = self.ws.parent._cell_styles[cell['style_id']]
                c = Cell(self.ws, row=cell['row'], column=cell['column'], style_array=style)
                c._value = cell['value']
                c.data_type = cell['data_type']
                self.ws._cells[(cell['row'], cell['column'])] = c

        if self.ws._cells:
            self.ws._current_row = self.ws.max_row # use cells not row dimensions


    def bind_formatting(self):
        for cf in self.parser.formatting:
            for rule in cf.rules:
                if rule.dxfId is not None:
                    rule.dxf = self.ws.parent._differential_styles[rule.dxfId]
                self.ws.conditional_formatting[cf] = rule


    def bind_tables(self):
        for t in self.parser.tables.tablePart:
            rel = self.ws._rels.get(t.id)
            self.tables.append(rel.Target)


    def bind_merged_cells(self):
        from openpyxl.worksheet.cell_range import MultiCellRange
        from openpyxl.worksheet.merge import MergedCellRange
        if not self.parser.merged_cells:
            return

        ranges = []
        for cr in self.parser.merged_cells.mergeCell:
            mcr = MergedCellRange(self.ws, cr.ref)
            self.ws._clean_merge_range(mcr)
            ranges.append(mcr)
        self.ws.merged_cells = MultiCellRange(ranges)


    def bind_hyperlinks(self):
        for link in self.parser.hyperlinks.hyperlink:
            if link.id:
                rel = self.ws._rels.get(link.id)
                link.target = rel.Target
            if ":" in link.ref:
                # range of cells
                for row in self.ws[link.ref]:
                    for cell in row:
                        try:
                            cell.hyperlink = copy(link)
                        except AttributeError:
                            pass
            else:
                cell = self.ws[link.ref]
                if isinstance(cell, MergedCell):
                    cell = self.normalize_merged_cell_link(cell.coordinate)
                cell.hyperlink = link

    def normalize_merged_cell_link(self, coord):
        """
        Returns the appropriate cell to which a hyperlink, which references a merged cell at the specified coordinates,
        should be bound.
        """
        for rng in self.ws.merged_cells:
            if coord in rng:
                return self.ws.cell(*rng.top[0])

    def bind_col_dimensions(self):
        for col, cd in self.parser.column_dimensions.items():
            if 'style' in cd:
                key = int(cd['style'])
                cd['style'] = self.ws.parent._cell_styles[key]
            self.ws.column_dimensions[col] = ColumnDimension(self.ws, **cd)


    def bind_row_dimensions(self):
        for row, rd in self.parser.row_dimensions.items():
            if 's' in rd:
                key = int(rd['s'])
                rd['s'] = self.ws.parent._cell_styles[key]
            self.ws.row_dimensions[int(row)] = RowDimension(self.ws, **rd)


    def bind_properties(self):
        for k in ('print_options', 'page_margins', 'page_setup',
                  'HeaderFooter', 'auto_filter', 'data_validations',
                  'sheet_properties', 'views', 'sheet_format',
                  'row_breaks', 'col_breaks', 'scenarios', 'legacy_drawing',
                  'protection',
                  ):
            v = getattr(self.parser, k, None)
            if v is not None:
                setattr(self.ws, k, v)


    def bind_all(self):
        self.bind_cells()
        self.bind_merged_cells()
        self.bind_hyperlinks()
        self.bind_formatting()
        self.bind_col_dimensions()
        self.bind_row_dimensions()
        self.bind_tables()
        self.bind_properties()

''',
    },
    'openpyxl.worksheet._write_only': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


"""Write worksheets to xml representations in an optimized way"""

from inspect import isgenerator

from openpyxl.cell import Cell, WriteOnlyCell
from openpyxl.workbook.child import _WorkbookChild
from .worksheet import Worksheet
from openpyxl.utils.exceptions import WorkbookAlreadySaved

from ._writer import WorksheetWriter


class WriteOnlyWorksheet(_WorkbookChild):
    """
    Streaming worksheet. Optimised to reduce memory by writing rows just in
    time.
    Cells can be styled and have comments Styles for rows and columns
    must be applied before writing cells
    """

    __saved = False
    _writer = None
    _rows = None
    _rel_type = Worksheet._rel_type
    _path = Worksheet._path
    mime_type = Worksheet.mime_type

    # copy methods from Standard worksheet
    _add_row = Worksheet._add_row
    _add_column = Worksheet._add_column
    add_chart = Worksheet.add_chart
    add_image = Worksheet.add_image
    add_table = Worksheet.add_table
    tables = Worksheet.tables
    print_titles = Worksheet.print_titles
    print_title_cols = Worksheet.print_title_cols
    print_title_rows = Worksheet.print_title_rows
    freeze_panes = Worksheet.freeze_panes
    print_area = Worksheet.print_area
    sheet_view = Worksheet.sheet_view
    _setup = Worksheet._setup

    def __init__(self, parent, title):
        super().__init__(parent, title)
        self._max_col = 0
        self._max_row = 0
        self._setup()

    @property
    def closed(self):
        return self.__saved


    def _write_rows(self):
        """
        Send rows to the writer's stream
        """
        try:
            xf = self._writer.xf.send(True)
        except StopIteration:
            self._already_saved()

        with xf.element("sheetData"):
            row_idx = 1
            try:
                while True:
                    row = (yield)
                    row = self._values_to_row(row, row_idx)
                    self._writer.write_row(xf, row, row_idx)
                    row_idx += 1
            except GeneratorExit:
                pass

        self._writer.xf.send(None)


    def _get_writer(self):
        if self._writer is None:
            self._writer = WorksheetWriter(self)
            self._writer.write_top()


    def close(self):
        if self.__saved:
            self._already_saved()

        self._get_writer()

        if self._rows is None:
            self._writer.write_rows()
        else:
            self._rows.close()

        self._writer.write_tail()

        self._writer.close()
        self.__saved = True


    def append(self, row):
        """
        :param row: iterable containing values to append
        :type row: iterable
        """

        if (not isgenerator(row) and
            not isinstance(row, (list, tuple, range))
            ):
            self._invalid_row(row)

        self._get_writer()

        if self._rows is None:
            self._rows = self._write_rows()
            next(self._rows)

        self._rows.send(row)


    def _values_to_row(self, values, row_idx):
        """
        Convert whatever has been appended into a form suitable for work_rows
        """
        cell = WriteOnlyCell(self)

        for col_idx, value in enumerate(values, 1):
            if value is None:
                continue
            try:
                cell.value = value
            except ValueError:
                if isinstance(value, Cell):
                    cell = value
                else:
                    raise ValueError

            cell.column = col_idx
            cell.row = row_idx

            if cell.hyperlink is not None:
                cell.hyperlink.ref = cell.coordinate

            yield cell

            # reset cell if style applied
            if cell.has_style or cell.hyperlink:
                cell = WriteOnlyCell(self)


    def _already_saved(self):
        raise WorkbookAlreadySaved('Workbook has already been saved and cannot be modified or saved anymore.')


    def _invalid_row(self, iterable):
        raise TypeError('Value must be a list, tuple, range or a generator Supplied value is {0}'.format(
            type(iterable))
                        )

''',
    },
    'openpyxl.worksheet._writer': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import atexit
from collections import defaultdict
from io import BytesIO
import os
from tempfile import NamedTemporaryFile
from warnings import warn

from openpyxl.xml.functions import xmlfile
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.comments.comment_sheet import CommentRecord
from openpyxl.packaging.relationship import Relationship, RelationshipList
from openpyxl.styles.differential import DifferentialStyle

from .dimensions import SheetDimension
from .hyperlink import HyperlinkList
from .merge import MergeCell, MergeCells
from .related import Related
from .table import TablePartList

from openpyxl.cell._writer import write_cell


ALL_TEMP_FILES = []

@atexit.register
def _openpyxl_shutdown():
    for path in ALL_TEMP_FILES:
        if os.path.exists(path):
            os.remove(path)


def create_temporary_file(suffix=''):
    fobj = NamedTemporaryFile(mode='w+', suffix=suffix,
                              prefix='openpyxl.', delete=False)
    filename = fobj.name
    fobj.close()
    ALL_TEMP_FILES.append(filename)
    return filename


class WorksheetWriter:


    def __init__(self, ws, out=None):
        self.ws = ws
        self.ws._hyperlinks = []
        self.ws._comments = []
        if out is None:
            out = create_temporary_file()
        self.out = out
        self._rels = RelationshipList()
        self.xf = self.get_stream()
        next(self.xf) # start generator


    def write_properties(self):
        props = self.ws.sheet_properties
        self.xf.send(props.to_tree())


    def write_dimensions(self):
        """
        Write worksheet size if known
        """
        ref = getattr(self.ws, 'calculate_dimension', None)
        if ref:
            dim = SheetDimension(ref())
            self.xf.send(dim.to_tree())


    def write_format(self):
        self.ws.sheet_format.outlineLevelCol = self.ws.column_dimensions.max_outline
        fmt = self.ws.sheet_format
        self.xf.send(fmt.to_tree())


    def write_views(self):
        views = self.ws.views
        self.xf.send(views.to_tree())


    def write_cols(self):
        cols = self.ws.column_dimensions
        self.xf.send(cols.to_tree())


    def write_top(self):
        """
        Write all elements up to rows:
        properties
        dimensions
        views
        format
        cols
        """
        self.write_properties()
        self.write_dimensions()
        self.write_views()
        self.write_format()
        self.write_cols()


    def rows(self):
        """Return all rows, and any cells that they contain"""
        # order cells by row
        rows = defaultdict(list)
        for (row, col), cell in sorted(self.ws._cells.items()):
            rows[row].append(cell)

        # add empty rows if styling has been applied
        for row in self.ws.row_dimensions.keys() - rows.keys():
            rows[row] = []

        return sorted(rows.items())


    def write_rows(self):
        xf = self.xf.send(True)

        with xf.element("sheetData"):
            for row_idx, row in self.rows():
                self.write_row(xf, row, row_idx)

        self.xf.send(None) # return control to generator


    def write_row(self, xf, row, row_idx):
        attrs = {'r': f"{row_idx}"}
        dims = self.ws.row_dimensions
        attrs.update(dims.get(row_idx, {}))

        with xf.element("row", attrs):

            for cell in row:
                if cell._comment is not None:
                    comment = CommentRecord.from_cell(cell)
                    self.ws._comments.append(comment)
                if (
                    cell._value is None
                    and not cell.has_style
                    and not cell._comment
                    ):
                    continue
                write_cell(xf, self.ws, cell, cell.has_style)


    def write_protection(self):
        prot = self.ws.protection
        if prot:
            self.xf.send(prot.to_tree())


    def write_scenarios(self):
        scenarios = self.ws.scenarios
        if scenarios:
            self.xf.send(scenarios.to_tree())


    def write_filter(self):
        flt = self.ws.auto_filter
        if flt:
            self.xf.send(flt.to_tree())


    def write_sort(self):
        """
        As per discusion with the OOXML Working Group global sort state is not required.
        openpyxl never reads it from existing files
        """
        pass


    def write_merged_cells(self):
        merged = self.ws.merged_cells
        if merged:
            cells = [MergeCell(str(ref)) for ref in self.ws.merged_cells]
            self.xf.send(MergeCells(mergeCell=cells).to_tree())


    def write_formatting(self):
        df = DifferentialStyle()
        wb = self.ws.parent
        for cf in self.ws.conditional_formatting:
            for rule in cf.rules:
                if rule.dxf and rule.dxf != df:
                    rule.dxfId = wb._differential_styles.add(rule.dxf)
            self.xf.send(cf.to_tree())


    def write_validations(self):
        dv = self.ws.data_validations
        if dv:
            self.xf.send(dv.to_tree())


    def write_hyperlinks(self):

        links = self.ws._hyperlinks

        for link in links:
            if link.target:
                rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target)
                self._rels.append(rel)
                link.id = rel.id

        if links:
            self.xf.send(HyperlinkList(links).to_tree())


    def write_print(self):
        print_options = self.ws.print_options
        if print_options:
            self.xf.send(print_options.to_tree())


    def write_margins(self):
        margins = self.ws.page_margins
        if margins:
            self.xf.send(margins.to_tree())


    def write_page(self):
        setup = self.ws.page_setup
        if setup:
            self.xf.send(setup.to_tree())


    def write_header(self):
        hf = self.ws.HeaderFooter
        if hf:
            self.xf.send(hf.to_tree())


    def write_breaks(self):
        brks = (self.ws.row_breaks, self.ws.col_breaks)
        for brk in brks:
            if brk:
                self.xf.send(brk.to_tree())


    def write_drawings(self):
        if self.ws._charts or self.ws._images:
            rel = Relationship(type="drawing", Target="")
            self._rels.append(rel)
            drawing = Related()
            drawing.id = rel.id
            self.xf.send(drawing.to_tree("drawing"))


    def write_legacy(self):
        """
        Comments & VBA controls use VML and require an additional element
        that is no longer in the specification.
        """
        if (self.ws.legacy_drawing is not None or self.ws._comments):
            legacy = Related(id="anysvml")
            self.xf.send(legacy.to_tree("legacyDrawing"))


    def write_tables(self):
        tables = TablePartList()

        for table in self.ws.tables.values():
            if not table.tableColumns:
                table._initialise_columns()
                if table.headerRowCount:
                    try:
                        row = self.ws[table.ref][0]
                        for cell, col in zip(row, table.tableColumns):
                            if cell.data_type != "s":
                                warn("File may not be readable: column headings must be strings.")
                            col.name = str(cell.value)
                    except TypeError:
                        warn("Column headings are missing, file may not be readable")
            rel = Relationship(Type=table._rel_type, Target="")
            self._rels.append(rel)
            table._rel_id = rel.Id
            tables.append(Related(id=rel.Id))

        if tables:
            self.xf.send(tables.to_tree())


    def get_stream(self):
        with xmlfile(self.out) as xf:
            with xf.element("worksheet", xmlns=SHEET_MAIN_NS):
                try:
                    while True:
                        el = (yield)
                        if el is True:
                            yield xf
                        elif el is None: # et_xmlfile chokes
                            continue
                        else:
                            xf.write(el)
                except GeneratorExit:
                    pass


    def write_tail(self):
        """
        Write all elements after the rows
        calc properties
        protection
        protected ranges #
        scenarios
        filters
        sorts # always ignored
        data consolidation #
        custom views #
        merged cells
        phonetic properties #
        conditional formatting
        data validation
        hyperlinks
        print options
        page margins
        page setup
        header
        row breaks
        col breaks
        custom properties #
        cell watches #
        ignored errors #
        smart tags #
        drawing
        drawingHF #
        background #
        OLE objects #
        controls #
        web publishing #
        tables
        """
        self.write_protection()
        self.write_scenarios()
        self.write_filter()
        self.write_merged_cells()
        self.write_formatting()
        self.write_validations()
        self.write_hyperlinks()
        self.write_print()
        self.write_margins()
        self.write_page()
        self.write_header()
        self.write_breaks()
        self.write_drawings()
        self.write_legacy()
        self.write_tables()


    def write(self):
        """
        High level
        """
        self.write_top()
        self.write_rows()
        self.write_tail()
        self.close()


    def close(self):
        """
        Close the context manager
        """
        if self.xf:
            self.xf.close()


    def read(self):
        """
        Close the context manager and return serialised XML
        """
        self.close()
        if isinstance(self.out, BytesIO):
            return self.out.getvalue()
        with open(self.out, "rb") as src:
            out = src.read()

        return out


    def cleanup(self):
        """
        Remove tempfile
        """
        os.remove(self.out)
        ALL_TEMP_FILES.remove(self.out)

''',
    },
    'openpyxl.worksheet.cell_range': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from copy import copy
from operator import attrgetter

from openpyxl.descriptors import Strict
from openpyxl.descriptors import MinMax
from openpyxl.descriptors.sequence import UniqueSequence
from openpyxl.descriptors.serialisable import Serialisable

from openpyxl.utils import (
    range_boundaries,
    range_to_tuple,
    get_column_letter,
    quote_sheetname,
)

class CellRange(Serialisable):
    """
    Represents a range in a sheet: title and coordinates.

    This object is used to perform operations on ranges, like:

    - shift, expand or shrink
    - union/intersection with another sheet range,

    We can check whether a range is:

    - equal or not equal to another,
    - disjoint of another,
    - contained in another.

    We can get:

    - the size of a range.
    - the range bounds (vertices)
    - the coordinates,
    - the string representation,

    """

    min_col = MinMax(min=1, max=18278, expected_type=int)
    min_row = MinMax(min=1, max=1048576, expected_type=int)
    max_col = MinMax(min=1, max=18278, expected_type=int)
    max_row = MinMax(min=1, max=1048576, expected_type=int)


    def __init__(self, range_string=None, min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None):
        if range_string is not None:
            if "!" in range_string:
                title, (min_col, min_row, max_col, max_row) = range_to_tuple(range_string)
            else:
                min_col, min_row, max_col, max_row = range_boundaries(range_string)

        self.min_col = min_col
        self.min_row = min_row
        self.max_col = max_col
        self.max_row = max_row
        self.title = title

        if min_col > max_col:
            fmt = "{max_col} must be greater than {min_col}"
            raise ValueError(fmt.format(min_col=min_col, max_col=max_col))
        if min_row > max_row:
            fmt = "{max_row} must be greater than {min_row}"
            raise ValueError(fmt.format(min_row=min_row, max_row=max_row))


    @property
    def bounds(self):
        """
        Vertices of the range as a tuple
        """
        return self.min_col, self.min_row, self.max_col, self.max_row


    @property
    def coord(self):
        """
        Excel-style representation of the range
        """
        fmt = "{min_col}{min_row}:{max_col}{max_row}"
        if (self.min_col == self.max_col
            and self.min_row == self.max_row):
            fmt = "{min_col}{min_row}"

        return fmt.format(
            min_col=get_column_letter(self.min_col),
            min_row=self.min_row,
            max_col=get_column_letter(self.max_col),
            max_row=self.max_row
        )

    @property
    def rows(self):
        """
        Return cell coordinates as rows
        """
        for row in range(self.min_row, self.max_row+1):
            yield [(row, col) for col in range(self.min_col, self.max_col+1)]


    @property
    def cols(self):
        """
        Return cell coordinates as columns
        """
        for col in range(self.min_col, self.max_col+1):
            yield [(row, col) for row in range(self.min_row, self.max_row+1)]


    @property
    def cells(self):
        from itertools import product
        return product(range(self.min_row, self.max_row+1), range(self.min_col, self.max_col+1))


    def _check_title(self, other):
        """
        Check whether comparisons between ranges are possible.
        Cannot compare ranges from different worksheets
        Skip if the range passed in has no title.
        """
        if not isinstance(other, CellRange):
            raise TypeError(repr(type(other)))

        if other.title and self.title != other.title:
            raise ValueError("Cannot work with ranges from different worksheets")


    def __repr__(self):
        fmt = u"<{cls} {coord}>"
        if self.title:
            fmt = u"<{cls} {title!r}!{coord}>"
        return fmt.format(cls=self.__class__.__name__, title=self.title, coord=self.coord)


    def __hash__(self):
        return hash((self.min_row, self.min_col, self.max_row, self.max_col))


    def __str__(self):
        fmt = "{coord}"
        title = self.title
        if title:
            fmt = u"{title}!{coord}"
            title = quote_sheetname(title)
        return fmt.format(title=title, coord=self.coord)


    def __copy__(self):
        return self.__class__(min_col=self.min_col, min_row=self.min_row,
                              max_col=self.max_col, max_row=self.max_row,
                              title=self.title)


    def shift(self, col_shift=0, row_shift=0):
        """
        Shift the focus of the range according to the shift values (*col_shift*, *row_shift*).

        :type col_shift: int
        :param col_shift: number of columns to be moved by, can be negative
        :type row_shift: int
        :param row_shift: number of rows to be moved by, can be negative
        :raise: :class:`ValueError` if any row or column index < 1
        """

        if (self.min_col + col_shift <= 0
            or self.min_row + row_shift <= 0):
            raise ValueError("Invalid shift value: col_shift={0}, row_shift={1}".format(col_shift, row_shift))
        self.min_col += col_shift
        self.min_row += row_shift
        self.max_col += col_shift
        self.max_row += row_shift


    def __ne__(self, other):
        """
        Test whether the ranges are not equal.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* != *other*.
        """
        try:
            self._check_title(other)
        except ValueError:
            return True

        return (
            other.min_row != self.min_row
            or self.max_row != other.max_row
            or other.min_col != self.min_col
            or self.max_col != other.max_col
        )


    def __eq__(self, other):
        """
        Test whether the ranges are equal.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* == *other*.
        """
        return not self.__ne__(other)


    def issubset(self, other):
        """
        Test whether every cell in this range is also in *other*.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* <= *other*.
        """
        self._check_title(other)

        return other.__superset(self)

    __le__ = issubset


    def __lt__(self, other):
        """
        Test whether *other* contains every cell of this range, and more.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* < *other*.
        """
        return self.__le__(other) and self.__ne__(other)


    def __superset(self, other):
        return (
            (self.min_row <= other.min_row <= other.max_row <= self.max_row)
            and
            (self.min_col <= other.min_col <= other.max_col <= self.max_col)
        )


    def issuperset(self, other):
        """
        Test whether every cell in *other* is in this range.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* >= *other* (or *other* in *range*).
        """
        self._check_title(other)

        return self.__superset(other)

    __ge__ = issuperset


    def __contains__(self, coord):
        """
        Check whether the range contains a particular cell coordinate
        """
        cr = self.__class__(coord)
        return self.__superset(cr)


    def __gt__(self, other):
        """
        Test whether this range contains every cell in *other*, and more.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* > *other*.
        """
        return self.__ge__(other) and self.__ne__(other)


    def isdisjoint(self, other):
        """
        Return ``True`` if this range has no cell in common with *other*.
        Ranges are disjoint if and only if their intersection is the empty range.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range.
        :return: ``True`` if the range has no cells in common with other.
        """
        self._check_title(other)

        # Sort by top-left vertex
        if self.bounds > other.bounds:
            self, other = other, self

        return (self.max_col < other.min_col
                or self.max_row < other.min_row
                or other.max_row < self.min_row)


    def intersection(self, other):
        """
        Return a new range with cells common to this range and *other*

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range.
        :return: the intersecting sheet range.
        :raise: :class:`ValueError` if the *other* range doesn't intersect
            with this range.
        """
        if self.isdisjoint(other):
            raise ValueError("Range {0} doesn't intersect {0}".format(self, other))

        min_row = max(self.min_row, other.min_row)
        max_row = min(self.max_row, other.max_row)
        min_col = max(self.min_col, other.min_col)
        max_col = min(self.max_col, other.max_col)

        return CellRange(min_col=min_col, min_row=min_row, max_col=max_col,
                         max_row=max_row)

    __and__ = intersection


    def union(self, other):
        """
        Return the minimal superset of this range and *other*. This new range
        will contain all cells from this range, *other*, and any additional
        cells required to form a rectangular ``CellRange``.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range.
        :return: a ``CellRange`` that is a superset of this and *other*.
        """
        self._check_title(other)

        min_row = min(self.min_row, other.min_row)
        max_row = max(self.max_row, other.max_row)
        min_col = min(self.min_col, other.min_col)
        max_col = max(self.max_col, other.max_col)
        return CellRange(min_col=min_col, min_row=min_row, max_col=max_col,
                         max_row=max_row, title=self.title)

    __or__ = union


    def __iter__(self):
        """
        For use as a dictionary elsewhere in the library.
        """
        for x in self.__attrs__:
            if x == "title":
                continue
            v = getattr(self, x)
            yield x, v


    def expand(self, right=0, down=0, left=0, up=0):
        """
        Expand the range by the dimensions provided.

        :type right: int
        :param right: expand range to the right by this number of cells
        :type down: int
        :param down: expand range down by this number of cells
        :type left: int
        :param left: expand range to the left by this number of cells
        :type up: int
        :param up: expand range up by this number of cells
        """
        self.min_col -= left
        self.min_row -= up
        self.max_col += right
        self.max_row += down


    def shrink(self, right=0, bottom=0, left=0, top=0):
        """
        Shrink the range by the dimensions provided.

        :type right: int
        :param right: shrink range from the right by this number of cells
        :type down: int
        :param down: shrink range from the top by this number of cells
        :type left: int
        :param left: shrink range from the left by this number of cells
        :type up: int
        :param up: shrink range from the bottom by this number of cells
        """
        self.min_col += left
        self.min_row += top
        self.max_col -= right
        self.max_row -= bottom


    @property
    def size(self):
        """ Return the size of the range as a dictionary of rows and columns. """
        cols = self.max_col + 1 - self.min_col
        rows = self.max_row + 1 - self.min_row
        return {'columns':cols, 'rows':rows}


    @property
    def top(self):
        """A list of cell coordinates that comprise the top of the range"""
        return [(self.min_row, col) for col in range(self.min_col, self.max_col+1)]


    @property
    def bottom(self):
        """A list of cell coordinates that comprise the bottom of the range"""
        return [(self.max_row, col) for col in range(self.min_col, self.max_col+1)]


    @property
    def left(self):
        """A list of cell coordinates that comprise the left-side of the range"""
        return [(row, self.min_col) for row in range(self.min_row, self.max_row+1)]


    @property
    def right(self):
        """A list of cell coordinates that comprise the right-side of the range"""
        return [(row, self.max_col) for row in range(self.min_row, self.max_row+1)]


class MultiCellRange(Strict):


    ranges = UniqueSequence(expected_type=CellRange)


    def __init__(self, ranges=set()):
        if isinstance(ranges, str):
            ranges = [CellRange(r) for r in ranges.split()]
        self.ranges = set(ranges)


    def __contains__(self, coord):
        if isinstance(coord, str):
            coord = CellRange(coord)
        for r in self.ranges:
            if coord <= r:
                return True
        return False


    def __repr__(self):
        ranges = " ".join([str(r) for r in self.sorted()])
        return f"<{self.__class__.__name__} [{ranges}]>"


    def __str__(self):
        ranges = u" ".join([str(r) for r in self.sorted()])
        return ranges


    def __hash__(self):
        return hash(str(self))


    def sorted(self):
        """
        Return a sorted list of items
        """
        return sorted(self.ranges, key=attrgetter('min_col', 'min_row', 'max_col', 'max_row'))


    def add(self, coord):
        """
        Add a cell coordinate or CellRange
        """
        cr = coord
        if isinstance(coord, str):
            cr = CellRange(coord)
        elif not isinstance(coord, CellRange):
            raise ValueError("You can only add CellRanges")
        if cr not in self:
            self.ranges.add(cr)


    def __iadd__(self, coord):
        self.add(coord)
        return self


    def __eq__(self, other):
        if  isinstance(other, str):
            other = self.__class__(other)
        return self.ranges == other.ranges


    def __ne__(self, other):
        return not self == other


    def __bool__(self):
        return bool(self.ranges)


    def remove(self, coord):
        if not isinstance(coord, CellRange):
            coord = CellRange(coord)
        self.ranges.remove(coord)


    def __iter__(self):
        for cr in self.ranges:
            yield cr


    def __copy__(self):
        ranges = {copy(r) for r in self.ranges}
        return MultiCellRange(ranges)

''',
    },
    'openpyxl.worksheet.cell_watch': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Sequence,
    String,
)

# could be done using a nestedSequence

class CellWatch(Serialisable):

    tagname = "cellWatch"

    r = String()

    def __init__(self,
                 r=None,
                ):
        self.r = r


class CellWatches(Serialisable):

    tagname = "cellWatches"

    cellWatch = Sequence(expected_type=CellWatch)

    __elements__ = ('cellWatch',)

    def __init__(self,
                 cellWatch=(),
                ):
        self.cellWatch = cellWatch


""",
    },
    'openpyxl.worksheet.controls': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Bool,
    Integer,
    String,
    Sequence,
)

from openpyxl.descriptors.excel import Relation
from .ole import ObjectAnchor


class ControlProperty(Serialisable):

    tagname = "controlPr"

    anchor = Typed(expected_type=ObjectAnchor, )
    locked = Bool(allow_none=True)
    defaultSize = Bool(allow_none=True)
    _print = Bool(allow_none=True)
    disabled = Bool(allow_none=True)
    recalcAlways = Bool(allow_none=True)
    uiObject = Bool(allow_none=True)
    autoFill = Bool(allow_none=True)
    autoLine = Bool(allow_none=True)
    autoPict = Bool(allow_none=True)
    macro = String(allow_none=True)
    altText = String(allow_none=True)
    linkedCell = String(allow_none=True)
    listFillRange = String(allow_none=True)
    cf = String(allow_none=True)
    id = Relation(allow_none=True)

    __elements__ = ('anchor',)

    def __init__(self,
                 anchor=None,
                 locked=True,
                 defaultSize=True,
                 _print=True,
                 disabled=False,
                 recalcAlways=False,
                 uiObject=False,
                 autoFill=True,
                 autoLine=True,
                 autoPict=True,
                 macro=None,
                 altText=None,
                 linkedCell=None,
                 listFillRange=None,
                 cf='pict',
                 id=None,
                ):
        self.anchor = anchor
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.recalcAlways = recalcAlways
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.autoPict = autoPict
        self.macro = macro
        self.altText = altText
        self.linkedCell = linkedCell
        self.listFillRange = listFillRange
        self.cf = cf
        self.id = id


class Control(Serialisable):

    tagname = "control"

    controlPr = Typed(expected_type=ControlProperty, allow_none=True)
    shapeId = Integer()
    name = String(allow_none=True)

    __elements__ = ('controlPr',)

    def __init__(self,
                 controlPr=None,
                 shapeId=None,
                 name=None,
                ):
        self.controlPr = controlPr
        self.shapeId = shapeId
        self.name = name


class Controls(Serialisable):

    tagname = "controls"

    control = Sequence(expected_type=Control)

    __elements__ = ('control',)

    def __init__(self,
                 control=(),
                ):
        self.control = control


""",
    },
    'openpyxl.worksheet.copier': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

#standard lib imports
from copy import copy

from .worksheet import Worksheet


class WorksheetCopy:
    """
    Copy the values, styles, dimensions, merged cells, margins, and
    print/page setup from one worksheet to another within the same
    workbook.
    """

    def __init__(self, source_worksheet, target_worksheet):
        self.source = source_worksheet
        self.target = target_worksheet
        self._verify_resources()


    def _verify_resources(self):

        if (not isinstance(self.source, Worksheet)
            and not isinstance(self.target, Worksheet)):
            raise TypeError("Can only copy worksheets")

        if self.source is self.target:
            raise ValueError("Cannot copy a worksheet to itself")

        if self.source.parent != self.target.parent:
            raise ValueError('Cannot copy between worksheets from different workbooks')


    def copy_worksheet(self):
        self._copy_cells()
        self._copy_dimensions()

        self.target.sheet_format = copy(self.source.sheet_format)
        self.target.sheet_properties = copy(self.source.sheet_properties)
        self.target.merged_cells = copy(self.source.merged_cells)
        self.target.page_margins = copy(self.source.page_margins)
        self.target.page_setup = copy(self.source.page_setup)
        self.target.print_options = copy(self.source.print_options)


    def _copy_cells(self):
        for (row, col), source_cell  in self.source._cells.items():
            target_cell = self.target.cell(column=col, row=row)

            target_cell._value = source_cell._value
            target_cell.data_type = source_cell.data_type

            if source_cell.has_style:
                target_cell._style = copy(source_cell._style)

            if source_cell.hyperlink:
                target_cell._hyperlink = copy(source_cell.hyperlink)

            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)


    def _copy_dimensions(self):
        for attr in ('row_dimensions', 'column_dimensions'):
            src = getattr(self.source, attr)
            target = getattr(self.target, attr)
            for key, dim in src.items():
                target[key] = copy(dim)
                target[key].worksheet = self.target

''',
    },
    'openpyxl.worksheet.custom': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    String,
    Sequence,
)

# can be done with a nested sequence


class CustomProperty(Serialisable):

    tagname = "customProperty"

    name = String()

    def __init__(self,
                 name=None,
                ):
        self.name = name


class CustomProperties(Serialisable):

    tagname = "customProperties"

    customPr = Sequence(expected_type=CustomProperty)

    __elements__ = ('customPr',)

    def __init__(self,
                 customPr=(),
                ):
        self.customPr = customPr


""",
    },
    'openpyxl.worksheet.datavalidation': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from collections import defaultdict
from itertools import chain
from operator import itemgetter

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    NoneSet,
    String,
    Sequence,
    Alias,
    Integer,
    Convertible,
)
from openpyxl.descriptors.nested import NestedText

from openpyxl.utils import (
    rows_from_range,
    coordinate_to_tuple,
    get_column_letter,
)


def collapse_cell_addresses(cells, input_ranges=()):
    """ Collapse a collection of cell co-ordinates down into an optimal
        range or collection of ranges.

        E.g. Cells A1, A2, A3, B1, B2 and B3 should have the data-validation
        object applied, attempt to collapse down to a single range, A1:B3.

        Currently only collapsing contiguous vertical ranges (i.e. above
        example results in A1:A3 B1:B3).
    """

    ranges = list(input_ranges)

    # convert cell into row, col tuple
    raw_coords = (coordinate_to_tuple(cell) for cell in cells)

    # group by column in order
    grouped_coords = defaultdict(list)
    for row, col in sorted(raw_coords, key=itemgetter(1)):
        grouped_coords[col].append(row)

    # create range string from first and last row in column
    for col, cells in grouped_coords.items():
        col = get_column_letter(col)
        fmt = "{0}{1}:{2}{3}"
        if len(cells) == 1:
            fmt = "{0}{1}"
        r = fmt.format(col, min(cells), col, max(cells))
        ranges.append(r)

    return " ".join(ranges)


def expand_cell_ranges(range_string):
    """
    Expand cell ranges to a sequence of addresses.
    Reverse of collapse_cell_addresses
    Eg. converts "A1:A2 B1:B2" to (A1, A2, B1, B2)
    """
    # expand ranges to rows and then flatten
    rows = (rows_from_range(rs) for rs in range_string.split()) # list of rows
    cells = (chain(*row) for row in rows) # flatten rows
    return set(chain(*cells))


from .cell_range import MultiCellRange


class DataValidation(Serialisable):

    tagname = "dataValidation"

    sqref = Convertible(expected_type=MultiCellRange)
    cells = Alias("sqref")
    ranges = Alias("sqref")

    showDropDown = Bool(allow_none=True)
    hide_drop_down = Alias('showDropDown')
    showInputMessage = Bool(allow_none=True)
    showErrorMessage = Bool(allow_none=True)
    allowBlank = Bool(allow_none=True)
    allow_blank = Alias('allowBlank')

    errorTitle = String(allow_none = True)
    error = String(allow_none = True)
    promptTitle = String(allow_none = True)
    prompt = String(allow_none = True)
    formula1 = NestedText(allow_none=True, expected_type=str)
    formula2 = NestedText(allow_none=True, expected_type=str)

    type = NoneSet(values=("whole", "decimal", "list", "date", "time",
                           "textLength", "custom"))
    errorStyle = NoneSet(values=("stop", "warning", "information"))
    imeMode = NoneSet(values=("noControl", "off", "on", "disabled",
                              "hiragana", "fullKatakana", "halfKatakana", "fullAlpha","halfAlpha",
                              "fullHangul", "halfHangul"))
    operator = NoneSet(values=("between", "notBetween", "equal", "notEqual",
                               "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"))
    validation_type = Alias('type')

    def __init__(self,
                 type=None,
                 formula1=None,
                 formula2=None,
                 showErrorMessage=False,
                 showInputMessage=False,
                 showDropDown=False,
                 allowBlank=False,
                 sqref=(),
                 promptTitle=None,
                 errorStyle=None,
                 error=None,
                 prompt=None,
                 errorTitle=None,
                 imeMode=None,
                 operator=None,
                 allow_blank=None,
                 ):
        self.sqref = sqref
        self.showDropDown = showDropDown
        self.imeMode = imeMode
        self.operator = operator
        self.formula1 = formula1
        self.formula2 = formula2
        if allow_blank is not None:
            allowBlank = allow_blank
        self.allowBlank = allowBlank
        self.showErrorMessage = showErrorMessage
        self.showInputMessage = showInputMessage
        self.type = type
        self.promptTitle = promptTitle
        self.errorStyle = errorStyle
        self.error = error
        self.prompt = prompt
        self.errorTitle = errorTitle


    def add(self, cell):
        """Adds a cell or cell coordinate to this validator"""
        if hasattr(cell, "coordinate"):
            cell = cell.coordinate
        self.sqref += cell


    def __contains__(self, cell):
        if hasattr(cell, "coordinate"):
            cell = cell.coordinate
        return cell in self.sqref


class DataValidationList(Serialisable):

    tagname = "dataValidations"

    disablePrompts = Bool(allow_none=True)
    xWindow = Integer(allow_none=True)
    yWindow = Integer(allow_none=True)
    dataValidation = Sequence(expected_type=DataValidation)

    __elements__ = ('dataValidation',)
    __attrs__ = ('disablePrompts', 'xWindow', 'yWindow', 'count')

    def __init__(self,
                 disablePrompts=None,
                 xWindow=None,
                 yWindow=None,
                 count=None,
                 dataValidation=(),
                ):
        self.disablePrompts = disablePrompts
        self.xWindow = xWindow
        self.yWindow = yWindow
        self.dataValidation = dataValidation


    @property
    def count(self):
        return len(self)


    def __len__(self):
        return len(self.dataValidation)


    def append(self, dv):
        self.dataValidation.append(dv)


    def to_tree(self, tagname=None):
        """
        Need to skip validations that have no cell ranges
        """
        ranges = self.dataValidation # copy
        self.dataValidation = [r for r in self.dataValidation if bool(r.sqref)]
        xml = super().to_tree(tagname)
        self.dataValidation = ranges
        return xml

''',
    },
    'openpyxl.worksheet.dimensions': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from copy import copy

from openpyxl.compat import safe_string
from openpyxl.utils import (
    get_column_letter,
    get_column_interval,
    column_index_from_string,
    range_boundaries,
)
from openpyxl.utils.units import DEFAULT_COLUMN_WIDTH
from openpyxl.descriptors import (
    Integer,
    Float,
    Bool,
    Strict,
    String,
    Alias,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles.styleable import StyleableObject
from openpyxl.utils.bound_dictionary import BoundDictionary
from openpyxl.xml.functions import Element


class Dimension(Strict, StyleableObject):
    """Information about the display properties of a row or column."""
    __fields__ = ('hidden',
                 'outlineLevel',
                 'collapsed',)

    index = Integer()
    hidden = Bool()
    outlineLevel = Integer(allow_none=True)
    outline_level = Alias('outlineLevel')
    collapsed = Bool()
    style = Alias('style_id')


    def __init__(self, index, hidden, outlineLevel,
                 collapsed, worksheet, visible=True, style=None):
        super().__init__(sheet=worksheet, style_array=style)
        self.index = index
        self.hidden = hidden
        self.outlineLevel = outlineLevel
        self.collapsed = collapsed


    def __iter__(self):
        for key in self.__fields__:
            value = getattr(self, key, None)
            if value:
                yield key, safe_string(value)


    def __copy__(self):
        cp = self.__new__(self.__class__)
        attrib = self.__dict__
        attrib['worksheet'] = self.parent
        cp.__init__(**attrib)
        cp._style = copy(self._style)
        return cp


    def __repr__(self):
        return f"<{self.__class__.__name__} Instance, Attributes={dict(self)}>"


class RowDimension(Dimension):
    """Information about the display properties of a row."""

    __fields__ = Dimension.__fields__ + ('ht', 'customFormat', 'customHeight', 's',
                                         'thickBot', 'thickTop')
    r = Alias('index')
    s = Alias('style_id')
    ht = Float(allow_none=True)
    height = Alias('ht')
    thickBot = Bool()
    thickTop = Bool()

    def __init__(self,
                 worksheet,
                 index=0,
                 ht=None,
                 customHeight=None, # do not write
                 s=None,
                 customFormat=None, # do not write
                 hidden=False,
                 outlineLevel=0,
                 outline_level=None,
                 collapsed=False,
                 visible=None,
                 height=None,
                 r=None,
                 spans=None,
                 thickBot=None,
                 thickTop=None,
                 **kw
                 ):
        if r is not None:
            index = r
        if height is not None:
            ht = height
        self.ht = ht
        if visible is not None:
            hidden = not visible
        if outline_level is not None:
            outlineLevel = outline_level
        self.thickBot = thickBot
        self.thickTop = thickTop
        super().__init__(index, hidden, outlineLevel,
                                           collapsed, worksheet, style=s)

    @property
    def customFormat(self):
        """Always true if there is a style for the row"""
        return self.has_style

    @property
    def customHeight(self):
        """Always true if there is a height for the row"""
        return self.ht is not None


class ColumnDimension(Dimension):
    """Information about the display properties of a column."""

    width = Float()
    bestFit = Bool()
    auto_size = Alias('bestFit')
    index = String()
    min = Integer(allow_none=True)
    max = Integer(allow_none=True)
    collapsed = Bool()

    __fields__ = Dimension.__fields__ + ('width', 'bestFit', 'customWidth', 'style',
                                         'min', 'max')

    def __init__(self,
                 worksheet,
                 index='A',
                 width=DEFAULT_COLUMN_WIDTH,
                 bestFit=False,
                 hidden=False,
                 outlineLevel=0,
                 outline_level=None,
                 collapsed=False,
                 style=None,
                 min=None,
                 max=None,
                 customWidth=False, # do not write
                 visible=None,
                 auto_size=None,):
        self.width = width
        self.min = min
        self.max = max
        if visible is not None:
            hidden = not visible
        if auto_size is not None:
            bestFit = auto_size
        self.bestFit = bestFit
        if outline_level is not None:
            outlineLevel = outline_level
        self.collapsed = collapsed
        super().__init__(index, hidden, outlineLevel,
                                              collapsed, worksheet, style=style)


    @property
    def customWidth(self):
        """Always true if there is a width for the column"""
        return bool(self.width)


    def reindex(self):
        """
        Set boundaries for column definition
        """
        if not all([self.min, self.max]):
            self.min = self.max = column_index_from_string(self.index)

    @property
    def range(self):
        """Return the range of cells actually covered"""
        return f"{get_column_letter(self.min)}:{get_column_letter(self.max)}"


    def to_tree(self):
        attrs = dict(self)
        if attrs.keys() != {'min', 'max'}:
            return Element("col", **attrs)


class DimensionHolder(BoundDictionary):
    """
    Allow columns to be grouped
    """

    def __init__(self, worksheet, reference="index", default_factory=None):
        self.worksheet = worksheet
        self.max_outline = None
        self.default_factory = default_factory
        super().__init__(reference, default_factory)


    def group(self, start, end=None, outline_level=1, hidden=False):
        """allow grouping a range of consecutive rows or columns together

        :param start: first row or column to be grouped (mandatory)
        :param end: last row or column to be grouped (optional, default to start)
        :param outline_level: outline level
        :param hidden: should the group be hidden on workbook open or not
        """
        if end is None:
            end = start

        if isinstance(self.default_factory(), ColumnDimension):
            new_dim = self[start]
            new_dim.outline_level = outline_level
            new_dim.hidden = hidden
            work_sequence = get_column_interval(start, end)[1:]
            for column_letter in work_sequence:
                if column_letter in self:
                    del self[column_letter]
            new_dim.min, new_dim.max = map(column_index_from_string, (start, end))
        elif isinstance(self.default_factory(), RowDimension):
            for el in range(start, end + 1):
                new_dim = self.worksheet.row_dimensions[el]
                new_dim.outline_level = outline_level
                new_dim.hidden = hidden


    def to_tree(self):

        def sorter(value):
            value.reindex()
            return value.min

        el = Element('cols')
        outlines = set()

        for col in sorted(self.values(), key=sorter):
            obj = col.to_tree()
            if obj is not None:
                outlines.add(col.outlineLevel)
                el.append(obj)

        if outlines:
            self.max_outline = max(outlines)

        if len(el):
            return el # must have at least one child


class SheetFormatProperties(Serialisable):

    tagname = "sheetFormatPr"

    baseColWidth = Integer(allow_none=True)
    defaultColWidth = Float(allow_none=True)
    defaultRowHeight = Float()
    customHeight = Bool(allow_none=True)
    zeroHeight = Bool(allow_none=True)
    thickTop = Bool(allow_none=True)
    thickBottom = Bool(allow_none=True)
    outlineLevelRow = Integer(allow_none=True)
    outlineLevelCol = Integer(allow_none=True)

    def __init__(self,
                 baseColWidth=8, #according to spec
                 defaultColWidth=None,
                 defaultRowHeight=15,
                 customHeight=None,
                 zeroHeight=None,
                 thickTop=None,
                 thickBottom=None,
                 outlineLevelRow=None,
                 outlineLevelCol=None,
                ):
        self.baseColWidth = baseColWidth
        self.defaultColWidth = defaultColWidth
        self.defaultRowHeight = defaultRowHeight
        self.customHeight = customHeight
        self.zeroHeight = zeroHeight
        self.thickTop = thickTop
        self.thickBottom = thickBottom
        self.outlineLevelRow = outlineLevelRow
        self.outlineLevelCol = outlineLevelCol


class SheetDimension(Serialisable):

    tagname = "dimension"

    ref = String()

    def __init__(self,
                 ref=None,
                ):
        self.ref = ref


    @property
    def boundaries(self):
        return range_boundaries(self.ref)

''',
    },
    'openpyxl.worksheet.drawing': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.excel import Relation


class Drawing(Serialisable):

    tagname = "drawing"

    id = Relation()

    def __init__(self, id=None):
        self.id = id

""",
    },
    'openpyxl.worksheet.errors': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Bool,
    Sequence,
)
from openpyxl.descriptors.excel import CellRange


class Extension(Serialisable):

    tagname = "extension"

    uri = String(allow_none=True)

    def __init__(self,
                 uri=None,
                ):
        self.uri = uri


class ExtensionList(Serialisable):

    tagname = "extensionList"

    # uses element group EG_ExtensionList
    ext = Sequence(expected_type=Extension)

    __elements__ = ('ext',)

    def __init__(self,
                 ext=(),
                ):
        self.ext = ext


class IgnoredError(Serialisable):

    tagname = "ignoredError"

    sqref = CellRange
    evalError = Bool(allow_none=True)
    twoDigitTextYear = Bool(allow_none=True)
    numberStoredAsText = Bool(allow_none=True)
    formula = Bool(allow_none=True)
    formulaRange = Bool(allow_none=True)
    unlockedFormula = Bool(allow_none=True)
    emptyCellReference = Bool(allow_none=True)
    listDataValidation = Bool(allow_none=True)
    calculatedColumn = Bool(allow_none=True)

    def __init__(self,
                 sqref=None,
                 evalError=False,
                 twoDigitTextYear=False,
                 numberStoredAsText=False,
                 formula=False,
                 formulaRange=False,
                 unlockedFormula=False,
                 emptyCellReference=False,
                 listDataValidation=False,
                 calculatedColumn=False,
                ):
        self.sqref = sqref
        self.evalError = evalError
        self.twoDigitTextYear = twoDigitTextYear
        self.numberStoredAsText = numberStoredAsText
        self.formula = formula
        self.formulaRange = formulaRange
        self.unlockedFormula = unlockedFormula
        self.emptyCellReference = emptyCellReference
        self.listDataValidation = listDataValidation
        self.calculatedColumn = calculatedColumn


class IgnoredErrors(Serialisable):

    tagname = "ignoredErrors"

    ignoredError = Sequence(expected_type=IgnoredError)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ignoredError', 'extLst')

    def __init__(self,
                 ignoredError=(),
                 extLst=None,
                ):
        self.ignoredError = ignoredError
        self.extLst = extLst


""",
    },
    'openpyxl.worksheet.filters': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import re

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    Set,
    Float,
    DateTime,
    NoneSet,
    Bool,
    Integer,
    String,
    Sequence,
    MinMax,
)
from openpyxl.descriptors.excel import ExtensionList, CellRange
from openpyxl.descriptors.sequence import ValueSequence
from openpyxl.utils import absolute_coordinate


class SortCondition(Serialisable):

    tagname = "sortCondition"

    descending = Bool(allow_none=True)
    sortBy = NoneSet(values=(['value', 'cellColor', 'fontColor', 'icon']))
    ref = CellRange()
    customList = String(allow_none=True)
    dxfId = Integer(allow_none=True)
    iconSet = NoneSet(values=(['3Arrows', '3ArrowsGray', '3Flags',
                           '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
                           '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
                           '5Arrows', '5ArrowsGray', '5Rating', '5Quarters']))
    iconId = Integer(allow_none=True)

    def __init__(self,
                 ref=None,
                 descending=None,
                 sortBy=None,
                 customList=None,
                 dxfId=None,
                 iconSet=None,
                 iconId=None,
                ):
        self.descending = descending
        self.sortBy = sortBy
        self.ref = ref
        self.customList = customList
        self.dxfId = dxfId
        self.iconSet = iconSet
        self.iconId = iconId


class SortState(Serialisable):

    tagname = "sortState"

    columnSort = Bool(allow_none=True)
    caseSensitive = Bool(allow_none=True)
    sortMethod = NoneSet(values=(['stroke', 'pinYin']))
    ref = CellRange()
    sortCondition = Sequence(expected_type=SortCondition, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('sortCondition',)

    def __init__(self,
                 columnSort=None,
                 caseSensitive=None,
                 sortMethod=None,
                 ref=None,
                 sortCondition=(),
                 extLst=None,
                ):
        self.columnSort = columnSort
        self.caseSensitive = caseSensitive
        self.sortMethod = sortMethod
        self.ref = ref
        self.sortCondition = sortCondition


    def __bool__(self):
        return self.ref is not None



class IconFilter(Serialisable):

    tagname = "iconFilter"

    iconSet = Set(values=(['3Arrows', '3ArrowsGray', '3Flags',
                           '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
                           '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
                           '5Arrows', '5ArrowsGray', '5Rating', '5Quarters']))
    iconId = Integer(allow_none=True)

    def __init__(self,
                 iconSet=None,
                 iconId=None,
                ):
        self.iconSet = iconSet
        self.iconId = iconId


class ColorFilter(Serialisable):

    tagname = "colorFilter"

    dxfId = Integer(allow_none=True)
    cellColor = Bool(allow_none=True)

    def __init__(self,
                 dxfId=None,
                 cellColor=None,
                ):
        self.dxfId = dxfId
        self.cellColor = cellColor


class DynamicFilter(Serialisable):

    tagname = "dynamicFilter"

    type = Set(values=(['null', 'aboveAverage', 'belowAverage', 'tomorrow',
                        'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth',
                        'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter',
                        'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4',
                        'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11',
                        'M12']))
    val = Float(allow_none=True)
    valIso = DateTime(allow_none=True)
    maxVal = Float(allow_none=True)
    maxValIso = DateTime(allow_none=True)

    def __init__(self,
                 type=None,
                 val=None,
                 valIso=None,
                 maxVal=None,
                 maxValIso=None,
                ):
        self.type = type
        self.val = val
        self.valIso = valIso
        self.maxVal = maxVal
        self.maxValIso = maxValIso


class CustomFilter(Serialisable):

    tagname = "customFilter"

    val = String()
    operator = Set(values=['equal', 'lessThan', 'lessThanOrEqual',
                           'notEqual', 'greaterThanOrEqual', 'greaterThan'])

    def __init__(self, operator="equal", val=None):
        self.operator = operator
        self.val = val


    def _get_subtype(self):
        if self.val == " ":
            subtype = BlankFilter
        else:
            try:
                float(self.val)
                subtype = NumberFilter
            except ValueError:
                subtype = StringFilter
        return subtype


    def convert(self):
        """Convert to more specific filter"""
        typ = self._get_subtype()
        if typ in (BlankFilter, NumberFilter):
            return typ(**dict(self))

        operator, term = StringFilter._guess_operator(self.val)
        flt = StringFilter(operator, term)
        if self.operator == "notEqual":
            flt.exclude = True
        return flt


class BlankFilter(CustomFilter):
    """
    Exclude blanks
    """

    __attrs__ = ("operator", "val")

    def __init__(self, **kw):
        pass


    @property
    def operator(self):
        return "notEqual"


    @property
    def val(self):
        return " "


class NumberFilter(CustomFilter):


    operator = Set(values=
                   ['equal', 'lessThan', 'lessThanOrEqual',
                    'notEqual', 'greaterThanOrEqual', 'greaterThan'])
    val = Float()

    def __init__(self, operator="equal", val=None):
        self.operator = operator
        self.val = val


string_format_mapping = {
    "contains": "*{}*",
    "startswith": "{}*",
    "endswith": "*{}",
    "wildcard":  "{}",
}


class StringFilter(CustomFilter):

    operator = Set(values=['contains', 'startswith', 'endswith', 'wildcard']
                   )
    val = String()
    exclude = Bool()


    def __init__(self, operator="contains", val=None, exclude=False):
        self.operator = operator
        self.val = val
        self.exclude = exclude


    def _escape(self):
        """Escape wildcards ~, * ? when serialising"""
        if self.operator == "wildcard":
            return self.val
        return re.sub(r"~|\*|\?", r"~\g<0>", self.val)


    @staticmethod
    def _unescape(value):
        """
        Unescape value
        """
        return re.sub(r"~(?P<op>[~*?])", r"\g<op>", value)


    @staticmethod
    def _guess_operator(value):
        value = StringFilter._unescape(value)
        endswith = r"^(?P<endswith>\*)(?P<term>[^\*\?]*$)"
        startswith = r"^(?P<term>[^\*\?]*)(?P<startswith>\*)$"
        contains = r"^(?P<contains>\*)(?P<term>[^\*\?]*)\*$"
        d = {"wildcard": True, "term": value}
        for pat in [contains, startswith, endswith]:
            m = re.match(pat, value)
            if m:
                d = m.groupdict()

        term = d.pop("term")
        op = list(d)[0]
        return op, term


    def to_tree(self, tagname=None, idx=None, namespace=None):
        fmt = string_format_mapping[self.operator]
        op = self.exclude and "notEqual" or "equal"
        value = fmt.format(self._escape())
        flt = CustomFilter(op, value)
        return flt.to_tree(tagname, idx, namespace)


class CustomFilters(Serialisable):

    tagname = "customFilters"

    _and = Bool(allow_none=True)
    customFilter = Sequence(expected_type=CustomFilter) # min 1, max 2

    __elements__ = ('customFilter',)

    def __init__(self,
                 _and=None,
                 customFilter=(),
                ):
        self._and = _and
        self.customFilter = customFilter


class Top10(Serialisable):

    tagname = "top10"

    top = Bool(allow_none=True)
    percent = Bool(allow_none=True)
    val = Float()
    filterVal = Float(allow_none=True)

    def __init__(self,
                 top=None,
                 percent=None,
                 val=None,
                 filterVal=None,
                ):
        self.top = top
        self.percent = percent
        self.val = val
        self.filterVal = filterVal


class DateGroupItem(Serialisable):

    tagname = "dateGroupItem"

    year = Integer()
    month = MinMax(min=1, max=12, allow_none=True)
    day = MinMax(min=1, max=31, allow_none=True)
    hour = MinMax(min=0, max=23, allow_none=True)
    minute = MinMax(min=0, max=59, allow_none=True)
    second = Integer(min=0, max=59, allow_none=True)
    dateTimeGrouping = Set(values=(['year', 'month', 'day', 'hour', 'minute',
                                    'second']))

    def __init__(self,
                 year=None,
                 month=None,
                 day=None,
                 hour=None,
                 minute=None,
                 second=None,
                 dateTimeGrouping=None,
                ):
        self.year = year
        self.month = month
        self.day = day
        self.hour = hour
        self.minute = minute
        self.second = second
        self.dateTimeGrouping = dateTimeGrouping


class Filters(Serialisable):

    tagname = "filters"

    blank = Bool(allow_none=True)
    calendarType = NoneSet(values=["gregorian","gregorianUs",
                                   "gregorianMeFrench","gregorianArabic", "hijri","hebrew",
                                   "taiwan","japan", "thai","korea",
                                   "saka","gregorianXlitEnglish","gregorianXlitFrench"])
    filter = ValueSequence(expected_type=str)
    dateGroupItem = Sequence(expected_type=DateGroupItem, allow_none=True)

    __elements__ = ('filter', 'dateGroupItem')

    def __init__(self,
                 blank=None,
                 calendarType=None,
                 filter=(),
                 dateGroupItem=(),
                ):
        self.blank = blank
        self.calendarType = calendarType
        self.filter = filter
        self.dateGroupItem = dateGroupItem


class FilterColumn(Serialisable):

    tagname = "filterColumn"

    colId = Integer()
    col_id = Alias('colId')
    hiddenButton = Bool(allow_none=True)
    showButton = Bool(allow_none=True)
    # some elements are choice
    filters = Typed(expected_type=Filters, allow_none=True)
    top10 = Typed(expected_type=Top10, allow_none=True)
    customFilters = Typed(expected_type=CustomFilters, allow_none=True)
    dynamicFilter = Typed(expected_type=DynamicFilter, allow_none=True)
    colorFilter = Typed(expected_type=ColorFilter, allow_none=True)
    iconFilter = Typed(expected_type=IconFilter, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('filters', 'top10', 'customFilters', 'dynamicFilter',
                    'colorFilter', 'iconFilter')

    def __init__(self,
                 colId=None,
                 hiddenButton=False,
                 showButton=True,
                 filters=None,
                 top10=None,
                 customFilters=None,
                 dynamicFilter=None,
                 colorFilter=None,
                 iconFilter=None,
                 extLst=None,
                 blank=None,
                 vals=None,
                ):
        self.colId = colId
        self.hiddenButton = hiddenButton
        self.showButton = showButton
        self.filters = filters
        self.top10 = top10
        self.customFilters = customFilters
        self.dynamicFilter = dynamicFilter
        self.colorFilter = colorFilter
        self.iconFilter = iconFilter
        if blank is not None and self.filters:
            self.filters.blank = blank
        if vals is not None and self.filters:
            self.filters.filter = vals


class AutoFilter(Serialisable):

    tagname = "autoFilter"

    ref = CellRange()
    filterColumn = Sequence(expected_type=FilterColumn, allow_none=True)
    sortState = Typed(expected_type=SortState, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('filterColumn', 'sortState')

    def __init__(self,
                 ref=None,
                 filterColumn=(),
                 sortState=None,
                 extLst=None,
                ):
        self.ref = ref
        self.filterColumn = filterColumn
        self.sortState = sortState


    def __bool__(self):
        return self.ref is not None


    def __str__(self):
        return absolute_coordinate(self.ref)


    def add_filter_column(self, col_id, vals, blank=False):
        """
        Add row filter for specified column.

        :param col_id: Zero-origin column id. 0 means first column.
        :type  col_id: int
        :param vals: Value list to show.
        :type  vals: str[]
        :param blank: Show rows that have blank cell if True (default=``False``)
        :type  blank: bool
        """
        self.filterColumn.append(FilterColumn(colId=col_id, filters=Filters(blank=blank, filter=vals)))


    def add_sort_condition(self, ref, descending=False):
        """
        Add sort condition for cpecified range of cells.

        :param ref: range of the cells (e.g. 'A2:A150')
        :type  ref: string, is the same as that of the filter
        :param descending: Descending sort order (default=``False``)
        :type  descending: bool
        """
        cond = SortCondition(ref, descending)
        if self.sortState is None:
            self.sortState = SortState(ref=self.ref)
        self.sortState.sortCondition.append(cond)

''',
    },
    'openpyxl.worksheet.formula': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.compat import safe_string

class DataTableFormula:


    t = "dataTable"

    def __init__(self,
                 ref,
                 ca=False,
                 dt2D=False,
                 dtr=False,
                 r1=None,
                 r2=None,
                 del1=False,
                 del2=False,
                 **kw):
        self.ref = ref
        self.ca = ca
        self.dt2D = dt2D
        self.dtr = dtr
        self.r1 = r1
        self.r2 = r2
        self.del1 = del1
        self.del2 = del2


    def __iter__(self):
        for k in ["t", "ref", "dt2D", "dtr", "r1", "r2", "del1", "del2", "ca"]:
            v = getattr(self, k)
            if v:
                yield k, safe_string(v)


class ArrayFormula:

    t = "array"


    def __init__(self, ref, text=None):
        self.ref = ref
        self.text = text


    def __iter__(self):
        for k in ["t", "ref"]:
            v = getattr(self, k)
            if v:
                yield k, safe_string(v)

""",
    },
    'openpyxl.worksheet.header_footer': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

# Simplified implementation of headers and footers: let worksheets have separate items

import re
from warnings import warn

from openpyxl.descriptors import (
    Alias,
    Bool,
    Strict,
    String,
    Integer,
    MatchPattern,
    Typed,
)
from openpyxl.descriptors.serialisable import Serialisable


from openpyxl.xml.functions import Element
from openpyxl.utils.escape import escape, unescape


FONT_PATTERN = '&"(?P<font>.+)"'
COLOR_PATTERN  = "&K(?P<color>[A-F0-9]{6})"
SIZE_REGEX = r"&(?P<size>\d+\s?)"
FORMAT_REGEX = re.compile("{0}|{1}|{2}".format(FONT_PATTERN, COLOR_PATTERN,
                                               SIZE_REGEX)
                          )

def _split_string(text):
    """
    Split the combined (decoded) string into left, center and right parts

    # See http://stackoverflow.com/questions/27711175/regex-with-multiple-optional-groups for discussion
    """

    ITEM_REGEX = re.compile("""
    (&L(?P<left>.+?))?
    (&C(?P<center>.+?))?
    (&R(?P<right>.+?))?
    $""", re.VERBOSE | re.DOTALL)

    m = ITEM_REGEX.match(text)
    try:
        parts = m.groupdict()
    except AttributeError:
        warn("""Cannot parse header or footer so it will be ignored""")
        parts = {'left':'', 'right':'', 'center':''}
    return parts


class _HeaderFooterPart(Strict):

    """
    Individual left/center/right header/footer part

    Do not use directly.

    Header & Footer ampersand codes:

    * &A   Inserts the worksheet name
    * &B   Toggles bold
    * &D or &[Date]   Inserts the current date
    * &E   Toggles double-underline
    * &F or &[File]   Inserts the workbook name
    * &I   Toggles italic
    * &N or &[Pages]   Inserts the total page count
    * &S   Toggles strikethrough
    * &T   Inserts the current time
    * &[Tab]   Inserts the worksheet name
    * &U   Toggles underline
    * &X   Toggles superscript
    * &Y   Toggles subscript
    * &P or &[Page]   Inserts the current page number
    * &P+n   Inserts the page number incremented by n
    * &P-n   Inserts the page number decremented by n
    * &[Path]   Inserts the workbook path
    * &&   Escapes the ampersand character
    * &"fontname"   Selects the named font
    * &nn   Selects the specified 2-digit font point size

    Colours are in RGB Hex
    """

    text = String(allow_none=True)
    font = String(allow_none=True)
    size = Integer(allow_none=True)
    RGB = ("^[A-Fa-f0-9]{6}$")
    color = MatchPattern(allow_none=True, pattern=RGB)


    def __init__(self, text=None, font=None, size=None, color=None):
        self.text = text
        self.font = font
        self.size = size
        self.color = color


    def __str__(self):
        """
        Convert to Excel HeaderFooter miniformat minus position
        """
        fmt = []
        if self.font:
            fmt.append(u'&"{0}"'.format(self.font))
        if self.size:
            fmt.append("&{0} ".format(self.size))
        if self.color:
            fmt.append("&K{0}".format(self.color))
        return u"".join(fmt + [self.text])

    def __bool__(self):
        return bool(self.text)



    @classmethod
    def from_str(cls, text):
        """
        Convert from miniformat to object
        """
        keys = ('font', 'color', 'size')
        kw = dict((k, v) for match in FORMAT_REGEX.findall(text)
                  for k, v in zip(keys, match) if v)

        kw['text'] = FORMAT_REGEX.sub('', text)

        return cls(**kw)


class HeaderFooterItem(Strict):
    """
    Header or footer item

    """

    left = Typed(expected_type=_HeaderFooterPart)
    center = Typed(expected_type=_HeaderFooterPart)
    centre = Alias("center")
    right = Typed(expected_type=_HeaderFooterPart)

    __keys = ('L', 'C', 'R')


    def __init__(self, left=None, right=None, center=None):
        if left is None:
            left = _HeaderFooterPart()
        self.left = left
        if center is None:
            center = _HeaderFooterPart()
        self.center = center
        if right is None:
            right = _HeaderFooterPart()
        self.right = right


    def __str__(self):
        """
        Pack parts into a single string
        """
        TRANSFORM = {'&[Tab]': '&A', '&[Pages]': '&N', '&[Date]': '&D',
                     '&[Path]': '&Z', '&[Page]': '&P', '&[Time]': '&T', '&[File]': '&F',
                     '&[Picture]': '&G'}

        # escape keys and create regex
        SUBS_REGEX = re.compile("|".join(["({0})".format(re.escape(k))
                                          for k in TRANSFORM]))

        def replace(match):
            """
            Callback for re.sub
            Replace expanded control with mini-format equivalent
            """
            sub = match.group(0)
            return TRANSFORM[sub]

        txt = []
        for key, part in zip(
            self.__keys, [self.left, self.center, self.right]):
            if part.text is not None:
                txt.append(u"&{0}{1}".format(key, str(part)))
        txt = "".join(txt)
        txt = SUBS_REGEX.sub(replace, txt)
        return escape(txt)


    def __bool__(self):
        return any([self.left, self.center, self.right])



    def to_tree(self, tagname):
        """
        Return as XML node
        """
        el = Element(tagname)
        el.text = str(self)
        return el


    @classmethod
    def from_tree(cls, node):
        if node.text:
            text = unescape(node.text)
            parts = _split_string(text)
            for k, v in parts.items():
                if v is not None:
                    parts[k] = _HeaderFooterPart.from_str(v)
            self = cls(**parts)
            return self


class HeaderFooter(Serialisable):

    tagname = "headerFooter"

    differentOddEven = Bool(allow_none=True)
    differentFirst = Bool(allow_none=True)
    scaleWithDoc = Bool(allow_none=True)
    alignWithMargins = Bool(allow_none=True)
    oddHeader = Typed(expected_type=HeaderFooterItem, allow_none=True)
    oddFooter = Typed(expected_type=HeaderFooterItem, allow_none=True)
    evenHeader = Typed(expected_type=HeaderFooterItem, allow_none=True)
    evenFooter = Typed(expected_type=HeaderFooterItem, allow_none=True)
    firstHeader = Typed(expected_type=HeaderFooterItem, allow_none=True)
    firstFooter = Typed(expected_type=HeaderFooterItem, allow_none=True)

    __elements__ = ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter")

    def __init__(self,
                 differentOddEven=None,
                 differentFirst=None,
                 scaleWithDoc=None,
                 alignWithMargins=None,
                 oddHeader=None,
                 oddFooter=None,
                 evenHeader=None,
                 evenFooter=None,
                 firstHeader=None,
                 firstFooter=None,
                ):
        self.differentOddEven = differentOddEven
        self.differentFirst = differentFirst
        self.scaleWithDoc = scaleWithDoc
        self.alignWithMargins = alignWithMargins
        if oddHeader is None:
            oddHeader = HeaderFooterItem()
        self.oddHeader = oddHeader
        if oddFooter is None:
            oddFooter = HeaderFooterItem()
        self.oddFooter = oddFooter
        if evenHeader is None:
            evenHeader = HeaderFooterItem()
        self.evenHeader = evenHeader
        if evenFooter is None:
            evenFooter = HeaderFooterItem()
        self.evenFooter = evenFooter
        if firstHeader is None:
            firstHeader = HeaderFooterItem()
        self.firstHeader = firstHeader
        if firstFooter is None:
            firstFooter = HeaderFooterItem()
        self.firstFooter = firstFooter


    def __bool__(self):
        parts = [getattr(self, attr) for attr in self.__attrs__ + self.__elements__]
        return any(parts)


''',
    },
    'openpyxl.worksheet.hyperlink': {
        'is_package': False,
        'source': r"""
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    String,
    Sequence,
)
from openpyxl.descriptors.excel import Relation


class Hyperlink(Serialisable):

    tagname = "hyperlink"

    ref = String()
    location = String(allow_none=True)
    tooltip = String(allow_none=True)
    display = String(allow_none=True)
    id = Relation()
    target = String(allow_none=True)

    __attrs__ = ("ref", "location", "tooltip", "display", "id")

    def __init__(self,
                 ref=None,
                 location=None,
                 tooltip=None,
                 display=None,
                 id=None,
                 target=None,
                ):
        self.ref = ref
        self.location = location
        self.tooltip = tooltip
        self.display = display
        self.id = id
        self.target = target


class HyperlinkList(Serialisable):

    tagname = "hyperlinks"

    __expected_type = Hyperlink
    hyperlink = Sequence(expected_type=__expected_type)

    def __init__(self, hyperlink=()):
        self.hyperlink = hyperlink

""",
    },
    'openpyxl.worksheet.merge': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import copy

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Integer,
    Sequence,
)

from openpyxl.cell.cell import MergedCell
from openpyxl.styles.borders import Border

from .cell_range import CellRange


class MergeCell(CellRange):

    tagname = "mergeCell"
    ref = CellRange.coord

    __attrs__ = ("ref",)


    def __init__(self,
                 ref=None,
                ):
        super().__init__(ref)


    def __copy__(self):
        return self.__class__(self.ref)


class MergeCells(Serialisable):

    tagname = "mergeCells"

    count = Integer(allow_none=True)
    mergeCell = Sequence(expected_type=MergeCell, )

    __elements__ = ('mergeCell',)
    __attrs__ = ('count',)

    def __init__(self,
                 count=None,
                 mergeCell=(),
                ):
        self.mergeCell = mergeCell


    @property
    def count(self):
        return len(self.mergeCell)


class MergedCellRange(CellRange):

    """
    MergedCellRange stores the border information of a merged cell in the top
    left cell of the merged cell.
    The remaining cells in the merged cell are stored as MergedCell objects and
    get their border information from the upper left cell.
    """

    def __init__(self, worksheet, coord):
        self.ws = worksheet
        super().__init__(range_string=coord)
        self.start_cell = None
        self._get_borders()


    def _get_borders(self):
        """
        If the upper left cell of the merged cell does not yet exist, it is
        created.
        The upper left cell gets the border information of the bottom and right
        border from the bottom right cell of the merged cell, if available.
        """

        # Top-left cell.
        self.start_cell = self.ws._cells.get((self.min_row, self.min_col))
        if self.start_cell is None:
            self.start_cell = self.ws.cell(row=self.min_row, column=self.min_col)

        # Bottom-right cell
        end_cell = self.ws._cells.get((self.max_row, self.max_col))
        if end_cell is not None:
            self.start_cell.border += Border(right=end_cell.border.right,
                                             bottom=end_cell.border.bottom)


    def format(self):
        """
        Each cell of the merged cell is created as MergedCell if it does not
        already exist.

        The MergedCells at the edge of the merged cell gets its borders from
        the upper left cell.

         - The top MergedCells get the top border from the top left cell.
         - The bottom MergedCells get the bottom border from the top left cell.
         - The left MergedCells get the left border from the top left cell.
         - The right MergedCells get the right border from the top left cell.
        """

        names = ['top', 'left', 'right', 'bottom']

        for name in names:
            side = getattr(self.start_cell.border, name)
            if side and side.style is None:
                continue # don't need to do anything if there is no border style
            border = Border(**{name:side})
            for coord in getattr(self, name):
                cell = self.ws._cells.get(coord)
                if cell is None:
                    row, col = coord
                    cell = MergedCell(self.ws, row=row, column=col)
                    self.ws._cells[(cell.row, cell.column)] = cell
                cell.border += border

        protected = self.start_cell.protection is not None
        if protected:
            protection = copy.copy(self.start_cell.protection)
        for coord in self.cells:
            cell = self.ws._cells.get(coord)
            if cell is None:
                row, col = coord
                cell = MergedCell(self.ws, row=row, column=col)
                self.ws._cells[(cell.row, cell.column)] = cell

            if protected:
                cell.protection = protection


    def __contains__(self, coord):
        return coord in CellRange(self.coord)


    def __copy__(self):
        return self.__class__(self.ws, self.coord)

''',
    },
    'openpyxl.worksheet.ole': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
    String,
    Set,
    Bool,
    Sequence,
)

from openpyxl.drawing.spreadsheet_drawing import AnchorMarker
from openpyxl.xml.constants import SHEET_DRAWING_NS


class ObjectAnchor(Serialisable):

    tagname = "anchor"

    _from = Typed(expected_type=AnchorMarker, namespace=SHEET_DRAWING_NS)
    to = Typed(expected_type=AnchorMarker, namespace=SHEET_DRAWING_NS)
    moveWithCells = Bool(allow_none=True)
    sizeWithCells = Bool(allow_none=True)
    z_order = Integer(allow_none=True, hyphenated=True)


    def __init__(self,
                 _from=None,
                 to=None,
                 moveWithCells=False,
                 sizeWithCells=False,
                 z_order=None,
                ):
        self._from = _from
        self.to = to
        self.moveWithCells = moveWithCells
        self.sizeWithCells = sizeWithCells
        self.z_order = z_order


class ObjectPr(Serialisable):

    tagname = "objectPr"

    anchor = Typed(expected_type=ObjectAnchor, )
    locked = Bool(allow_none=True)
    defaultSize = Bool(allow_none=True)
    _print = Bool(allow_none=True)
    disabled = Bool(allow_none=True)
    uiObject = Bool(allow_none=True)
    autoFill = Bool(allow_none=True)
    autoLine = Bool(allow_none=True)
    autoPict = Bool(allow_none=True)
    macro = String()
    altText = String(allow_none=True)
    dde = Bool(allow_none=True)

    __elements__ = ('anchor',)

    def __init__(self,
                 anchor=None,
                 locked=True,
                 defaultSize=True,
                 _print=True,
                 disabled=False,
                 uiObject=False,
                 autoFill=True,
                 autoLine=True,
                 autoPict=True,
                 macro=None,
                 altText=None,
                 dde=False,
                ):
        self.anchor = anchor
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.autoPict = autoPict
        self.macro = macro
        self.altText = altText
        self.dde = dde


class OleObject(Serialisable):

    tagname = "oleObject"

    objectPr = Typed(expected_type=ObjectPr, allow_none=True)
    progId = String(allow_none=True)
    dvAspect = Set(values=(['DVASPECT_CONTENT', 'DVASPECT_ICON']))
    link = String(allow_none=True)
    oleUpdate = Set(values=(['OLEUPDATE_ALWAYS', 'OLEUPDATE_ONCALL']))
    autoLoad = Bool(allow_none=True)
    shapeId = Integer()

    __elements__ = ('objectPr',)

    def __init__(self,
                 objectPr=None,
                 progId=None,
                 dvAspect='DVASPECT_CONTENT',
                 link=None,
                 oleUpdate=None,
                 autoLoad=False,
                 shapeId=None,
                ):
        self.objectPr = objectPr
        self.progId = progId
        self.dvAspect = dvAspect
        self.link = link
        self.oleUpdate = oleUpdate
        self.autoLoad = autoLoad
        self.shapeId = shapeId


class OleObjects(Serialisable):

    tagname = "oleObjects"

    oleObject = Sequence(expected_type=OleObject)

    __elements__ = ('oleObject',)

    def __init__(self,
                 oleObject=(),
                ):
        self.oleObject = oleObject


""",
    },
    'openpyxl.worksheet.page': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Float,
    Bool,
    Integer,
    NoneSet,
    )
from openpyxl.descriptors.excel import UniversalMeasure, Relation


class PrintPageSetup(Serialisable):
    """ Worksheet print page setup """

    tagname = "pageSetup"

    orientation = NoneSet(values=("default", "portrait", "landscape"))
    paperSize = Integer(allow_none=True)
    scale = Integer(allow_none=True)
    fitToHeight = Integer(allow_none=True)
    fitToWidth = Integer(allow_none=True)
    firstPageNumber = Integer(allow_none=True)
    useFirstPageNumber = Bool(allow_none=True)
    paperHeight = UniversalMeasure(allow_none=True)
    paperWidth = UniversalMeasure(allow_none=True)
    pageOrder = NoneSet(values=("downThenOver", "overThenDown"))
    usePrinterDefaults = Bool(allow_none=True)
    blackAndWhite = Bool(allow_none=True)
    draft = Bool(allow_none=True)
    cellComments = NoneSet(values=("asDisplayed", "atEnd"))
    errors = NoneSet(values=("displayed", "blank", "dash", "NA"))
    horizontalDpi = Integer(allow_none=True)
    verticalDpi = Integer(allow_none=True)
    copies = Integer(allow_none=True)
    id = Relation()


    def __init__(self,
                 worksheet=None,
                 orientation=None,
                 paperSize=None,
                 scale=None,
                 fitToHeight=None,
                 fitToWidth=None,
                 firstPageNumber=None,
                 useFirstPageNumber=None,
                 paperHeight=None,
                 paperWidth=None,
                 pageOrder=None,
                 usePrinterDefaults=None,
                 blackAndWhite=None,
                 draft=None,
                 cellComments=None,
                 errors=None,
                 horizontalDpi=None,
                 verticalDpi=None,
                 copies=None,
                 id=None):
        self._parent = worksheet
        self.orientation = orientation
        self.paperSize = paperSize
        self.scale = scale
        self.fitToHeight = fitToHeight
        self.fitToWidth = fitToWidth
        self.firstPageNumber = firstPageNumber
        self.useFirstPageNumber = useFirstPageNumber
        self.paperHeight = paperHeight
        self.paperWidth = paperWidth
        self.pageOrder = pageOrder
        self.usePrinterDefaults = usePrinterDefaults
        self.blackAndWhite = blackAndWhite
        self.draft = draft
        self.cellComments = cellComments
        self.errors = errors
        self.horizontalDpi = horizontalDpi
        self.verticalDpi = verticalDpi
        self.copies = copies
        self.id = id


    def __bool__(self):
        return bool(dict(self))




    @property
    def sheet_properties(self):
        """
        Proxy property
        """
        return self._parent.sheet_properties.pageSetUpPr


    @property
    def fitToPage(self):
        return self.sheet_properties.fitToPage


    @fitToPage.setter
    def fitToPage(self, value):
        self.sheet_properties.fitToPage = value


    @property
    def autoPageBreaks(self):
        return self.sheet_properties.autoPageBreaks


    @autoPageBreaks.setter
    def autoPageBreaks(self, value):
        self.sheet_properties.autoPageBreaks = value


    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        self.id = None # strip link to binary settings
        return self


class PrintOptions(Serialisable):
    """ Worksheet print options """

    tagname = "printOptions"
    horizontalCentered = Bool(allow_none=True)
    verticalCentered = Bool(allow_none=True)
    headings = Bool(allow_none=True)
    gridLines = Bool(allow_none=True)
    gridLinesSet = Bool(allow_none=True)

    def __init__(self, horizontalCentered=None,
                 verticalCentered=None,
                 headings=None,
                 gridLines=None,
                 gridLinesSet=None,
                 ):
        self.horizontalCentered = horizontalCentered
        self.verticalCentered = verticalCentered
        self.headings = headings
        self.gridLines = gridLines
        self.gridLinesSet = gridLinesSet


    def __bool__(self):
        return bool(dict(self))


class PageMargins(Serialisable):
    """
    Information about page margins for view/print layouts.
    Standard values (in inches)
    left, right = 0.75
    top, bottom = 1
    header, footer = 0.5
    """
    tagname = "pageMargins"

    left = Float()
    right = Float()
    top = Float()
    bottom = Float()
    header = Float()
    footer = Float()

    def __init__(self, left=0.75, right=0.75, top=1, bottom=1, header=0.5,
                 footer=0.5):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.header = header
        self.footer = footer

''',
    },
    'openpyxl.worksheet.pagebreak': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Integer,
    Bool,
    Sequence,
)


class Break(Serialisable):

    tagname = "brk"

    id = Integer(allow_none=True)
    min = Integer(allow_none=True)
    max = Integer(allow_none=True)
    man = Bool(allow_none=True)
    pt = Bool(allow_none=True)

    def __init__(self,
                 id=0,
                 min=0,
                 max=16383,
                 man=True,
                 pt=None,
                ):
        self.id = id
        self.min = min
        self.max = max
        self.man = man
        self.pt = pt


class RowBreak(Serialisable):

    tagname = "rowBreaks"

    count = Integer(allow_none=True)
    manualBreakCount = Integer(allow_none=True)
    brk = Sequence(expected_type=Break, allow_none=True)

    __elements__ = ('brk',)
    __attrs__ = ("count", "manualBreakCount",)

    def __init__(self,
                 count=None,
                 manualBreakCount=None,
                 brk=(),
                ):
        self.brk = brk


    def __bool__(self):
        return len(self.brk) > 0


    def __len__(self):
        return len(self.brk)


    @property
    def count(self):
        return len(self)


    @property
    def manualBreakCount(self):
        return len(self)


    def append(self, brk=None):
        """
        Add a page break
        """
        vals = list(self.brk)
        if not isinstance(brk, Break):
            brk = Break(id=self.count+1)
        vals.append(brk)
        self.brk = vals


PageBreak = RowBreak


class ColBreak(RowBreak):

    tagname = "colBreaks"

    count = RowBreak.count
    manualBreakCount = RowBreak.manualBreakCount
    brk = RowBreak.brk

    __attrs__ = RowBreak.__attrs__

''',
    },
    'openpyxl.worksheet.picture': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable

# same as related

class SheetBackgroundPicture(Serialisable):

    tagname = "sheetBackgroundPicture"

""",
    },
    'openpyxl.worksheet.print_settings': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

import re
from openpyxl.descriptors import (
    Strict,
    Integer,
    String,
    Typed,
)
from openpyxl.utils import quote_sheetname, absolute_coordinate
from openpyxl.utils.cell import SHEET_TITLE, SHEETRANGE_RE, RANGE_EXPR

from .cell_range import MultiCellRange

COL_RANGE = r"""(?P<cols>[$]?(?P<min_col>[a-zA-Z]{1,3}):[$]?(?P<max_col>[a-zA-Z]{1,3}))"""
COL_RANGE_RE = re.compile(COL_RANGE)
ROW_RANGE = r"""(?P<rows>[$]?(?P<min_row>\d+):[$]?(?P<max_row>\d+))"""
ROW_RANGE_RE = re.compile(ROW_RANGE)
TITLES_REGEX = re.compile("""{0}{1}?,?{2}?,?""".format(SHEET_TITLE, ROW_RANGE, COL_RANGE),
                          re.VERBOSE)
PRINT_AREA_RE = re.compile(f"({SHEET_TITLE})?(?P<cells>{RANGE_EXPR})", re.VERBOSE)

class ColRange(Strict):
    """
    Represent a range of at least one column
    """

    min_col = String()
    max_col = String()


    def __init__(self, range_string=None, min_col=None, max_col=None):
        if range_string is not None:
            match = COL_RANGE_RE.match(range_string)
            if not match:
                raise ValueError(f"{range_string} is not a valid column range")
            min_col, max_col = match.groups()[1:]
        self.min_col = min_col
        self.max_col = max_col


    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return (self.min_col == other.min_col
                    and
                    self.max_col == other.max_col)
        elif isinstance(other, str):
            return (str(self) == other
                    or
                    f"{self.min_col}:{self.max_col}")
        return False


    def __repr__(self):
        return f"Range of columns from '{self.min_col}' to '{self.max_col}'"


    def __str__(self):
        return f"${self.min_col}:${self.max_col}"


class RowRange(Strict):
    """
    Represent a range of at least one row
    """

    min_row = Integer()
    max_row = Integer()

    def __init__(self, range_string=None, min_row=None, max_row=None):
        if range_string is not None:
            match = ROW_RANGE_RE.match(range_string)
            if not match:
                raise ValueError(f"{range_string} is not a valid row range")
            min_row, max_row = match.groups()[1:]
        self.min_row = min_row
        self.max_row = max_row


    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return (self.min_row == other.min_row
                    and
                    self.max_row == other.max_row)
        elif isinstance(other, str):
            return (str(self) == other
                    or
                    f"{self.min_row}:{self.max_row}")
        return False

    def __repr__(self):
        return f"Range of rows from '{self.min_row}' to '{self.max_row}'"


    def __str__(self):
        return f"${self.min_row}:${self.max_row}"


class PrintTitles(Strict):
    """
    Contains at least either a range of rows or columns
    """

    cols = Typed(expected_type=ColRange, allow_none=True)
    rows = Typed(expected_type=RowRange, allow_none=True)
    title = String()


    def __init__(self, cols=None, rows=None, title=""):
        self.cols = cols
        self.rows = rows
        self.title = title


    @classmethod
    def from_string(cls, value):
        kw = dict((k, v) for match in TITLES_REGEX.finditer(value)
                  for k, v in match.groupdict().items() if v)

        if not kw:
            raise ValueError(f"{value} is not a valid print titles definition")

        cols = rows = None

        if "cols" in kw:
            cols = ColRange(kw["cols"])
        if "rows" in kw:
            rows = RowRange(kw["rows"])

        title = kw.get("quoted") or kw.get("notquoted")

        return cls(cols=cols, rows=rows, title=title)


    def __eq__(self, other):
        if isinstance(other, self.__class__):
            return (self.cols == other.cols
                    and
                    self.rows == other.rows
                    and
                    self.title == other.title)
        elif isinstance(other, str):
            return str(self) == other
        return False

    def __repr__(self):
        return f"Print titles for sheet {self.title} cols {self.rows}, rows {self.cols}"


    def __str__(self):
        title = quote_sheetname(self.title)
        titles = ",".join([f"{title}!{value}" for value in (self.rows, self.cols) if value])
        return titles or ""


class PrintArea(MultiCellRange):


    @classmethod
    def from_string(cls, value):
        new = []
        for m in PRINT_AREA_RE.finditer(value): # can be multiple
            coord = m.group("cells")
            if coord:
                new.append(coord)
        return cls(new)


    def __init__(self, ranges=(), title=""):
        self.title = ""
        super().__init__(ranges)


    def __str__(self):
        if self.ranges:
            return ",".join([f"{quote_sheetname(self.title)}!{absolute_coordinate(str(range))}"
                             for range in self.sorted()])
        return ""


    def __eq__(self, other):
        super().__eq__(other)
        if isinstance(other, str):
            return str(self) == other

''',
    },
    'openpyxl.worksheet.properties': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Worksheet Properties"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import String, Bool, Typed
from openpyxl.styles.colors import ColorDescriptor


class Outline(Serialisable):

    tagname = "outlinePr"

    applyStyles = Bool(allow_none=True)
    summaryBelow = Bool(allow_none=True)
    summaryRight = Bool(allow_none=True)
    showOutlineSymbols = Bool(allow_none=True)


    def __init__(self,
                 applyStyles=None,
                 summaryBelow=None,
                 summaryRight=None,
                 showOutlineSymbols=None
                 ):
        self.applyStyles = applyStyles
        self.summaryBelow = summaryBelow
        self.summaryRight = summaryRight
        self.showOutlineSymbols = showOutlineSymbols


class PageSetupProperties(Serialisable):

    tagname = "pageSetUpPr"

    autoPageBreaks = Bool(allow_none=True)
    fitToPage = Bool(allow_none=True)

    def __init__(self, autoPageBreaks=None, fitToPage=None):
        self.autoPageBreaks = autoPageBreaks
        self.fitToPage = fitToPage


class WorksheetProperties(Serialisable):

    tagname = "sheetPr"

    codeName = String(allow_none=True)
    enableFormatConditionsCalculation = Bool(allow_none=True)
    filterMode = Bool(allow_none=True)
    published = Bool(allow_none=True)
    syncHorizontal = Bool(allow_none=True)
    syncRef = String(allow_none=True)
    syncVertical = Bool(allow_none=True)
    transitionEvaluation = Bool(allow_none=True)
    transitionEntry = Bool(allow_none=True)
    tabColor = ColorDescriptor(allow_none=True)
    outlinePr = Typed(expected_type=Outline, allow_none=True)
    pageSetUpPr = Typed(expected_type=PageSetupProperties, allow_none=True)

    __elements__ = ('tabColor', 'outlinePr', 'pageSetUpPr')


    def __init__(self,
                 codeName=None,
                 enableFormatConditionsCalculation=None,
                 filterMode=None,
                 published=None,
                 syncHorizontal=None,
                 syncRef=None,
                 syncVertical=None,
                 transitionEvaluation=None,
                 transitionEntry=None,
                 tabColor=None,
                 outlinePr=None,
                 pageSetUpPr=None
                 ):
        """ Attributes """
        self.codeName = codeName
        self.enableFormatConditionsCalculation = enableFormatConditionsCalculation
        self.filterMode = filterMode
        self.published = published
        self.syncHorizontal = syncHorizontal
        self.syncRef = syncRef
        self.syncVertical = syncVertical
        self.transitionEvaluation = transitionEvaluation
        self.transitionEntry = transitionEntry
        """ Elements """
        self.tabColor = tabColor
        if outlinePr is None:
            self.outlinePr = Outline(summaryBelow=True, summaryRight=True)
        else:
            self.outlinePr = outlinePr

        if pageSetUpPr is None:
            pageSetUpPr = PageSetupProperties()
        self.pageSetUpPr = pageSetUpPr

''',
    },
    'openpyxl.worksheet.protection': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Bool,
    String,
    Alias,
    Integer,
)
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.excel import (
    Base64Binary,
)
from openpyxl.utils.protection import hash_password


class _Protected:
    _password = None

    def set_password(self, value='', already_hashed=False):
        """Set a password on this sheet."""
        if not already_hashed:
            value = hash_password(value)
        self._password = value

    @property
    def password(self):
        """Return the password value, regardless of hash."""
        return self._password

    @password.setter
    def password(self, value):
        """Set a password directly, forcing a hash step."""
        self.set_password(value)


class SheetProtection(Serialisable, _Protected):
    """
    Information about protection of various aspects of a sheet. True values
    mean that protection for the object or action is active This is the
    **default** when protection is active, ie. users cannot do something
    """

    tagname = "sheetProtection"

    sheet = Bool()
    enabled = Alias('sheet')
    objects = Bool()
    scenarios = Bool()
    formatCells = Bool()
    formatColumns = Bool()
    formatRows = Bool()
    insertColumns = Bool()
    insertRows = Bool()
    insertHyperlinks = Bool()
    deleteColumns = Bool()
    deleteRows = Bool()
    selectLockedCells = Bool()
    selectUnlockedCells = Bool()
    sort = Bool()
    autoFilter = Bool()
    pivotTables = Bool()
    saltValue = Base64Binary(allow_none=True)
    spinCount = Integer(allow_none=True)
    algorithmName = String(allow_none=True)
    hashValue = Base64Binary(allow_none=True)


    __attrs__ = ('selectLockedCells', 'selectUnlockedCells', 'algorithmName',
              'sheet', 'objects', 'insertRows', 'insertHyperlinks', 'autoFilter',
              'scenarios', 'formatColumns', 'deleteColumns', 'insertColumns',
              'pivotTables', 'deleteRows', 'formatCells', 'saltValue', 'formatRows',
              'sort', 'spinCount', 'password', 'hashValue')


    def __init__(self, sheet=False, objects=False, scenarios=False,
                 formatCells=True, formatRows=True, formatColumns=True,
                 insertColumns=True, insertRows=True, insertHyperlinks=True,
                 deleteColumns=True, deleteRows=True, selectLockedCells=False,
                 selectUnlockedCells=False, sort=True, autoFilter=True, pivotTables=True,
                 password=None, algorithmName=None, saltValue=None, spinCount=None, hashValue=None):
        self.sheet = sheet
        self.objects = objects
        self.scenarios = scenarios
        self.formatCells = formatCells
        self.formatColumns = formatColumns
        self.formatRows = formatRows
        self.insertColumns = insertColumns
        self.insertRows = insertRows
        self.insertHyperlinks = insertHyperlinks
        self.deleteColumns = deleteColumns
        self.deleteRows = deleteRows
        self.selectLockedCells = selectLockedCells
        self.selectUnlockedCells = selectUnlockedCells
        self.sort = sort
        self.autoFilter = autoFilter
        self.pivotTables = pivotTables
        if password is not None:
            self.password = password
        self.algorithmName = algorithmName
        self.saltValue = saltValue
        self.spinCount = spinCount
        self.hashValue = hashValue


    def set_password(self, value='', already_hashed=False):
        super().set_password(value, already_hashed)
        self.enable()


    def enable(self):
        self.sheet = True


    def disable(self):
        self.sheet = False


    def  __bool__(self):
        return self.sheet


''',
    },
    'openpyxl.worksheet.related': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors.excel import Relation


class Related(Serialisable):

    id = Relation()


    def __init__(self, id=None):
        self.id = id


    def to_tree(self, tagname, idx=None):
        return super().to_tree(tagname)

""",
    },
    'openpyxl.worksheet.scenario': {
        'is_package': False,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    String,
    Integer,
    Bool,
    Sequence,
    Convertible,
)
from .cell_range import MultiCellRange


class InputCells(Serialisable):

    tagname = "inputCells"

    r = String()
    deleted = Bool(allow_none=True)
    undone = Bool(allow_none=True)
    val = String()
    numFmtId = Integer(allow_none=True)

    def __init__(self,
                 r=None,
                 deleted=False,
                 undone=False,
                 val=None,
                 numFmtId=None,
                ):
        self.r = r
        self.deleted = deleted
        self.undone = undone
        self.val = val
        self.numFmtId = numFmtId


class Scenario(Serialisable):

    tagname = "scenario"

    inputCells = Sequence(expected_type=InputCells)
    name = String()
    locked = Bool(allow_none=True)
    hidden = Bool(allow_none=True)
    user = String(allow_none=True)
    comment = String(allow_none=True)

    __elements__ = ('inputCells',)
    __attrs__ = ('name', 'locked', 'hidden', 'user', 'comment', 'count')

    def __init__(self,
                 inputCells=(),
                 name=None,
                 locked=False,
                 hidden=False,
                 count=None,
                 user=None,
                 comment=None,
                ):
        self.inputCells = inputCells
        self.name = name
        self.locked = locked
        self.hidden = hidden
        self.user = user
        self.comment = comment


    @property
    def count(self):
        return len(self.inputCells)


class ScenarioList(Serialisable):

    tagname = "scenarios"

    scenario = Sequence(expected_type=Scenario)
    current = Integer(allow_none=True)
    show = Integer(allow_none=True)
    sqref = Convertible(expected_type=MultiCellRange, allow_none=True)

    __elements__ = ('scenario',)

    def __init__(self,
                 scenario=(),
                 current=None,
                 show=None,
                 sqref=None,
                ):
        self.scenario = scenario
        self.current = current
        self.show = show
        self.sqref = sqref


    def append(self, scenario):
        s = self.scenario
        s.append(scenario)
        self.scenario = s


    def __bool__(self):
        return bool(self.scenario)


""",
    },
    'openpyxl.worksheet.smart_tag': {
        'is_package': False,
        'source': r"""
#Autogenerated schema
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Integer,
    String,
    Sequence,
)


class CellSmartTagPr(Serialisable):

    tagname = "cellSmartTagPr"

    key = String()
    val = String()

    def __init__(self,
                 key=None,
                 val=None,
                ):
        self.key = key
        self.val = val


class CellSmartTag(Serialisable):

    tagname = "cellSmartTag"

    cellSmartTagPr = Sequence(expected_type=CellSmartTagPr)
    type = Integer()
    deleted = Bool(allow_none=True)
    xmlBased = Bool(allow_none=True)

    __elements__ = ('cellSmartTagPr',)

    def __init__(self,
                 cellSmartTagPr=(),
                 type=None,
                 deleted=False,
                 xmlBased=False,
                ):
        self.cellSmartTagPr = cellSmartTagPr
        self.type = type
        self.deleted = deleted
        self.xmlBased = xmlBased


class CellSmartTags(Serialisable):

    tagname = "cellSmartTags"

    cellSmartTag = Sequence(expected_type=CellSmartTag)
    r = String()

    __elements__ = ('cellSmartTag',)

    def __init__(self,
                 cellSmartTag=(),
                 r=None,
                ):
        self.cellSmartTag = cellSmartTag
        self.r = r


class SmartTags(Serialisable):

    tagname = "smartTags"

    cellSmartTags = Sequence(expected_type=CellSmartTags)

    __elements__ = ('cellSmartTags',)

    def __init__(self,
                 cellSmartTags=(),
                ):
        self.cellSmartTags = cellSmartTags


""",
    },
    'openpyxl.worksheet.table': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Descriptor,
    Alias,
    Typed,
    Bool,
    Integer,
    NoneSet,
    String,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList, CellRange
from openpyxl.descriptors.sequence import NestedSequence
from openpyxl.xml.constants import SHEET_MAIN_NS, REL_NS
from openpyxl.xml.functions import tostring
from openpyxl.utils import range_boundaries
from openpyxl.utils.escape import escape, unescape

from .related import Related

from .filters import (
    AutoFilter,
    SortState,
)

TABLESTYLES = tuple(
    ["TableStyleMedium{0}".format(i) for i in range(1, 29)]
    + ["TableStyleLight{0}".format(i) for i in range(1, 22)]
    + ["TableStyleDark{0}".format(i) for i in range(1, 12)]
)

PIVOTSTYLES = tuple(
    ["PivotStyleMedium{0}".format(i) for i in range(1, 29)]
    + ["PivotStyleLight{0}".format(i) for i in range(1, 29)]
    + ["PivotStyleDark{0}".format(i) for i in range(1, 29)]
)


class TableStyleInfo(Serialisable):

    tagname = "tableStyleInfo"

    name = String(allow_none=True)
    showFirstColumn = Bool(allow_none=True)
    showLastColumn = Bool(allow_none=True)
    showRowStripes = Bool(allow_none=True)
    showColumnStripes = Bool(allow_none=True)

    def __init__(self,
                 name=None,
                 showFirstColumn=None,
                 showLastColumn=None,
                 showRowStripes=None,
                 showColumnStripes=None,
                ):
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes


class XMLColumnProps(Serialisable):

    tagname = "xmlColumnPr"

    mapId = Integer()
    xpath = String()
    denormalized = Bool(allow_none=True)
    xmlDataType = String()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 mapId=None,
                 xpath=None,
                 denormalized=None,
                 xmlDataType=None,
                 extLst=None,
                ):
        self.mapId = mapId
        self.xpath = xpath
        self.denormalized = denormalized
        self.xmlDataType = xmlDataType


class TableFormula(Serialisable):

    tagname = "tableFormula"

    ## Note formula is stored as the text value

    array = Bool(allow_none=True)
    attr_text = Descriptor()
    text = Alias('attr_text')


    def __init__(self,
                 array=None,
                 attr_text=None,
                ):
        self.array = array
        self.attr_text = attr_text


class TableColumn(Serialisable):

    tagname = "tableColumn"

    id = Integer()
    uniqueName = String(allow_none=True)
    name = String()
    totalsRowFunction = NoneSet(values=(['sum', 'min', 'max', 'average',
                                         'count', 'countNums', 'stdDev', 'var', 'custom']))
    totalsRowLabel = String(allow_none=True)
    queryTableFieldId = Integer(allow_none=True)
    headerRowDxfId = Integer(allow_none=True)
    dataDxfId = Integer(allow_none=True)
    totalsRowDxfId = Integer(allow_none=True)
    headerRowCellStyle = String(allow_none=True)
    dataCellStyle = String(allow_none=True)
    totalsRowCellStyle = String(allow_none=True)
    calculatedColumnFormula = Typed(expected_type=TableFormula, allow_none=True)
    totalsRowFormula = Typed(expected_type=TableFormula, allow_none=True)
    xmlColumnPr = Typed(expected_type=XMLColumnProps, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('calculatedColumnFormula', 'totalsRowFormula',
                    'xmlColumnPr', 'extLst')

    def __init__(self,
                 id=None,
                 uniqueName=None,
                 name=None,
                 totalsRowFunction=None,
                 totalsRowLabel=None,
                 queryTableFieldId=None,
                 headerRowDxfId=None,
                 dataDxfId=None,
                 totalsRowDxfId=None,
                 headerRowCellStyle=None,
                 dataCellStyle=None,
                 totalsRowCellStyle=None,
                 calculatedColumnFormula=None,
                 totalsRowFormula=None,
                 xmlColumnPr=None,
                 extLst=None,
                ):
        self.id = id
        self.uniqueName = uniqueName
        self.name = name
        self.totalsRowFunction = totalsRowFunction
        self.totalsRowLabel = totalsRowLabel
        self.queryTableFieldId = queryTableFieldId
        self.headerRowDxfId = headerRowDxfId
        self.dataDxfId = dataDxfId
        self.totalsRowDxfId = totalsRowDxfId
        self.headerRowCellStyle = headerRowCellStyle
        self.dataCellStyle = dataCellStyle
        self.totalsRowCellStyle = totalsRowCellStyle
        self.calculatedColumnFormula = calculatedColumnFormula
        self.totalsRowFormula = totalsRowFormula
        self.xmlColumnPr = xmlColumnPr
        self.extLst = extLst


    def __iter__(self):
        for k, v in super().__iter__():
            if k == 'name':
                v = escape(v)
            yield k, v


    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        self.name = unescape(self.name)
        return self


class TableNameDescriptor(String):

    """
    Table names cannot have spaces in them
    """

    def __set__(self, instance, value):
        if value is not None and " " in value:
            raise ValueError("Table names cannot have spaces")
        super().__set__(instance, value)


class Table(Serialisable):

    _path = "/tables/table{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    _rel_type = REL_NS + "/table"
    _rel_id = None

    tagname = "table"

    id = Integer()
    name = String(allow_none=True)
    displayName = TableNameDescriptor()
    comment = String(allow_none=True)
    ref = CellRange()
    tableType = NoneSet(values=(['worksheet', 'xml', 'queryTable']))
    headerRowCount = Integer(allow_none=True)
    insertRow = Bool(allow_none=True)
    insertRowShift = Bool(allow_none=True)
    totalsRowCount = Integer(allow_none=True)
    totalsRowShown = Bool(allow_none=True)
    published = Bool(allow_none=True)
    headerRowDxfId = Integer(allow_none=True)
    dataDxfId = Integer(allow_none=True)
    totalsRowDxfId = Integer(allow_none=True)
    headerRowBorderDxfId = Integer(allow_none=True)
    tableBorderDxfId = Integer(allow_none=True)
    totalsRowBorderDxfId = Integer(allow_none=True)
    headerRowCellStyle = String(allow_none=True)
    dataCellStyle = String(allow_none=True)
    totalsRowCellStyle = String(allow_none=True)
    connectionId = Integer(allow_none=True)
    autoFilter = Typed(expected_type=AutoFilter, allow_none=True)
    sortState = Typed(expected_type=SortState, allow_none=True)
    tableColumns = NestedSequence(expected_type=TableColumn, count=True)
    tableStyleInfo = Typed(expected_type=TableStyleInfo, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('autoFilter', 'sortState', 'tableColumns',
                    'tableStyleInfo')

    def __init__(self,
                 id=1,
                 displayName=None,
                 ref=None,
                 name=None,
                 comment=None,
                 tableType=None,
                 headerRowCount=1,
                 insertRow=None,
                 insertRowShift=None,
                 totalsRowCount=None,
                 totalsRowShown=None,
                 published=None,
                 headerRowDxfId=None,
                 dataDxfId=None,
                 totalsRowDxfId=None,
                 headerRowBorderDxfId=None,
                 tableBorderDxfId=None,
                 totalsRowBorderDxfId=None,
                 headerRowCellStyle=None,
                 dataCellStyle=None,
                 totalsRowCellStyle=None,
                 connectionId=None,
                 autoFilter=None,
                 sortState=None,
                 tableColumns=(),
                 tableStyleInfo=None,
                 extLst=None,
                ):
        self.id = id
        self.displayName = displayName
        if name is None:
            name = displayName
        self.name = name
        self.comment = comment
        self.ref = ref
        self.tableType = tableType
        self.headerRowCount = headerRowCount
        self.insertRow = insertRow
        self.insertRowShift = insertRowShift
        self.totalsRowCount = totalsRowCount
        self.totalsRowShown = totalsRowShown
        self.published = published
        self.headerRowDxfId = headerRowDxfId
        self.dataDxfId = dataDxfId
        self.totalsRowDxfId = totalsRowDxfId
        self.headerRowBorderDxfId = headerRowBorderDxfId
        self.tableBorderDxfId = tableBorderDxfId
        self.totalsRowBorderDxfId = totalsRowBorderDxfId
        self.headerRowCellStyle = headerRowCellStyle
        self.dataCellStyle = dataCellStyle
        self.totalsRowCellStyle = totalsRowCellStyle
        self.connectionId = connectionId
        self.autoFilter = autoFilter
        self.sortState = sortState
        self.tableColumns = tableColumns
        self.tableStyleInfo = tableStyleInfo


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        """
        Return path within the archive
        """
        return "/xl" + self._path.format(self.id)


    def _write(self, archive):
        """
        Serialise to XML and write to archive
        """
        xml = self.to_tree()
        archive.writestr(self.path[1:], tostring(xml))


    def _initialise_columns(self):
        """
        Create a list of table columns from a cell range
        Always set a ref if we have headers (the default)
        Column headings must be strings and must match cells in the worksheet.
        """

        min_col, min_row, max_col, max_row = range_boundaries(self.ref)
        for idx in range(min_col, max_col+1):
            col = TableColumn(id=idx, name="Column{0}".format(idx))
            self.tableColumns.append(col)
        if self.headerRowCount and not self.autoFilter:
            self.autoFilter = AutoFilter(ref=self.ref)


    @property
    def column_names(self):
        return [column.name for column in self.tableColumns]


class TablePartList(Serialisable):

    tagname = "tableParts"

    count = Integer(allow_none=True)
    tablePart = Sequence(expected_type=Related)

    __elements__ = ('tablePart',)
    __attrs__ = ('count',)

    def __init__(self,
                 count=None,
                 tablePart=(),
                ):
        self.tablePart = tablePart


    def append(self, part):
        self.tablePart.append(part)


    @property
    def count(self):
        return len(self.tablePart)


    def __bool__(self):
        return bool(self.tablePart)


class TableList(dict):


    def add(self, table):
        if not isinstance(table, Table):
            raise TypeError("You can only add tables")
        self[table.name] = table


    def get(self, name=None, table_range=None):
        if name is not None:
            return super().get(name)
        for table in self.values():
            if table_range == table.ref:
                return table


    def items(self):
        return [(name, table.ref) for name, table in super().items()]

''',
    },
    'openpyxl.worksheet.views': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

from openpyxl.descriptors import (
    Bool,
    Integer,
    String,
    Set,
    Float,
    Typed,
    NoneSet,
    Sequence,
)
from openpyxl.descriptors.excel import ExtensionList
from openpyxl.descriptors.serialisable import Serialisable


class Pane(Serialisable):
    xSplit = Float(allow_none=True)
    ySplit = Float(allow_none=True)
    topLeftCell = String(allow_none=True)
    activePane = Set(values=("bottomRight", "topRight", "bottomLeft", "topLeft"))
    state = Set(values=("split", "frozen", "frozenSplit"))

    def __init__(self,
                 xSplit=None,
                 ySplit=None,
                 topLeftCell=None,
                 activePane="topLeft",
                 state="split"):
        self.xSplit = xSplit
        self.ySplit = ySplit
        self.topLeftCell = topLeftCell
        self.activePane = activePane
        self.state = state


class Selection(Serialisable):
    pane = NoneSet(values=("bottomRight", "topRight", "bottomLeft", "topLeft"))
    activeCell = String(allow_none=True)
    activeCellId = Integer(allow_none=True)
    sqref = String(allow_none=True)

    def __init__(self,
                 pane=None,
                 activeCell="A1",
                 activeCellId=None,
                 sqref="A1"):
        self.pane = pane
        self.activeCell = activeCell
        self.activeCellId = activeCellId
        self.sqref = sqref


class SheetView(Serialisable):

    """Information about the visible portions of this sheet."""

    tagname = "sheetView"

    windowProtection = Bool(allow_none=True)
    showFormulas = Bool(allow_none=True)
    showGridLines = Bool(allow_none=True)
    showRowColHeaders = Bool(allow_none=True)
    showZeros = Bool(allow_none=True)
    rightToLeft = Bool(allow_none=True)
    tabSelected = Bool(allow_none=True)
    showRuler = Bool(allow_none=True)
    showOutlineSymbols = Bool(allow_none=True)
    defaultGridColor = Bool(allow_none=True)
    showWhiteSpace = Bool(allow_none=True)
    view = NoneSet(values=("normal", "pageBreakPreview", "pageLayout"))
    topLeftCell = String(allow_none=True)
    colorId = Integer(allow_none=True)
    zoomScale = Integer(allow_none=True)
    zoomScaleNormal = Integer(allow_none=True)
    zoomScaleSheetLayoutView = Integer(allow_none=True)
    zoomScalePageLayoutView = Integer(allow_none=True)
    zoomToFit = Bool(allow_none=True) # Chart sheets only
    workbookViewId = Integer()
    selection = Sequence(expected_type=Selection)
    pane = Typed(expected_type=Pane, allow_none=True)

    def __init__(self,
                 windowProtection=None,
                 showFormulas=None,
                 showGridLines=None,
                 showRowColHeaders=None,
                 showZeros=None,
                 rightToLeft=None,
                 tabSelected=None,
                 showRuler=None,
                 showOutlineSymbols=None,
                 defaultGridColor=None,
                 showWhiteSpace=None,
                 view=None,
                 topLeftCell=None,
                 colorId=None,
                 zoomScale=None,
                 zoomScaleNormal=None,
                 zoomScaleSheetLayoutView=None,
                 zoomScalePageLayoutView=None,
                 zoomToFit=None,
                 workbookViewId=0,
                 selection=None,
                 pane=None,):
        self.windowProtection = windowProtection
        self.showFormulas = showFormulas
        self.showGridLines = showGridLines
        self.showRowColHeaders = showRowColHeaders
        self.showZeros = showZeros
        self.rightToLeft = rightToLeft
        self.tabSelected = tabSelected
        self.showRuler = showRuler
        self.showOutlineSymbols = showOutlineSymbols
        self.defaultGridColor = defaultGridColor
        self.showWhiteSpace = showWhiteSpace
        self.view = view
        self.topLeftCell = topLeftCell
        self.colorId = colorId
        self.zoomScale = zoomScale
        self.zoomScaleNormal = zoomScaleNormal
        self.zoomScaleSheetLayoutView = zoomScaleSheetLayoutView
        self.zoomScalePageLayoutView = zoomScalePageLayoutView
        self.zoomToFit = zoomToFit
        self.workbookViewId = workbookViewId
        self.pane = pane
        if selection is None:
            selection = (Selection(), )
        self.selection = selection


class SheetViewList(Serialisable):

    tagname = "sheetViews"

    sheetView = Sequence(expected_type=SheetView, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('sheetView',)

    def __init__(self,
                 sheetView=None,
                 extLst=None,
                ):
        if sheetView is None:
            sheetView = [SheetView()]
        self.sheetView = sheetView


    @property
    def active(self):
        """
        Returns the first sheet view which is assumed to be active
        """
        return self.sheetView[0]

''',
    },
    'openpyxl.worksheet.worksheet': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Worksheet is the 2nd-level container in Excel."""


# Python stdlib imports
from itertools import chain
from operator import itemgetter
from inspect import isgenerator
from warnings import warn

# compatibility imports
from openpyxl.compat import (
    deprecated,
)

# package imports
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
    coordinate_to_tuple,
)
from openpyxl.cell import Cell, MergedCell
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.packaging.relationship import RelationshipList
from openpyxl.workbook.child import _WorkbookChild
from openpyxl.workbook.defined_name import (
    DefinedNameDict,
)

from openpyxl.formula.translate import Translator

from .datavalidation import DataValidationList
from .page import (
    PrintPageSetup,
    PageMargins,
    PrintOptions,
)
from .dimensions import (
    ColumnDimension,
    RowDimension,
    DimensionHolder,
    SheetFormatProperties,
)
from .protection import SheetProtection
from .filters import AutoFilter
from .views import (
    Pane,
    Selection,
    SheetViewList,
)
from .cell_range import MultiCellRange, CellRange
from .merge import MergedCellRange
from .properties import WorksheetProperties
from .pagebreak import RowBreak, ColBreak
from .scenario import ScenarioList
from .table import TableList
from .formula import ArrayFormula
from .print_settings import (
    PrintTitles,
    ColRange,
    RowRange,
    PrintArea,
)


class Worksheet(_WorkbookChild):
    """Represents a worksheet.

    Do not create worksheets yourself,
    use :func:`openpyxl.workbook.Workbook.create_sheet` instead

    """

    _rel_type = "worksheet"
    _path = "/xl/worksheets/sheet{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"

    BREAK_NONE = 0
    BREAK_ROW = 1
    BREAK_COLUMN = 2

    SHEETSTATE_VISIBLE = 'visible'
    SHEETSTATE_HIDDEN = 'hidden'
    SHEETSTATE_VERYHIDDEN = 'veryHidden'

    # Paper size
    PAPERSIZE_LETTER = '1'
    PAPERSIZE_LETTER_SMALL = '2'
    PAPERSIZE_TABLOID = '3'
    PAPERSIZE_LEDGER = '4'
    PAPERSIZE_LEGAL = '5'
    PAPERSIZE_STATEMENT = '6'
    PAPERSIZE_EXECUTIVE = '7'
    PAPERSIZE_A3 = '8'
    PAPERSIZE_A4 = '9'
    PAPERSIZE_A4_SMALL = '10'
    PAPERSIZE_A5 = '11'

    # Page orientation
    ORIENTATION_PORTRAIT = 'portrait'
    ORIENTATION_LANDSCAPE = 'landscape'

    def __init__(self, parent, title=None):
        _WorkbookChild.__init__(self, parent, title)
        self._setup()

    def _setup(self):
        self.row_dimensions = DimensionHolder(worksheet=self,
                                              default_factory=self._add_row)
        self.column_dimensions = DimensionHolder(worksheet=self,
                                                 default_factory=self._add_column)
        self.row_breaks = RowBreak()
        self.col_breaks = ColBreak()
        self._cells = {}
        self._charts = []
        self._images = []
        self._rels = RelationshipList()
        self._drawing = None
        self._comments = []
        self.merged_cells = MultiCellRange()
        self._tables = TableList()
        self._pivots = []
        self.data_validations = DataValidationList()
        self._hyperlinks = []
        self.sheet_state = 'visible'
        self.page_setup = PrintPageSetup(worksheet=self)
        self.print_options = PrintOptions()
        self._print_rows = None
        self._print_cols = None
        self._print_area = PrintArea()
        self.page_margins = PageMargins()
        self.views = SheetViewList()
        self.protection = SheetProtection()
        self.defined_names = DefinedNameDict()

        self._current_row = 0
        self.auto_filter = AutoFilter()
        self.conditional_formatting = ConditionalFormattingList()
        self.legacy_drawing = None
        self.sheet_properties = WorksheetProperties()
        self.sheet_format = SheetFormatProperties()
        self.scenarios = ScenarioList()


    @property
    def sheet_view(self):
        return self.views.active


    @property
    def selected_cell(self):
        return self.sheet_view.selection[0].sqref


    @property
    def active_cell(self):
        return self.sheet_view.selection[0].activeCell


    @property
    def array_formulae(self):
        """Returns a dictionary of cells with array formulae and the cells in array"""
        result = {}
        for c in self._cells.values():
            if c.data_type == "f":
                if isinstance(c.value, ArrayFormula):
                    result[c.coordinate] = c.value.ref
        return result


    @property
    def show_gridlines(self):
        return self.sheet_view.showGridLines


    @property
    def freeze_panes(self):
        if self.sheet_view.pane is not None:
            return self.sheet_view.pane.topLeftCell


    @freeze_panes.setter
    def freeze_panes(self, topLeftCell=None):
        if isinstance(topLeftCell, Cell):
            topLeftCell = topLeftCell.coordinate
        if topLeftCell == 'A1':
            topLeftCell = None

        if not topLeftCell:
            self.sheet_view.pane = None
            return

        row, column = coordinate_to_tuple(topLeftCell)

        view = self.sheet_view
        view.pane = Pane(topLeftCell=topLeftCell,
                        activePane="topRight",
                        state="frozen")
        view.selection[0].pane = "topRight"

        if column > 1:
            view.pane.xSplit = column - 1
        if row > 1:
            view.pane.ySplit = row - 1
            view.pane.activePane = 'bottomLeft'
            view.selection[0].pane = "bottomLeft"
            if column > 1:
                view.selection[0].pane = "bottomRight"
                view.pane.activePane = 'bottomRight'

        if row > 1 and column > 1:
            sel = list(view.selection)
            sel.insert(0, Selection(pane="topRight", activeCell=None, sqref=None))
            sel.insert(1, Selection(pane="bottomLeft", activeCell=None, sqref=None))
            view.selection = sel


    def cell(self, row, column, value=None):
        """
        Returns a cell object based on the given coordinates.

        Usage: cell(row=15, column=1, value=5)

        Calling `cell` creates cells in memory when they
        are first accessed.

        :param row: row index of the cell (e.g. 4)
        :type row: int

        :param column: column index of the cell (e.g. 3)
        :type column: int

        :param value: value of the cell (e.g. 5)
        :type value: numeric or time or string or bool or none

        :rtype: openpyxl.cell.cell.Cell
        """

        if row < 1 or column < 1:
            raise ValueError("Row or column values must be at least 1")

        cell = self._get_cell(row, column)
        if value is not None:
            cell.value = value

        return cell


    def _get_cell(self, row, column):
        """
        Internal method for getting a cell from a worksheet.
        Will create a new cell if one doesn't already exist.
        """
        if not 0 < row < 1048577:
            raise ValueError(f"Row numbers must be between 1 and 1048576. Row number supplied was {row}")
        coordinate = (row, column)
        if not coordinate in self._cells:
            cell = Cell(self, row=row, column=column)
            self._add_cell(cell)
        return self._cells[coordinate]


    def _add_cell(self, cell):
        """
        Internal method for adding cell objects.
        """
        column = cell.col_idx
        row = cell.row
        self._current_row = max(row, self._current_row)
        self._cells[(row, column)] = cell


    def __getitem__(self, key):
        """Convenience access by Excel style coordinates

        The key can be a single cell coordinate 'A1', a range of cells 'A1:D25',
        individual rows or columns 'A', 4 or ranges of rows or columns 'A:D',
        4:10.

        Single cells will always be created if they do not exist.

        Returns either a single cell or a tuple of rows or columns.
        """
        if isinstance(key, slice):
            if not all([key.start, key.stop]):
                raise IndexError("{0} is not a valid coordinate or range".format(key))
            key = "{0}:{1}".format(key.start, key.stop)

        if isinstance(key, int):
            key = str(key
                      )
        min_col, min_row, max_col, max_row = range_boundaries(key)

        if not any([min_col, min_row, max_col, max_row]):
            raise IndexError("{0} is not a valid coordinate or range".format(key))

        if min_row is None:
            cols = tuple(self.iter_cols(min_col, max_col))
            if min_col == max_col:
                cols = cols[0]
            return cols
        if min_col is None:
            rows = tuple(self.iter_rows(min_col=min_col, min_row=min_row,
                                        max_col=self.max_column, max_row=max_row))
            if min_row == max_row:
                rows = rows[0]
            return rows
        if ":" not in key:
            return self._get_cell(min_row, min_col)
        return tuple(self.iter_rows(min_row=min_row, min_col=min_col,
                                    max_row=max_row, max_col=max_col))


    def __setitem__(self, key, value):
        self[key].value = value


    def __iter__(self):
        return self.iter_rows()


    def __delitem__(self, key):
        row, column = coordinate_to_tuple(key)
        if (row, column) in self._cells:
            del self._cells[(row, column)]


    @property
    def min_row(self):
        """The minimum row index containing data (1-based)

        :type: int
        """
        min_row = 1
        if self._cells:
            min_row = min(self._cells)[0]
        return min_row


    @property
    def max_row(self):
        """The maximum row index containing data (1-based)

        :type: int
        """
        max_row = 1
        if self._cells:
            max_row = max(self._cells)[0]
        return max_row


    @property
    def min_column(self):
        """The minimum column index containing data (1-based)

        :type: int
        """
        min_col = 1
        if self._cells:
            min_col = min(c[1] for c in self._cells)
        return min_col


    @property
    def max_column(self):
        """The maximum column index containing data (1-based)

        :type: int
        """
        max_col = 1
        if self._cells:
            max_col = max(c[1] for c in self._cells)
        return max_col


    def calculate_dimension(self):
        """Return the minimum bounding range for all cells containing data (ex. 'A1:M24')

        :rtype: string
        """
        if self._cells:
            rows = set()
            cols = set()
            for row, col in self._cells:
                rows.add(row)
                cols.add(col)
            max_row = max(rows)
            max_col = max(cols)
            min_col = min(cols)
            min_row = min(rows)
        else:
            return "A1:A1"

        return f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"


    @property
    def dimensions(self):
        """Returns the result of :func:`calculate_dimension`"""
        return self.calculate_dimension()


    def iter_rows(self, min_row=None, max_row=None, min_col=None, max_col=None, values_only=False):
        """
        Produces cells from the worksheet, by row. Specify the iteration range
        using indices of rows and columns.

        If no indices are specified the range starts at A1.

        If no cells are in the worksheet an empty tuple will be returned.

        :param min_col: smallest column index (1-based index)
        :type min_col: int

        :param min_row: smallest row index (1-based index)
        :type min_row: int

        :param max_col: largest column index (1-based index)
        :type max_col: int

        :param max_row: largest row index (1-based index)
        :type max_row: int

        :param values_only: whether only cell values should be returned
        :type values_only: bool

        :rtype: generator
        """

        if self._current_row == 0 and not any([min_col, min_row, max_col, max_row ]):
            return iter(())


        min_col = min_col or 1
        min_row = min_row or 1
        max_col = max_col or self.max_column
        max_row = max_row or self.max_row

        return self._cells_by_row(min_col, min_row, max_col, max_row, values_only)


    def _cells_by_row(self, min_col, min_row, max_col, max_row, values_only=False):
        for row in range(min_row, max_row + 1):
            cells = (self.cell(row=row, column=column) for column in range(min_col, max_col + 1))
            if values_only:
                yield tuple(cell.value for cell in cells)
            else:
                yield tuple(cells)


    @property
    def rows(self):
        """Produces all cells in the worksheet, by row (see :func:`iter_rows`)

        :type: generator
        """
        return self.iter_rows()


    @property
    def values(self):
        """Produces all cell values in the worksheet, by row

        :type: generator
        """
        for row in self.iter_rows(values_only=True):
            yield row


    def iter_cols(self, min_col=None, max_col=None, min_row=None, max_row=None, values_only=False):
        """
        Produces cells from the worksheet, by column. Specify the iteration range
        using indices of rows and columns.

        If no indices are specified the range starts at A1.

        If no cells are in the worksheet an empty tuple will be returned.

        :param min_col: smallest column index (1-based index)
        :type min_col: int

        :param min_row: smallest row index (1-based index)
        :type min_row: int

        :param max_col: largest column index (1-based index)
        :type max_col: int

        :param max_row: largest row index (1-based index)
        :type max_row: int

        :param values_only: whether only cell values should be returned
        :type values_only: bool

        :rtype: generator
        """

        if self._current_row == 0 and not any([min_col, min_row, max_col, max_row]):
            return iter(())

        min_col = min_col or 1
        min_row = min_row or 1
        max_col = max_col or self.max_column
        max_row = max_row or self.max_row

        return self._cells_by_col(min_col, min_row, max_col, max_row, values_only)


    def _cells_by_col(self, min_col, min_row, max_col, max_row, values_only=False):
        """
        Get cells by column
        """
        for column in range(min_col, max_col+1):
            cells = (self.cell(row=row, column=column)
                        for row in range(min_row, max_row+1))
            if values_only:
                yield tuple(cell.value for cell in cells)
            else:
                yield tuple(cells)


    @property
    def columns(self):
        """Produces all cells in the worksheet, by column  (see :func:`iter_cols`)"""
        return self.iter_cols()


    @property
    def column_groups(self):
        """
        Return a list of column ranges where more than one column
        """
        return [cd.range for cd in self.column_dimensions.values() if cd.min and cd.max > cd.min]


    def set_printer_settings(self, paper_size, orientation):
        """Set printer settings """

        self.page_setup.paperSize = paper_size
        self.page_setup.orientation = orientation


    def add_data_validation(self, data_validation):
        """ Add a data-validation object to the sheet.  The data-validation
            object defines the type of data-validation to be applied and the
            cell or range of cells it should apply to.
        """
        self.data_validations.append(data_validation)


    def add_chart(self, chart, anchor=None):
        """
        Add a chart to the sheet
        Optionally provide a cell for the top-left anchor
        """
        if anchor is not None:
            chart.anchor = anchor
        self._charts.append(chart)


    def add_image(self, img, anchor=None):
        """
        Add an image to the sheet.
        Optionally provide a cell for the top-left anchor
        """
        if anchor is not None:
            img.anchor = anchor
        self._images.append(img)


    def add_table(self, table):
        """
        Check for duplicate name in definedNames and other worksheet tables
        before adding table.
        """

        if self.parent._duplicate_name(table.name):
            raise ValueError("Table with name {0} already exists".format(table.name))
        if not hasattr(self, "_get_cell"):
            warn("In write-only mode you must add table columns manually")
        self._tables.add(table)


    @property
    def tables(self):
        return self._tables


    def add_pivot(self, pivot):
        self._pivots.append(pivot)


    def merge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        """ Set merge on a cell range.  Range is a cell range (e.g. A1:E1) """
        if range_string is None:
            cr = CellRange(range_string=range_string, min_col=start_column, min_row=start_row,
                      max_col=end_column, max_row=end_row)
            range_string = cr.coord
        mcr = MergedCellRange(self, range_string)
        self.merged_cells.add(mcr)
        self._clean_merge_range(mcr)


    def _clean_merge_range(self, mcr):
        """
        Remove all but the top left-cell from a range of merged cells
        and recreate the lost border information.
        Borders are then applied
        """
        cells = mcr.cells
        next(cells) # skip first cell
        for row, col in cells:
            self._cells[row, col] = MergedCell(self, row, col)
        mcr.format()


    @property
    @deprecated("Use ws.merged_cells.ranges")
    def merged_cell_ranges(self):
        """Return a copy of cell ranges"""
        return self.merged_cells.ranges[:]


    def unmerge_cells(self, range_string=None, start_row=None, start_column=None, end_row=None, end_column=None):
        """ Remove merge on a cell range.  Range is a cell range (e.g. A1:E1) """
        cr = CellRange(range_string=range_string, min_col=start_column, min_row=start_row,
                      max_col=end_column, max_row=end_row)

        if cr.coord not in self.merged_cells:
            raise ValueError("Cell range {0} is not merged".format(cr.coord))

        self.merged_cells.remove(cr)

        cells = cr.cells
        next(cells) # skip first cell
        for row, col in cells:
            del self._cells[(row, col)]


    def append(self, iterable):
        """Appends a group of values at the bottom of the current sheet.

        * If it's a list: all values are added in order, starting from the first column
        * If it's a dict: values are assigned to the columns indicated by the keys (numbers or letters)

        :param iterable: list, range or generator, or dict containing values to append
        :type iterable: list|tuple|range|generator or dict

        Usage:

        * append(['This is A1', 'This is B1', 'This is C1'])
        * **or** append({'A' : 'This is A1', 'C' : 'This is C1'})
        * **or** append({1 : 'This is A1', 3 : 'This is C1'})

        :raise: TypeError when iterable is neither a list/tuple nor a dict

        """
        row_idx = self._current_row + 1

        if (isinstance(iterable, (list, tuple, range))
            or isgenerator(iterable)):
            for col_idx, content in enumerate(iterable, 1):
                if isinstance(content, Cell):
                    # compatible with write-only mode
                    cell = content
                    if cell.parent and cell.parent != self:
                        raise ValueError("Cells cannot be copied from other worksheets")
                    cell.parent = self
                    cell.column = col_idx
                    cell.row = row_idx
                else:
                    cell = Cell(self, row=row_idx, column=col_idx, value=content)
                self._cells[(row_idx, col_idx)] = cell

        elif isinstance(iterable, dict):
            for col_idx, content in iterable.items():
                if isinstance(col_idx, str):
                    col_idx = column_index_from_string(col_idx)
                cell = Cell(self, row=row_idx, column=col_idx, value=content)
                self._cells[(row_idx, col_idx)] = cell

        else:
            self._invalid_row(iterable)

        self._current_row = row_idx


    def _move_cells(self, min_row=None, min_col=None, offset=0, row_or_col="row"):
        """
        Move either rows or columns around by the offset
        """
        reverse = offset > 0 # start at the end if inserting
        row_offset = 0
        col_offset = 0

        # need to make affected ranges contiguous
        if row_or_col == 'row':
            cells = self.iter_rows(min_row=min_row)
            row_offset = offset
            key = 0
        else:
            cells = self.iter_cols(min_col=min_col)
            col_offset = offset
            key = 1
        cells = list(cells)

        for row, column in sorted(self._cells, key=itemgetter(key), reverse=reverse):
            if min_row and row < min_row:
                continue
            elif min_col and column < min_col:
                continue

            self._move_cell(row, column, row_offset, col_offset)


    def insert_rows(self, idx, amount=1):
        """
        Insert row or rows before row==idx
        """
        self._move_cells(min_row=idx, offset=amount, row_or_col="row")
        self._current_row = self.max_row


    def insert_cols(self, idx, amount=1):
        """
        Insert column or columns before col==idx
        """
        self._move_cells(min_col=idx, offset=amount, row_or_col="column")


    def delete_rows(self, idx, amount=1):
        """
        Delete row or rows from row==idx
        """

        remainder = _gutter(idx, amount, self.max_row)

        self._move_cells(min_row=idx+amount, offset=-amount, row_or_col="row")

        # calculating min and max col is an expensive operation, do it only once
        min_col = self.min_column
        max_col = self.max_column + 1
        for row in remainder:
            for col in range(min_col, max_col):
                if (row, col) in self._cells:
                    del self._cells[row, col]
        self._current_row = self.max_row
        if not self._cells:
            self._current_row = 0


    def delete_cols(self, idx, amount=1):
        """
        Delete column or columns from col==idx
        """

        remainder = _gutter(idx, amount, self.max_column)

        self._move_cells(min_col=idx+amount, offset=-amount, row_or_col="column")

        # calculating min and max row is an expensive operation, do it only once
        min_row = self.min_row
        max_row = self.max_row + 1
        for col in remainder:
            for row in range(min_row, max_row):
                if (row, col) in self._cells:
                    del self._cells[row, col]


    def move_range(self, cell_range, rows=0, cols=0, translate=False):
        """
        Move a cell range by the number of rows and/or columns:
        down if rows > 0 and up if rows < 0
        right if cols > 0 and left if cols < 0
        Existing cells will be overwritten.
        Formulae and references will not be updated.
        """
        if isinstance(cell_range, str):
            cell_range = CellRange(cell_range)
        if not isinstance(cell_range, CellRange):
            raise ValueError("Only CellRange objects can be moved")
        if not rows and not cols:
            return

        down = rows > 0
        right = cols > 0

        if rows:
            cells = sorted(cell_range.rows, reverse=down)
        else:
            cells = sorted(cell_range.cols, reverse=right)

        for row, col in chain.from_iterable(cells):
            self._move_cell(row, col, rows, cols, translate)

        # rebase moved range
        cell_range.shift(row_shift=rows, col_shift=cols)


    def _move_cell(self, row, column, row_offset, col_offset, translate=False):
        """
        Move a cell from one place to another.
        Delete at old index
        Rebase coordinate
        """
        cell = self._get_cell(row, column)
        new_row = cell.row + row_offset
        new_col = cell.column + col_offset
        self._cells[new_row, new_col] = cell
        del self._cells[(cell.row, cell.column)]
        cell.row = new_row
        cell.column = new_col
        if translate and cell.data_type == "f":
            t = Translator(cell.value, cell.coordinate)
            cell.value = t.translate_formula(row_delta=row_offset, col_delta=col_offset)


    def _invalid_row(self, iterable):
        raise TypeError('Value must be a list, tuple, range or generator, or a dict. Supplied value is {0}'.format(
            type(iterable))
                        )


    def _add_column(self):
        """Dimension factory for column information"""

        return ColumnDimension(self)

    def _add_row(self):
        """Dimension factory for row information"""

        return RowDimension(self)


    @property
    def print_title_rows(self):
        """Rows to be printed at the top of every page (ex: '1:3')"""
        if self._print_rows:
            return str(self._print_rows)


    @print_title_rows.setter
    def print_title_rows(self, rows):
        """
        Set rows to be printed on the top of every page
        format `1:3`
        """
        if rows is not None:
            self._print_rows = RowRange(rows)


    @property
    def print_title_cols(self):
        """Columns to be printed at the left side of every page (ex: 'A:C')"""
        if self._print_cols:
            return str(self._print_cols)


    @print_title_cols.setter
    def print_title_cols(self, cols):
        """
        Set cols to be printed on the left of every page
        format ``A:C`
        """
        if cols is not None:
            self._print_cols = ColRange(cols)


    @property
    def print_titles(self):
        titles = PrintTitles(cols=self._print_cols, rows=self._print_rows, title=self.title)
        return str(titles)


    @property
    def print_area(self):
        """
        The print area for the worksheet, or None if not set. To set, supply a range
        like 'A1:D4' or a list of ranges.
        """
        self._print_area.title = self.title
        return str(self._print_area)


    @print_area.setter
    def print_area(self, value):
        """
        Range of cells in the form A1:D4 or list of ranges. Print area can be cleared
        by passing `None` or an empty list
        """
        if not value:
            self._print_area = PrintArea()
        elif isinstance(value, str):
            self._print_area = PrintArea.from_string(value)
        elif hasattr(value, "__iter__"):
            self._print_area = PrintArea.from_string(",".join(value))


def _gutter(idx, offset, max_val):
    """
    When deleting rows and columns are deleted we rely on overwriting.
    This may not be the case for a large offset on small set of cells:
    range(cells_to_delete) > range(cell_to_be_moved)
    """
    gutter = range(max(max_val+1-offset, idx), min(idx+offset, max_val)+1)
    return gutter

''',
    },
    'openpyxl.writer': {
        'is_package': True,
        'source': r"""
# Copyright (c) 2010-2024 openpyxl

""",
    },
    'openpyxl.writer.excel': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


# Python stdlib imports
import datetime
import re
from zipfile import ZipFile, ZIP_DEFLATED

# package imports
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.xml.constants import (
    ARC_ROOT_RELS,
    ARC_WORKBOOK_RELS,
    ARC_APP,
    ARC_CORE,
    ARC_CUSTOM,
    CPROPS_TYPE,
    ARC_THEME,
    ARC_STYLE,
    ARC_WORKBOOK,
    )
from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.xml.functions import tostring, fromstring
from openpyxl.packaging.manifest import Manifest
from openpyxl.packaging.relationship import (
    get_rels_path,
    RelationshipList,
    Relationship,
)
from openpyxl.comments.comment_sheet import CommentSheet
from openpyxl.styles.stylesheet import write_stylesheet
from openpyxl.worksheet._writer import WorksheetWriter
from openpyxl.workbook._writer import WorkbookWriter
from .theme import theme_xml


class ExcelWriter:
    """Write a workbook object to an Excel file."""

    def __init__(self, workbook, archive):
        self._archive = archive
        self.workbook = workbook
        self.manifest = Manifest()
        self.vba_modified = set()
        self._tables = []
        self._charts = []
        self._images = []
        self._drawings = []
        self._comments = []
        self._pivots = []


    def write_data(self):
        from openpyxl.packaging.extended import ExtendedProperties
        """Write the various xml files into the zip archive."""
        # cleanup all worksheets
        archive = self._archive

        props = ExtendedProperties()
        archive.writestr(ARC_APP, tostring(props.to_tree()))

        archive.writestr(ARC_CORE, tostring(self.workbook.properties.to_tree()))
        if self.workbook.loaded_theme:
            archive.writestr(ARC_THEME, self.workbook.loaded_theme)
        else:
            archive.writestr(ARC_THEME, theme_xml)

        if len(self.workbook.custom_doc_props) >= 1:
            archive.writestr(ARC_CUSTOM, tostring(self.workbook.custom_doc_props.to_tree()))
            class CustomOverride():
                path = "/" + ARC_CUSTOM #PartName
                mime_type = CPROPS_TYPE #ContentType

            custom_override = CustomOverride()
            self.manifest.append(custom_override)

        self._write_worksheets()
        self._write_chartsheets()
        self._write_images()
        self._write_charts()

        self._write_external_links()

        stylesheet = write_stylesheet(self.workbook)
        archive.writestr(ARC_STYLE, tostring(stylesheet))

        writer = WorkbookWriter(self.workbook)
        archive.writestr(ARC_ROOT_RELS, writer.write_root_rels())
        archive.writestr(ARC_WORKBOOK, writer.write())
        archive.writestr(ARC_WORKBOOK_RELS, writer.write_rels())

        self._merge_vba()

        self.manifest._write(archive, self.workbook)

    def _merge_vba(self):
        """
        If workbook contains macros then extract associated files from cache
        of old file and add to archive
        """
        ARC_VBA = re.compile("|".join(
            ('xl/vba', r'xl/drawings/.*vmlDrawing\d\.vml',
             'xl/ctrlProps', 'customUI', 'xl/activeX', r'xl/media/.*\.emf')
        )
                             )

        if self.workbook.vba_archive:
            for name in set(self.workbook.vba_archive.namelist()) - self.vba_modified:
                if ARC_VBA.match(name):
                    self._archive.writestr(name, self.workbook.vba_archive.read(name))


    def _write_images(self):
        # delegate to object
        for img in self._images:
            self._archive.writestr(img.path[1:], img._data())


    def _write_charts(self):
        # delegate to object
        if len(self._charts) != len(set(self._charts)):
            raise InvalidFileException("The same chart cannot be used in more than one worksheet")
        for chart in self._charts:
            self._archive.writestr(chart.path[1:], tostring(chart._write()))
            self.manifest.append(chart)


    def _write_drawing(self, drawing):
        """
        Write a drawing
        """
        self._drawings.append(drawing)
        drawing._id = len(self._drawings)
        for chart in drawing.charts:
            self._charts.append(chart)
            chart._id = len(self._charts)
        for img in drawing.images:
            self._images.append(img)
            img._id = len(self._images)
        rels_path = get_rels_path(drawing.path)[1:]
        self._archive.writestr(drawing.path[1:], tostring(drawing._write()))
        self._archive.writestr(rels_path, tostring(drawing._write_rels()))
        self.manifest.append(drawing)


    def _write_chartsheets(self):
        for idx, sheet in enumerate(self.workbook.chartsheets, 1):

            sheet._id = idx
            xml = tostring(sheet.to_tree())

            self._archive.writestr(sheet.path[1:], xml)
            self.manifest.append(sheet)

            if sheet._drawing:
                self._write_drawing(sheet._drawing)

                rel = Relationship(type="drawing", Target=sheet._drawing.path)
                rels = RelationshipList()
                rels.append(rel)
                tree = rels.to_tree()

                rels_path = get_rels_path(sheet.path[1:])
                self._archive.writestr(rels_path, tostring(tree))


    def _write_comment(self, ws):

        cs = CommentSheet.from_comments(ws._comments)
        self._comments.append(cs)
        cs._id = len(self._comments)
        self._archive.writestr(cs.path[1:], tostring(cs.to_tree()))
        self.manifest.append(cs)

        if ws.legacy_drawing is None or self.workbook.vba_archive is None:
            ws.legacy_drawing = 'xl/drawings/commentsDrawing{0}.vml'.format(cs._id)
            vml = None
        else:
            vml = fromstring(self.workbook.vba_archive.read(ws.legacy_drawing))

        vml = cs.write_shapes(vml)

        self._archive.writestr(ws.legacy_drawing, vml)
        self.vba_modified.add(ws.legacy_drawing)

        comment_rel = Relationship(Id="comments", type=cs._rel_type, Target=cs.path)
        ws._rels.append(comment_rel)


    def write_worksheet(self, ws):
        ws._drawing = SpreadsheetDrawing()
        ws._drawing.charts = ws._charts
        ws._drawing.images = ws._images
        if self.workbook.write_only:
            if not ws.closed:
                ws.close()
            writer = ws._writer
        else:
            writer = WorksheetWriter(ws)
            writer.write()

        ws._rels = writer._rels
        self._archive.write(writer.out, ws.path[1:])
        self.manifest.append(ws)
        writer.cleanup()


    def _write_worksheets(self):

        pivot_caches = set()

        for idx, ws in enumerate(self.workbook.worksheets, 1):

            ws._id = idx
            self.write_worksheet(ws)

            if ws._drawing:
                self._write_drawing(ws._drawing)

                for r in ws._rels:
                    if "drawing" in r.Type:
                        r.Target = ws._drawing.path

            if ws._comments:
                self._write_comment(ws)

            if ws.legacy_drawing is not None:
                shape_rel = Relationship(type="vmlDrawing", Id="anysvml",
                                         Target="/" + ws.legacy_drawing)
                ws._rels.append(shape_rel)

            for t in ws._tables.values():
                self._tables.append(t)
                t.id = len(self._tables)
                t._write(self._archive)
                self.manifest.append(t)
                ws._rels.get(t._rel_id).Target = t.path

            for p in ws._pivots:
                if p.cache not in pivot_caches:
                    pivot_caches.add(p.cache)
                    p.cache._id = len(pivot_caches)

                self._pivots.append(p)
                p._id = len(self._pivots)
                p._write(self._archive, self.manifest)
                self.workbook._pivots.append(p)
                r = Relationship(Type=p.rel_type, Target=p.path)
                ws._rels.append(r)

            if ws._rels:
                tree = ws._rels.to_tree()
                rels_path = get_rels_path(ws.path)[1:]
                self._archive.writestr(rels_path, tostring(tree))


    def _write_external_links(self):
        # delegate to object
        """Write links to external workbooks"""
        wb = self.workbook
        for idx, link in enumerate(wb._external_links, 1):
            link._id = idx
            rels_path = get_rels_path(link.path[1:])

            xml = link.to_tree()
            self._archive.writestr(link.path[1:], tostring(xml))
            rels = RelationshipList()
            rels.append(link.file_link)
            self._archive.writestr(rels_path, tostring(rels.to_tree()))
            self.manifest.append(link)


    def save(self):
        """Write data into the archive."""
        self.write_data()
        self._archive.close()


def save_workbook(workbook, filename):
    """Save the given workbook on the filesystem under the name filename.

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool

    """
    archive = ZipFile(filename, 'w', ZIP_DEFLATED, allowZip64=True)
    workbook.properties.modified = datetime.datetime.now(tz=datetime.timezone.utc).replace(tzinfo=None)
    writer = ExcelWriter(workbook, archive)
    writer.save()
    return True

''',
    },
    'openpyxl.writer.theme': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""Write the theme xml based on a fixed string."""


theme_xml = """<?xml version="1.0"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000"/>
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF"/>
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D"/>
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1"/>
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD"/>
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D"/>
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59"/>
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2"/>
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6"/>
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646"/>
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF"/>
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080"/>
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="&#xFF2D;&#xFF33; &#xFF30;&#x30B4;&#x30B7;&#x30C3;&#x30AF;"/>
        <a:font script="Hang" typeface="&#xB9D1;&#xC740; &#xACE0;&#xB515;"/>
        <a:font script="Hans" typeface="&#x5B8B;&#x4F53;"/>
        <a:font script="Hant" typeface="&#x65B0;&#x7D30;&#x660E;&#x9AD4;"/>
        <a:font script="Arab" typeface="Times New Roman"/>
        <a:font script="Hebr" typeface="Times New Roman"/>
        <a:font script="Thai" typeface="Tahoma"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="MoolBoran"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Times New Roman"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
        <a:font script="Jpan" typeface="&#xFF2D;&#xFF33; &#xFF30;&#x30B4;&#x30B7;&#x30C3;&#x30AF;"/>
        <a:font script="Hang" typeface="&#xB9D1;&#xC740; &#xACE0;&#xB515;"/>
        <a:font script="Hans" typeface="&#x5B8B;&#x4F53;"/>
        <a:font script="Hant" typeface="&#x65B0;&#x7D30;&#x660E;&#x9AD4;"/>
        <a:font script="Arab" typeface="Arial"/>
        <a:font script="Hebr" typeface="Arial"/>
        <a:font script="Thai" typeface="Tahoma"/>
        <a:font script="Ethi" typeface="Nyala"/>
        <a:font script="Beng" typeface="Vrinda"/>
        <a:font script="Gujr" typeface="Shruti"/>
        <a:font script="Khmr" typeface="DaunPenh"/>
        <a:font script="Knda" typeface="Tunga"/>
        <a:font script="Guru" typeface="Raavi"/>
        <a:font script="Cans" typeface="Euphemia"/>
        <a:font script="Cher" typeface="Plantagenet Cherokee"/>
        <a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
        <a:font script="Tibt" typeface="Microsoft Himalaya"/>
        <a:font script="Thaa" typeface="MV Boli"/>
        <a:font script="Deva" typeface="Mangal"/>
        <a:font script="Telu" typeface="Gautami"/>
        <a:font script="Taml" typeface="Latha"/>
        <a:font script="Syrc" typeface="Estrangelo Edessa"/>
        <a:font script="Orya" typeface="Kalinga"/>
        <a:font script="Mlym" typeface="Kartika"/>
        <a:font script="Laoo" typeface="DokChampa"/>
        <a:font script="Sinh" typeface="Iskoola Pota"/>
        <a:font script="Mong" typeface="Mongolian Baiti"/>
        <a:font script="Viet" typeface="Arial"/>
        <a:font script="Uigh" typeface="Microsoft Uighur"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:shade val="51000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="80000">
              <a:schemeClr val="phClr">
                <a:shade val="93000"/>
                <a:satMod val="130000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="94000"/>
                <a:satMod val="135000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000"/>
              <a:satMod val="105000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr"/>
          </a:solidFill>
          <a:prstDash val="solid"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="38000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000"/>
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront">
              <a:rot lat="0" lon="0" rev="0"/>
            </a:camera>
            <a:lightRig rig="threePt" dir="t">
              <a:rot lat="0" lon="0" rev="1200000"/>
            </a:lightRig>
          </a:scene3d>
          <a:sp3d>
            <a:bevelT w="63500" h="25400"/>
          </a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr"/>
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000"/>
                <a:shade val="99000"/>
                <a:satMod val="350000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000"/>
                <a:satMod val="255000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000"/>
                <a:satMod val="300000"/>
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000"/>
                <a:satMod val="200000"/>
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>
"""

def write_theme():
    """Write the theme xml."""
    return theme_xml

''',
    },
    'openpyxl.xml': {
        'is_package': True,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


"""Collection of XML resources compatible across different Python versions"""
import os


def lxml_available():
    try:
        from lxml.etree import LXML_VERSION
        LXML = LXML_VERSION >= (3, 3, 1, 0)
        if not LXML:
            import warnings
            warnings.warn("The installed version of lxml is too old to be used with openpyxl")
            return False  # we have it, but too old
        else:
            return True  # we have it, and recent enough
    except ImportError:
        return False  # we don't even have it


def lxml_env_set():
    return os.environ.get("OPENPYXL_LXML", "True") == "True"


LXML = lxml_available() and lxml_env_set()


def defusedxml_available():
    try:
        import defusedxml # noqa
    except ImportError:
        return False
    else:
        return True


def defusedxml_env_set():
    return os.environ.get("OPENPYXL_DEFUSEDXML", "True") == "True"


DEFUSEDXML = defusedxml_available() and defusedxml_env_set()

''',
    },
    'openpyxl.xml.constants': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl


"""Constants for fixed paths in a file and xml namespace urls."""

MIN_ROW = 0
MIN_COLUMN = 0
MAX_COLUMN = 16384
MAX_ROW = 1048576

# constants
PACKAGE_PROPS = 'docProps'
PACKAGE_XL = 'xl'
PACKAGE_RELS = '_rels'
PACKAGE_THEME = PACKAGE_XL + '/' + 'theme'
PACKAGE_WORKSHEETS = PACKAGE_XL + '/' + 'worksheets'
PACKAGE_CHARTSHEETS = PACKAGE_XL + '/' + 'chartsheets'
PACKAGE_DRAWINGS = PACKAGE_XL + '/' + 'drawings'
PACKAGE_CHARTS = PACKAGE_XL + '/' + 'charts'
PACKAGE_IMAGES = PACKAGE_XL + '/' + 'media'
PACKAGE_WORKSHEET_RELS = PACKAGE_WORKSHEETS + '/' + '_rels'
PACKAGE_CHARTSHEETS_RELS = PACKAGE_CHARTSHEETS + '/' + '_rels'
PACKAGE_PIVOT_TABLE = PACKAGE_XL + '/' + 'pivotTables'
PACKAGE_PIVOT_CACHE = PACKAGE_XL + '/' + 'pivotCache'

ARC_CONTENT_TYPES = '[Content_Types].xml'
ARC_ROOT_RELS = PACKAGE_RELS + '/.rels'
ARC_WORKBOOK_RELS = PACKAGE_XL + '/' + PACKAGE_RELS + '/workbook.xml.rels'
ARC_CORE = PACKAGE_PROPS + '/core.xml'
ARC_APP = PACKAGE_PROPS + '/app.xml'
ARC_CUSTOM = PACKAGE_PROPS + '/custom.xml'
ARC_WORKBOOK = PACKAGE_XL + '/workbook.xml'
ARC_STYLE = PACKAGE_XL + '/styles.xml'
ARC_THEME = PACKAGE_THEME + '/theme1.xml'
ARC_SHARED_STRINGS = PACKAGE_XL + '/sharedStrings.xml'
ARC_CUSTOM_UI = 'customUI/customUI.xml'

## namespaces
# XML
XML_NS = "http://www.w3.org/XML/1998/namespace"
# Dublin Core
DCORE_NS = 'http://purl.org/dc/elements/1.1/'
DCTERMS_NS = 'http://purl.org/dc/terms/'
DCTERMS_PREFIX = 'dcterms'

# Document
DOC_NS = "http://schemas.openxmlformats.org/officeDocument/2006/"
REL_NS = DOC_NS + "relationships"
COMMENTS_NS = REL_NS + "/comments"
IMAGE_NS = REL_NS + "/image"
VML_NS =  REL_NS + "/vmlDrawing"
VTYPES_NS = DOC_NS + 'docPropsVTypes'
XPROPS_NS = DOC_NS + 'extended-properties'
CUSTPROPS_NS = DOC_NS + 'custom-properties'
EXTERNAL_LINK_NS = REL_NS + "/externalLink"

# CustomDocumentProperty FMTID:
CPROPS_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"

# Package
PKG_NS = "http://schemas.openxmlformats.org/package/2006/"
PKG_REL_NS = PKG_NS + "relationships"
COREPROPS_NS = PKG_NS + 'metadata/core-properties'
CONTYPES_NS = PKG_NS + 'content-types'

XSI_NS = 'http://www.w3.org/2001/XMLSchema-instance'
XML_NS = 'http://www.w3.org/XML/1998/namespace'
SHEET_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

# Drawing
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CHART_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"

CUSTOMUI_NS = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'


NAMESPACES = {
    'cp': COREPROPS_NS,
    'dc': DCORE_NS,
    DCTERMS_PREFIX: DCTERMS_NS,
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'xsi': XSI_NS,
    'vt': VTYPES_NS,
    'xml': XML_NS,
    'main': SHEET_MAIN_NS,
    'cust': CUSTPROPS_NS,
}

## Mime types
WORKBOOK_MACRO = "application/vnd.ms-excel.%s.macroEnabled.main+xml"
WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.%s.main+xml"
SPREADSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.%s+xml"
SHARED_STRINGS = SPREADSHEET % "sharedStrings"
EXTERNAL_LINK = SPREADSHEET % "externalLink"
WORKSHEET_TYPE = SPREADSHEET % "worksheet"
COMMENTS_TYPE = SPREADSHEET % "comments"
STYLES_TYPE = SPREADSHEET % "styles"
CHARTSHEET_TYPE = SPREADSHEET % "chartsheet"
DRAWING_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
CHART_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CHARTSHAPE_TYPE = "application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"
THEME_TYPE = "application/vnd.openxmlformats-officedocument.theme+xml"
CPROPS_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
XLTM = WORKBOOK_MACRO % 'template'
XLSM = WORKBOOK_MACRO % 'sheet'
XLTX = WORKBOOK % 'template'
XLSX = WORKBOOK % 'sheet'


# Extensions to the specification

EXT_TYPES = {
    '{78C0D931-6437-407D-A8EE-F0AAD7539E65}': 'Conditional Formatting',
    '{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}': 'Data Validation',
    '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}': 'Sparkline Group',
    '{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}': 'Slicer List',
    '{FC87AEE6-9EDD-4A0A-B7FB-166176984837}': 'Protected Range',
    '{01252117-D84E-4E92-8308-4BE1C098FCBB}': 'Ignored Error',
    '{F7C9EE02-42E1-4005-9D12-6889AFFD525C}': 'Web Extension',
    '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}': 'Slicer List',
    '{7E03D99C-DC04-49d9-9315-930204A7B6E9}': 'Timeline Ref',
}

# Objects related to macros that we preserve
CTRL = "application/vnd.ms-excel.controlproperties+xml"
ACTIVEX = "application/vnd.ms-office.activeX+xml"
VBA = "application/vnd.ms-office.vbaProject"

''',
    },
    'openpyxl.xml.functions': {
        'is_package': False,
        'source': r'''
# Copyright (c) 2010-2024 openpyxl

"""
XML compatibility functions
"""

# Python stdlib imports
import re
from functools import partial

from openpyxl import DEFUSEDXML, LXML

if LXML is True:
    from lxml.etree import (
    Element,
    SubElement,
    register_namespace,
    QName,
    xmlfile,
    XMLParser,
    )
    from lxml.etree import fromstring, tostring
    # do not resolve entities
    safe_parser = XMLParser(resolve_entities=False)
    fromstring = partial(fromstring, parser=safe_parser)

else:
    from xml.etree.ElementTree import (
    Element,
    SubElement,
    fromstring,
    tostring,
    QName,
    register_namespace
    )
    from et_xmlfile import xmlfile
    if DEFUSEDXML is True:
        from defusedxml.ElementTree import fromstring

from xml.etree.ElementTree import iterparse
if DEFUSEDXML is True:
    from defusedxml.ElementTree import iterparse

from openpyxl.xml.constants import (
    CHART_NS,
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_DRAWING_NS,
    SHEET_MAIN_NS,
    REL_NS,
    VTYPES_NS,
    COREPROPS_NS,
    CUSTPROPS_NS,
    DCTERMS_NS,
    DCTERMS_PREFIX,
    XML_NS
)

register_namespace(DCTERMS_PREFIX, DCTERMS_NS)
register_namespace('dcmitype', 'http://purl.org/dc/dcmitype/')
register_namespace('cp', COREPROPS_NS)
register_namespace('c', CHART_NS)
register_namespace('a', DRAWING_NS)
register_namespace('s', SHEET_MAIN_NS)
register_namespace('r', REL_NS)
register_namespace('vt', VTYPES_NS)
register_namespace('xdr', SHEET_DRAWING_NS)
register_namespace('cdr', CHART_DRAWING_NS)
register_namespace('xml', XML_NS)
register_namespace('cust', CUSTPROPS_NS)


tostring = partial(tostring, encoding="utf-8")

NS_REGEX = re.compile("({(?P<namespace>.*)})?(?P<localname>.*)")

def localname(node):
    if callable(node.tag):
        return "comment"
    m = NS_REGEX.match(node.tag)
    return m.group('localname')


def whitespace(node):
    stripped = node.text.strip()
    if stripped and node.text != stripped:
        node.set("{%s}space" % XML_NS, "preserve")

''',
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

