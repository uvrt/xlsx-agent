"""The streaming parser's options are ooxml-edit's ``parse_xml``'s (ROADMAP.md, O6): the same
parser options, and the same behaviour on entities, entity bombs and whitespace."""

import inspect
import re

import pytest
from lxml import etree
from ooxml_edit.xml import parse_xml

from xlsx_agent.xml import SAFE_OPTIONS, iterparse_xml


def _streamed_root(data: bytes) -> etree._Element:
    root = None
    for _, node in iterparse_xml(data):
        root = node
    return root


def test_every_option_parse_xml_sets_is_set_alike():
    source = inspect.getsource(parse_xml)
    options = dict(re.findall(r"(\w+)=(True|False)", source.split("XMLParser(", 1)[1].split(")", 1)[0]))
    assert options, "parse_xml's parser options were not found"
    for name, value in options.items():
        assert SAFE_OPTIONS[name] is (value == "True"), name
    assert SAFE_OPTIONS["huge_tree"] is False
    assert SAFE_OPTIONS["resolve_entities"] is False


def test_an_internal_entity_is_left_alike():
    data = b'<?xml version="1.0"?><!DOCTYPE a [<!ENTITY e "expanded">]><a><t>&e;</t>  <u/></a>'
    whole, streamed = parse_xml(data), _streamed_root(data)
    assert etree.tostring(whole) == etree.tostring(streamed)
    assert b"  <u/>" in etree.tostring(streamed)  # whitespace kept


def test_an_external_entity_is_never_read(tmp_path):
    secret = tmp_path / "secret.txt"
    secret.write_text("SECRET-CONTENT")
    data = (b'<?xml version="1.0"?><!DOCTYPE a [<!ENTITY x SYSTEM "' + secret.as_uri().encode()
            + b'">]><a>&x;</a>')
    for root in (parse_xml(data), _streamed_root(data)):
        assert b"SECRET-CONTENT" not in etree.tostring(root)


def test_an_entity_bomb_is_refused_alike():
    entities = b'<!ENTITY a0 "lol">' + b"".join(
        b'<!ENTITY a%d "%s">' % (n, (b"&a%d;" % (n - 1)) * 10) for n in range(1, 10))
    data = b'<?xml version="1.0"?><!DOCTYPE a [' + entities + b"]><a>&a9;</a>"
    with pytest.raises(etree.XMLSyntaxError):
        parse_xml(data)
    with pytest.raises(etree.XMLSyntaxError):
        _streamed_root(data)
