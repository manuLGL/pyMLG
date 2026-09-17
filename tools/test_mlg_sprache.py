# -*- coding: utf-8 -*-
"""Offline-Test der Sprachwahl (ohne Revit).

    python tools\\test_mlg_sprache.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

import mlg_sprache as sp  # noqa: E402


def test_aus_name():
    assert sp._aus_name("German") == sp.DE
    assert sp._aus_name("LanguageType.German") == sp.DE
    assert sp._aus_name("Spanish") == sp.ES
    assert sp._aus_name("English_USA") == sp.EN
    assert sp._aus_name("French") == sp.EN
    assert sp._aus_name("") == sp.EN


def test_t():
    sp.setze_sprache(sp.DE)
    assert sp.t(u"Schließen", u"Close", u"Cerrar") == u"Schließen"
    sp.setze_sprache(sp.EN)
    assert sp.t(u"Schließen", u"Close", u"Cerrar") == u"Close"
    sp.setze_sprache(sp.ES)
    assert sp.t(u"Schließen", u"Close", u"Cerrar") == u"Cerrar"
    # fehlende Übersetzung: Englisch, dann Deutsch
    assert sp.t(u"Schließen", u"Close") == u"Close"
    assert sp.t(u"Schließen") == u"Schließen"


def test_xaml():
    sp.setze_sprache(sp.EN)
    xaml = u'<Button Content="{{ok}}" Style="{StaticResource x}" ' \
           u'ToolTip="{{tipp}}" Tag="{{fehlt}}"/>'
    ergebnis = sp.uebersetze_xaml(xaml, {
        "ok": (u"Übernehmen", u"Apply", u"Aplicar"),
        "tipp": (u"a", u"\"A\" & <B>", u"c")})
    assert ergebnis == (u'<Button Content="Apply" Style="{StaticResource x}" '
                        u'ToolTip="&quot;A&quot; &amp; &lt;B&gt;" '
                        u'Tag="{{fehlt}}"/>')


def test_umgebungsvariable():
    os.environ["PYMLG_SPRACHE"] = "es"
    try:
        assert sp._ermittle() == sp.ES
    finally:
        del os.environ["PYMLG_SPRACHE"]


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
