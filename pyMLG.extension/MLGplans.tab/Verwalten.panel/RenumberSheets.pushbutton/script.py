# -*- coding: utf-8 -*-
# Nummeriert die ausgewählten Pläne (Projektbrowser-Auswahl oder Liste) ab
# einer Startnummer neu - mit Vorschau. Die Nummern werden erst auf
# Zwischenwerte und dann auf die neuen Werte gesetzt, damit Tauschen und
# Verschieben innerhalb der Auswahl ohne Konflikte geht. Konflikte mit
# anderen Plänen werden vorher angezeigt.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "RenumberSheets"
__author__ = "Manuel"

import io
import os

from Autodesk.Revit.DB import FilteredElementCollector, Transaction, ViewSheet
from pyrevit import forms, revit

from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_sprache import t, uebersetze_xaml

TITEL = t(u"Pläne nummerieren", u"Renumber Sheets", u"Renumerar planos")
NACH_NUMMER = t(u"nach aktueller Nummer", u"by current number", u"por número actual")
NACH_NAME = t(u"nach Planname", u"by sheet name", u"por nombre de plano")

XAML_TEXTE = {
    "titel": (u"Pläne nummerieren", u"Renumber Sheets", u"Renumerar planos"),
    "start": (u"Startnummer", u"Start number", u"Número inicial"),
    "schritt": (u"Schritt", u"Step", u"Incremento"),
    "reihenfolge": (u"Reihenfolge", u"Order", u"Orden"),
    "ok": (u"Nummerieren", u"Renumber", u"Renumerar"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}


def plaene_waehlen(doc, uidoc):
    gewaehlt = rv.gewaehlte_plaene(doc, uidoc)
    if len(gewaehlt) > 1:
        return gewaehlt
    alle = rv.plaene(doc)
    texte = forms.SelectFromList.show(
        [rv.plan_text(p) for p in alle], multiselect=True,
        title=t(u"Pläne zum Nummerieren wählen", u"Select sheets to renumber",
                u"Seleccione planos para renumerar"),
        button_name=t(u"Weiter", u"Next", u"Siguiente"))
    if not texte:
        return []
    texte = set(texte)
    return [p for p in alle if rv.plan_text(p) in texte]


class NummerierenDialog(forms.WPFWindow):
    def __init__(self, plaene, fremde):
        pfad = os.path.join(os.path.dirname(__file__), "ui.xaml")
        with io.open(pfad, encoding="utf-8") as datei:
            xaml = uebersetze_xaml(datei.read(), XAML_TEXTE)
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.plaene = plaene
        self.fremde = fremde
        self.ergebnis = None
        self._vorschau = None

        self.lbl_info.Text = t(
            u"{} Plan/Pläne ausgewählt. Die letzte Zahl der Startnummer wird "
            u"hochgezählt, führende Nullen bleiben (A-090, A-100 ...).",
            u"{} sheet(s) selected. The last number of the start number is "
            u"counted up, leading zeros are kept (A-090, A-100 ...).",
            u"{} plano(s) seleccionados. Se incrementa el último número del "
            u"número inicial y se conservan los ceros (A-090, A-100 ...).").format(len(plaene))
        self.tb_start.Text = plaene[0].SheetNumber
        self.tb_schritt.Text = u"1"
        self.cb_reihenfolge.ItemsSource = [NACH_NUMMER, NACH_NAME]
        self.cb_reihenfolge.SelectedIndex = 0

        self.tb_start.TextChanged += self.aktualisieren
        self.tb_schritt.TextChanged += self.aktualisieren
        self.cb_reihenfolge.SelectionChanged += self.aktualisieren
        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click
        self.aktualisieren(None, None)

    def _reihenfolge(self):
        if self.cb_reihenfolge.SelectedItem == NACH_NAME:
            return sorted(self.plaene, key=lambda p: lg.natuerlich(p.Name))
        return sorted(self.plaene, key=lambda p: lg.natuerlich(p.SheetNumber))

    def aktualisieren(self, sender, args):
        self._vorschau = None
        self.btn_ok.IsEnabled = False
        start = self.tb_start.Text.strip()
        if not start:
            self.lb_vorschau.ItemsSource = []
            self.lbl_status.Text = t(u"Bitte eine Startnummer angeben.",
                                     u"Please enter a start number.",
                                     u"Indique un número inicial.")
            return
        try:
            schritt = int(self.tb_schritt.Text.strip())
            plaene = self._reihenfolge()
            neue = lg.umnummerierung(len(plaene), start, schritt)
        except ValueError:
            self.lb_vorschau.ItemsSource = []
            self.lbl_status.Text = t(
                u"Schritt muss eine ganze Zahl ungleich 0 sein, und keine "
                u"Nummer darf negativ werden.",
                u"Step must be a whole number other than 0, and no number may "
                u"become negative.",
                u"El incremento debe ser un entero distinto de 0 y ningún "
                u"número puede ser negativo.")
            return

        breite = max(len(p.SheetNumber) for p in plaene)
        zeilen = [u"{}  ->  {}   {}".format(p.SheetNumber.ljust(breite), neu, p.Name)
                  for p, neu in zip(plaene, neue)]
        self.lb_vorschau.ItemsSource = zeilen
        doppelt = lg.konflikte(neue, self.fremde)
        if doppelt:
            self.lbl_status.Text = t(u"Schon von anderen Plänen belegt: {}",
                                     u"Already used by other sheets: {}",
                                     u"Ya usados por otros planos: {}").format(
                u", ".join(doppelt[:10]))
            return
        self.lbl_status.Text = u""
        self._vorschau = list(zip(plaene, neue))
        self.btn_ok.IsEnabled = True

    def ok_click(self, sender, args):
        if self._vorschau:
            self.ergebnis = self._vorschau
            self.Close()

    def cancel_click(self, sender, args):
        self.Close()


def main():
    doc, uidoc = revit.doc, revit.uidoc
    plaene = plaene_waehlen(doc, uidoc)
    if not plaene:
        return
    eigene = set(rv.id_wert(p.Id) for p in plaene)
    fremde = lg.nummern_menge(s.SheetNumber for s in
                              FilteredElementCollector(doc).OfClass(ViewSheet)
                              if rv.id_wert(s.Id) not in eigene)

    dlg = NummerierenDialog(plaene, fremde)
    dlg.ShowDialog()
    paare = dlg.ergebnis
    if not paare:
        return
    paare = [(p, neu) for p, neu in paare if p.SheetNumber != neu]
    if not paare:
        return

    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        # Zwischenwerte, damit z.B. A-101 und A-102 tauschen können
        for plan, _ in paare:
            plan.SheetNumber = u"~mlg~{}".format(rv.id_wert(plan.Id))
        for plan, neu in paare:
            plan.SheetNumber = neu
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Nummerieren fehlgeschlagen, nichts wurde geändert.",
                      u"Renumbering failed, nothing was changed.",
                      u"Error al renumerar, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    forms.alert(t(u"{} Plan/Pläne umnummeriert.", u"{} sheet(s) renumbered.",
                  u"{} plano(s) renumerados.").format(len(paare)), title=TITEL)


main()
