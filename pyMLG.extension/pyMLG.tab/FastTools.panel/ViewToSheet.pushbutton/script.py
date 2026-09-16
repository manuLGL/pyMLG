# -*- coding: utf-8 -*-
__title__ = "ViewToSheet"
__doc__ = ("Erstellt für jede ausgewählte Ansicht einen Plan: Ansichtsvorlage, "
           "Plankopf und Position (X/Y) werden im Dialog gewählt.")

import os

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import forms, script

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

KEINE_VORLAGE = "<Keine Vorlage – Ansicht unverändert>"
CM_PRO_FUSS = 30.48


#_________________________________________________________________________
# Daten sammeln
#_________________________________________________________________________

def platzierbar(el):
    return (isinstance(el, View)
            and not el.IsTemplate
            and el.CanBePrinted
            and not isinstance(el, ViewSheet)
            and el.ViewType != ViewType.Schedule)


def gewaehlte_ansichten():
    views = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    views = [v for v in views if platzierbar(v)]
    if views:
        return views
    return forms.select_views(title="Ansichten für neue Pläne wählen",
                              multiple=True,
                              filterfunc=platzierbar) or []


def ansichtsvorlagen():
    vorlagen = {}
    for v in FilteredElementCollector(doc).OfClass(View):
        if v.IsTemplate:
            vorlagen[v.Name] = v
    return vorlagen


def planfamilien():
    familien = {}
    typen = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsElementType()
    for tb in typen:
        name = tb.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
        if name:
            familien[name] = tb
    return familien


def zahl(text):
    return float(text.strip().replace(",", "."))


#_________________________________________________________________________
# Dialog
#_________________________________________________________________________

class ViewToSheetDialog(forms.WPFWindow):
    def __init__(self, anzahl, vorlagen, familien, cfg):
        forms.WPFWindow.__init__(
            self, os.path.join(os.path.dirname(__file__), "ui.xaml"))
        self.ergebnis = None
        self.lbl_info.Text = "{} Ansicht(en) ausgewählt.".format(anzahl)

        self.cb_template.ItemsSource = [KEINE_VORLAGE] + sorted(vorlagen)
        self.cb_template.SelectedItem = cfg.get_option("vorlage", KEINE_VORLAGE)
        if self.cb_template.SelectedIndex < 0:
            self.cb_template.SelectedIndex = 0

        self.cb_titleblock.ItemsSource = sorted(familien)
        self.cb_titleblock.SelectedItem = cfg.get_option("planfamilie", "")
        if self.cb_titleblock.SelectedIndex < 0:
            self.cb_titleblock.SelectedIndex = 0

        self.tb_x.Text = cfg.get_option("x_cm", "-57")
        self.tb_y.Text = cfg.get_option("y_cm", "40")
        self.tb_prefix.Text = cfg.get_option("praefix", "AP")

        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click

    def ok_click(self, sender, args):
        try:
            x = zahl(self.tb_x.Text)
            y = zahl(self.tb_y.Text)
        except ValueError:
            forms.alert("X und Y müssen Zahlen sein (in cm).", title="Eingabe prüfen")
            return
        self.ergebnis = {
            "vorlage": self.cb_template.SelectedItem,
            "planfamilie": self.cb_titleblock.SelectedItem,
            "x_cm": x,
            "y_cm": y,
            "praefix": self.tb_prefix.Text.strip(),
        }
        self.Close()

    def cancel_click(self, sender, args):
        self.Close()


#_________________________________________________________________________
# Pläne erstellen
#_________________________________________________________________________

def eindeutige_nummer(praefix, vergeben):
    num = 1
    while True:
        nummer = "{}-{:03d}".format(praefix, num) if praefix else "{:03d}".format(num)
        if nummer not in vergeben:
            vergeben.add(nummer)
            return nummer
        num += 1


def plaene_erstellen(views, vorlage, planfamilie, punkt, praefix):
    vergeben = {s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet)}
    erstellt = []
    fehler = []

    t = Transaction(doc, "Ansichten auf Pläne")
    t.Start()
    for view in views:
        try:
            sheet = ViewSheet.Create(doc, planfamilie.Id)
            if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
                doc.Delete(sheet.Id)
                fehler.append("{}: bereits auf einem Plan platziert".format(view.Name))
                continue

            if vorlage is not None:
                try:
                    view.ViewTemplateId = vorlage.Id
                except Exception as e:
                    fehler.append("{}: Vorlage nicht anwendbar ({})".format(view.Name, e))

            sheet.Name = view.Name
            sheet.SheetNumber = eindeutige_nummer(praefix, vergeben)
            Viewport.Create(doc, sheet.Id, view.Id, punkt)
            erstellt.append(sheet.SheetNumber)
        except Exception as e:
            fehler.append("{}: {}".format(view.Name, e))
    t.Commit()
    return erstellt, fehler


def main():
    views = gewaehlte_ansichten()
    if not views:
        forms.alert("Keine platzierbaren Ansichten ausgewählt.", title="ViewToSheet")
        return

    familien = planfamilien()
    if not familien:
        forms.alert("Keine Planfamilie (Plankopf) im Projekt gefunden.", title="ViewToSheet")
        return
    vorlagen = ansichtsvorlagen()

    cfg = script.get_config()
    dlg = ViewToSheetDialog(len(views), vorlagen, familien, cfg)
    dlg.ShowDialog()
    wahl = dlg.ergebnis
    if not wahl:
        return

    cfg.vorlage = wahl["vorlage"]
    cfg.planfamilie = wahl["planfamilie"]
    cfg.x_cm = str(wahl["x_cm"])
    cfg.y_cm = str(wahl["y_cm"])
    cfg.praefix = wahl["praefix"]
    script.save_config()

    punkt = XYZ(wahl["x_cm"] / CM_PRO_FUSS, wahl["y_cm"] / CM_PRO_FUSS, 0)
    erstellt, fehler = plaene_erstellen(
        views,
        vorlagen.get(wahl["vorlage"]),
        familien[wahl["planfamilie"]],
        punkt,
        wahl["praefix"])

    meldung = "{} Plan/Pläne erstellt.".format(len(erstellt))
    if erstellt:
        meldung += "\n" + ", ".join(erstellt)
    if fehler:
        meldung += "\n\nHinweise:\n" + "\n".join(fehler)
    forms.alert(meldung, title="ViewToSheet")


main()
