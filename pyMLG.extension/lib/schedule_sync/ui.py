# -*- coding: utf-8 -*-
"""Dialoge, die unter der CPython3-Engine funktionieren.

Hintergrund: pyRevit stellt `pyrevit.forms` unter CPython nur als Platzhalter
bereit. Jeder Zugriff - auch auf `alert`, `save_file` oder `pick_file` - wirft
dort PyRevitCPythonNotSupported, weil die Formulare per IronPython-XAML-Loader
gebaut werden, den es unter pythonnet nicht gibt.

Dieses Modul ersetzt die benötigten Dialoge:
    * Meldungen/Rückfragen -> Autodesk.Revit.UI.TaskDialog (reine Revit-API)
    * Mehrfachauswahl      -> WinForms-Fenster mit CheckedListBox
    * Datei öffnen/speichern -> WinForms-Dateidialoge, ersatzweise Microsoft.Win32
    * Fortschritt          -> Fortschrittsbalken des pyRevit-Ausgabefensters

Alles ist reiner Code ohne XAML und läuft deshalb unter beiden Engines.
"""

import clr

from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
)

EXCEL_FILTER = u"Excel-Arbeitsmappe (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*"

# TaskDialog.MainContent bei sehr langen Texten kürzen - lange Listen gehören
# ins Ausgabefenster, nicht in einen modalen Dialog.
MAX_TEXTLAENGE = 1800

_winforms = None


def _lade_winforms():
    """Lädt System.Windows.Forms einmalig (oder None, falls nicht verfügbar)."""
    global _winforms
    if _winforms is not None:
        return _winforms if _winforms is not False else None
    try:
        clr.AddReference("System.Windows.Forms")
        clr.AddReference("System.Drawing")
        import System.Drawing as Drawing
        import System.Windows.Forms as Forms
        _winforms = (Forms, Drawing)
    except Exception:
        _winforms = False
        return None
    return _winforms


def _kuerze(text):
    if text and len(text) > MAX_TEXTLAENGE:
        return text[:MAX_TEXTLAENGE] + u"\n\n[...] Vollständige Angaben im " \
                                       u"pyRevit-Ausgabefenster."
    return text


def _warnsymbol(dialog):
    """Warnsymbol setzen, ohne sich auf einen Enum-Namen festzulegen."""
    try:
        from Autodesk.Revit.UI import TaskDialogIcon
        dialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning
    except Exception:
        pass


def meldung(text, titel=u"pyMLG ScheduleSync", hauptzeile=None, warnung=False):
    """Einfache Hinweismeldung mit Schliessen-Schaltfläche."""
    dialog = TaskDialog(titel)
    dialog.TitleAutoPrefix = False
    dialog.MainInstruction = hauptzeile or titel
    dialog.MainContent = _kuerze(text)
    dialog.CommonButtons = TaskDialogCommonButtons.Close
    if warnung:
        _warnsymbol(dialog)
    dialog.Show()


def frage(text, titel=u"pyMLG ScheduleSync", hauptzeile=None, warnung=False,
          standard_ja=False):
    """Ja/Nein-Rückfrage. Rückgabe True bei 'Ja'."""
    dialog = TaskDialog(titel)
    dialog.TitleAutoPrefix = False
    dialog.MainInstruction = hauptzeile or titel
    dialog.MainContent = _kuerze(text)
    dialog.CommonButtons = (TaskDialogCommonButtons.Yes
                            | TaskDialogCommonButtons.No)
    dialog.DefaultButton = (TaskDialogResult.Yes if standard_ja
                            else TaskDialogResult.No)
    if warnung:
        _warnsymbol(dialog)
    return dialog.Show() == TaskDialogResult.Yes


def waehle_option(optionen, titel=u"pyMLG", hauptzeile=None, text=None):
    """Einfachauswahl als TaskDialog mit bis zu vier Befehlslinks.

    Ersetzt forms.CommandSwitchWindow / Radiobutton-Dialoge.
    optionen: Liste von (beschriftung, erlaeuterung).
    Rückgabe: Index der gewählten Option oder None bei Abbruch.
    """
    from Autodesk.Revit.UI import TaskDialogCommandLinkId

    links = [TaskDialogCommandLinkId.CommandLink1,
             TaskDialogCommandLinkId.CommandLink2,
             TaskDialogCommandLinkId.CommandLink3,
             TaskDialogCommandLinkId.CommandLink4]
    ergebnisse = [TaskDialogResult.CommandLink1,
                  TaskDialogResult.CommandLink2,
                  TaskDialogResult.CommandLink3,
                  TaskDialogResult.CommandLink4]

    dialog = TaskDialog(titel)
    dialog.TitleAutoPrefix = False
    dialog.MainInstruction = hauptzeile or titel
    if text:
        dialog.MainContent = _kuerze(text)
    for index, (beschriftung, erlaeuterung) in enumerate(optionen[:4]):
        if erlaeuterung:
            dialog.AddCommandLink(links[index], beschriftung, erlaeuterung)
        else:
            dialog.AddCommandLink(links[index], beschriftung)
    dialog.CommonButtons = TaskDialogCommonButtons.Cancel

    ergebnis = dialog.Show()
    for index, kandidat in enumerate(ergebnisse):
        if ergebnis == kandidat:
            return index
    return None


# ---------------------------------------------------------------------------
# Mehrfachauswahl
# ---------------------------------------------------------------------------

def baue_auswahlfenster(eintraege, titel=u"Auswahl", hinweis=None,
                        schaltflaeche=u"OK"):
    """Baut das Auswahlfenster, ohne es anzuzeigen.

    Rückgabe: (fenster, hole_auswahl, steuerelemente)
    Getrennt von waehle_mehrfach, damit der Fensteraufbau ohne modalen
    Dialog geprüft werden kann.
    """
    module = _lade_winforms()
    if module is None:
        raise RuntimeError(
            u"Die Auswahlliste konnte nicht geöffnet werden: "
            u"System.Windows.Forms ist in dieser Umgebung nicht verfügbar.")
    Forms, Drawing = module

    eintraege = list(eintraege)
    gewaehlt = set()

    fenster = Forms.Form()
    fenster.Text = titel
    fenster.ClientSize = Drawing.Size(520, 600)
    fenster.StartPosition = Forms.FormStartPosition.CenterScreen
    fenster.MinimizeBox = False
    fenster.MaximizeBox = False
    fenster.ShowInTaskbar = False
    # Ohne TopMost verschwindet das Fenster gelegentlich hinter Revit
    fenster.TopMost = True
    fenster.Font = Drawing.Font("Segoe UI", 9.0)

    anker_oben = (Forms.AnchorStyles.Top | Forms.AnchorStyles.Left
                  | Forms.AnchorStyles.Right)

    beschriftung = Forms.Label()
    beschriftung.Text = hinweis or u"Einträge auswählen:"
    beschriftung.Bounds = Drawing.Rectangle(12, 10, 496, 34)
    beschriftung.Anchor = anker_oben
    fenster.Controls.Add(beschriftung)

    suchfeld = Forms.TextBox()
    suchfeld.Bounds = Drawing.Rectangle(12, 48, 496, 24)
    suchfeld.Anchor = anker_oben
    fenster.Controls.Add(suchfeld)

    liste = Forms.CheckedListBox()
    liste.Bounds = Drawing.Rectangle(12, 80, 496, 462)
    liste.Anchor = (Forms.AnchorStyles.Top | Forms.AnchorStyles.Bottom
                    | Forms.AnchorStyles.Left | Forms.AnchorStyles.Right)
    liste.CheckOnClick = True
    liste.IntegralHeight = False
    fenster.Controls.Add(liste)

    def uebernehme_haken():
        """Aktuellen Zustand der sichtbaren Zeilen in die Auswahl übernehmen."""
        for index in range(liste.Items.Count):
            eintrag = liste.Items[index]
            if liste.GetItemChecked(index):
                gewaehlt.add(eintrag)
            else:
                gewaehlt.discard(eintrag)

    def fuelle(filtertext=u""):
        filtertext = (filtertext or u"").strip().lower()
        liste.BeginUpdate()
        liste.Items.Clear()
        for eintrag in eintraege:
            if filtertext and filtertext not in eintrag.lower():
                continue
            index = liste.Items.Add(eintrag)
            if eintrag in gewaehlt:
                liste.SetItemChecked(index, True)
        liste.EndUpdate()

    def bei_suche(sender, args):
        uebernehme_haken()
        fuelle(suchfeld.Text)

    suchfeld.TextChanged += bei_suche

    def setze_alle(zustand):
        for index in range(liste.Items.Count):
            liste.SetItemChecked(index, zustand)
        uebernehme_haken()

    anker_unten_links = Forms.AnchorStyles.Bottom | Forms.AnchorStyles.Left
    anker_unten_rechts = Forms.AnchorStyles.Bottom | Forms.AnchorStyles.Right

    knopf_alle = Forms.Button()
    knopf_alle.Text = u"Alle"
    knopf_alle.Bounds = Drawing.Rectangle(12, 554, 80, 30)
    knopf_alle.Anchor = anker_unten_links
    knopf_alle.Click += lambda sender, args: setze_alle(True)
    fenster.Controls.Add(knopf_alle)

    knopf_keine = Forms.Button()
    knopf_keine.Text = u"Keine"
    knopf_keine.Bounds = Drawing.Rectangle(98, 554, 80, 30)
    knopf_keine.Anchor = anker_unten_links
    knopf_keine.Click += lambda sender, args: setze_alle(False)
    fenster.Controls.Add(knopf_keine)

    knopf_ok = Forms.Button()
    knopf_ok.Text = schaltflaeche
    knopf_ok.Bounds = Drawing.Rectangle(280, 554, 140, 30)
    knopf_ok.Anchor = anker_unten_rechts
    knopf_ok.DialogResult = Forms.DialogResult.OK
    fenster.Controls.Add(knopf_ok)

    knopf_abbruch = Forms.Button()
    knopf_abbruch.Text = u"Abbrechen"
    knopf_abbruch.Bounds = Drawing.Rectangle(428, 554, 80, 30)
    knopf_abbruch.Anchor = anker_unten_rechts
    knopf_abbruch.DialogResult = Forms.DialogResult.Cancel
    fenster.Controls.Add(knopf_abbruch)

    fenster.AcceptButton = knopf_ok
    fenster.CancelButton = knopf_abbruch

    fuelle()

    def hole_auswahl():
        uebernehme_haken()
        # Ursprüngliche Reihenfolge beibehalten
        return [eintrag for eintrag in eintraege if eintrag in gewaehlt]

    steuerelemente = {
        "liste": liste,
        "suchfeld": suchfeld,
        "knopf_alle": knopf_alle,
        "knopf_keine": knopf_keine,
        "knopf_ok": knopf_ok,
        "knopf_abbruch": knopf_abbruch,
    }
    return fenster, hole_auswahl, steuerelemente


def waehle_mehrfach(eintraege, titel=u"Auswahl", hinweis=None,
                    schaltflaeche=u"OK"):
    """Mehrfachauswahl aus einer Liste von Texten.

    Rückgabe: Liste der gewählten Texte oder None bei Abbruch.
    """
    Forms, _Drawing = _lade_winforms() or (None, None)
    fenster, hole_auswahl, _steuerelemente = baue_auswahlfenster(
        eintraege, titel=titel, hinweis=hinweis, schaltflaeche=schaltflaeche)
    try:
        if fenster.ShowDialog() != Forms.DialogResult.OK:
            return None
        return hole_auswahl()
    finally:
        fenster.Dispose()


# ---------------------------------------------------------------------------
# Dateidialoge
# ---------------------------------------------------------------------------

def _dialog_win32(speichern, vorgabename, startordner, titel):
    """Ersatzweg über Microsoft.Win32 (WPF), falls WinForms fehlt."""
    try:
        clr.AddReference("PresentationFramework")
        from Microsoft.Win32 import OpenFileDialog, SaveFileDialog
    except Exception:
        return None
    dialog = SaveFileDialog() if speichern else OpenFileDialog()
    dialog.Title = titel
    dialog.Filter = EXCEL_FILTER
    dialog.DefaultExt = u".xlsx"
    if vorgabename:
        dialog.FileName = vorgabename
    if startordner:
        dialog.InitialDirectory = startordner
    if dialog.ShowDialog() == True:  # noqa: E712 - Nullable<bool> aus .NET
        return dialog.FileName
    return None


def _dateidialog(speichern, vorgabename, startordner, titel):
    module = _lade_winforms()
    if module is None:
        return _dialog_win32(speichern, vorgabename, startordner, titel)
    Forms, _Drawing = module

    dialog = Forms.SaveFileDialog() if speichern else Forms.OpenFileDialog()
    dialog.Title = titel
    dialog.Filter = EXCEL_FILTER
    dialog.DefaultExt = u"xlsx"
    dialog.AddExtension = True
    dialog.RestoreDirectory = True
    if vorgabename:
        dialog.FileName = vorgabename
    if startordner:
        dialog.InitialDirectory = startordner
    if speichern:
        dialog.OverwritePrompt = True
    else:
        dialog.CheckFileExists = True
        dialog.Multiselect = False

    try:
        if dialog.ShowDialog() != Forms.DialogResult.OK:
            return None
        return dialog.FileName
    finally:
        dialog.Dispose()


def datei_speichern(vorgabename=u"", startordner=u"",
                    titel=u"Excel-Datei speichern"):
    """Speichern-Dialog. Rückgabe: Pfad oder None."""
    return _dateidialog(True, vorgabename, startordner, titel)


def datei_oeffnen(startordner=u"", titel=u"Excel-Datei auswählen"):
    """Öffnen-Dialog. Rückgabe: Pfad oder None."""
    return _dateidialog(False, u"", startordner, titel)


# ---------------------------------------------------------------------------
# Fortschritt
# ---------------------------------------------------------------------------

class Fortschritt(object):
    """Fortschrittsbalken des pyRevit-Ausgabefensters als Kontextmanager.

    Ersetzt forms.ProgressBar, das unter CPython nicht verfügbar ist.
    Ein Abbrechen durch den Nutzer bietet das Ausgabefenster nicht an.
    """

    def __init__(self, output, max_wert=1):
        self._output = output
        self._max = max(1, max_wert)

    def __enter__(self):
        try:
            self._output.unhide_progress()
        except Exception:
            pass
        return self

    def aktualisiere(self, wert, max_wert=None):
        if max_wert:
            self._max = max(1, max_wert)
        try:
            self._output.update_progress(wert, self._max)
        except Exception:
            pass

    def __exit__(self, *_ausnahme):
        try:
            self._output.reset_progress()
            self._output.hide_progress()
        except Exception:
            pass
        return False
