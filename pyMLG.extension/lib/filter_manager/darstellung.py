# -*- coding: utf-8 -*-
"""Sichtbarkeit und Grafiküberschreibungen eines Filters in einer Ansicht.

Bildet die Filterzeile aus "Sichtbarkeit/Grafiken > Filter" nach:

    Sichtbarkeit | Projektion/Oberfläche: Linien, Muster, Transparenz
                 | Schnitt: Linien, Muster | Halbton

Jedes Feld hat drei Zustände:
    UNVERAENDERT   nichts schreiben - vorhandene Überschreibung bleibt
    KEINE          Überschreibung entfernen (Revit-Standard)
    Wert           Überschreibung setzen

Geschrieben wird über View.SetFilterOverrides(). Ausgangspunkt ist eine
Kopie von View.GetFilterOverrides(), damit unberührte Felder erhalten
bleiben (außer "Überschreibungen zurücksetzen" ist gewählt).

Musterfarbe ohne Muster: Revit zeigt eine Farbe nur mit Füllmuster. Hat
der Filter noch kein Muster, wird "Einfarbige Füllung" (IsSolidFill)
gesetzt - wie beim Farbwählen im nativen Dialog.
"""

from Autodesk.Revit.DB import (
    Color,
    ElementId,
    FillPatternElement,
    FillPatternTarget,
    FilteredElementCollector,
    OverrideGraphicSettings,
)
from System.Windows import FontWeights, Thickness, VerticalAlignment, Visibility
from System.Windows.Controls import (
    Border,
    Button,
    ComboBox,
    ComboBoxItem,
    Grid,
    RowDefinition,
    StackPanel,
    TextBlock,
    Orientation,
)
from System.Windows.Media import SolidColorBrush

from filter_manager import dialoge as dlg
from filter_manager.aufloesung import id_wert
from mlg_sprache import t, uebersetze_xaml

UNVERAENDERT = "unveraendert"
KEINE = "keine"

# Feldschlüssel
SICHTBAR = "sichtbar"
PROJ_LINIE_FARBE = "proj_linie_farbe"
PROJ_LINIE_STAERKE = "proj_linie_staerke"
OBERFL_MUSTER_FARBE = "oberfl_muster_farbe"
OBERFL_MUSTER = "oberfl_muster"
TRANSPARENZ = "transparenz"
SCHNITT_LINIE_FARBE = "schnitt_linie_farbe"
SCHNITT_LINIE_STAERKE = "schnitt_linie_staerke"
SCHNITT_MUSTER_FARBE = "schnitt_muster_farbe"
SCHNITT_MUSTER = "schnitt_muster"
HALBTON = "halbton"

UEBERSCHREIBUNGEN = (PROJ_LINIE_FARBE, PROJ_LINIE_STAERKE, OBERFL_MUSTER_FARBE,
                     OBERFL_MUSTER, TRANSPARENZ, SCHNITT_LINIE_FARBE,
                     SCHNITT_LINIE_STAERKE, SCHNITT_MUSTER_FARBE,
                     SCHNITT_MUSTER, HALBTON)


class Darstellung(object):
    """Gewählte Änderungen. Farben als (r, g, b), Muster als Id-Zahlenwert."""

    def __init__(self):
        self.werte = dict((k, UNVERAENDERT) for k in
                          (SICHTBAR,) + UEBERSCHREIBUNGEN)
        self.zuruecksetzen = False

    def sichtbarkeit(self):
        """True/False oder None (nicht ändern)."""
        wert = self.werte[SICHTBAR]
        return None if wert == UNVERAENDERT else bool(wert)

    def aendert_grafik(self):
        return self.zuruecksetzen or any(
            self.werte[k] != UNVERAENDERT for k in UEBERSCHREIBUNGEN)

    def zusammenfassung(self):
        """Kurztext der Änderungen, z.B. für Rückfragen und Berichte."""
        teile = []
        if self.werte[SICHTBAR] != UNVERAENDERT:
            teile.append(t(u"sichtbar", u"visible", u"visible")
                         if self.werte[SICHTBAR]
                         else t(u"unsichtbar", u"hidden", u"oculto"))
        if self.zuruecksetzen:
            teile.append(t(u"Überschreibungen zurückgesetzt",
                           u"overrides reset", u"modificaciones restablecidas"))
        anzahl = sum(1 for k in UEBERSCHREIBUNGEN
                     if self.werte[k] != UNVERAENDERT)
        if anzahl:
            teile.append(t(u"%d Grafikeinstellung(en)",
                           u"%d graphic setting(s)",
                           u"%d ajuste(s) gráfico(s)") % anzahl)
        return u", ".join(teile)

    # -- Anwenden -------------------------------------------------------

    def anwenden(self, ansicht, filter_id, doc):
        """Sichtbarkeit und Überschreibungen in die Ansicht schreiben."""
        sichtbar = self.sichtbarkeit()
        if sichtbar is not None:
            ansicht.SetFilterVisibility(filter_id, sichtbar)
        if not self.aendert_grafik():
            return
        if self.zuruecksetzen:
            ogs = OverrideGraphicSettings()
        else:
            ogs = OverrideGraphicSettings(ansicht.GetFilterOverrides(filter_id))
        self.uebertrage(ogs, doc)
        ansicht.SetFilterOverrides(filter_id, ogs)

    def uebertrage(self, ogs, doc):
        w = self.werte

        def farbe(wert):
            if wert == KEINE:
                return Color.InvalidColorValue
            return Color(wert[0], wert[1], wert[2])

        def staerke(wert):
            return OverrideGraphicSettings.InvalidPenNumber if wert == KEINE \
                else int(wert)

        def muster(wert):
            return ElementId.InvalidElementId if wert == KEINE \
                else ElementId(wert)

        if w[PROJ_LINIE_FARBE] != UNVERAENDERT:
            ogs.SetProjectionLineColor(farbe(w[PROJ_LINIE_FARBE]))
        if w[PROJ_LINIE_STAERKE] != UNVERAENDERT:
            ogs.SetProjectionLineWeight(staerke(w[PROJ_LINIE_STAERKE]))
        if w[SCHNITT_LINIE_FARBE] != UNVERAENDERT:
            ogs.SetCutLineColor(farbe(w[SCHNITT_LINIE_FARBE]))
        if w[SCHNITT_LINIE_STAERKE] != UNVERAENDERT:
            ogs.SetCutLineWeight(staerke(w[SCHNITT_LINIE_STAERKE]))

        # Oberflächenmuster (Vordergrund)
        if w[OBERFL_MUSTER] != UNVERAENDERT:
            ogs.SetSurfaceForegroundPatternId(muster(w[OBERFL_MUSTER]))
        if w[OBERFL_MUSTER_FARBE] != UNVERAENDERT:
            ogs.SetSurfaceForegroundPatternColor(farbe(w[OBERFL_MUSTER_FARBE]))
            if w[OBERFL_MUSTER_FARBE] != KEINE and w[OBERFL_MUSTER] \
                    == UNVERAENDERT and id_wert(
                        ogs.SurfaceForegroundPatternId) in (None, -1):
                ogs.SetSurfaceForegroundPatternId(volle_fuellung(doc))
        if w[OBERFL_MUSTER] != UNVERAENDERT \
                or w[OBERFL_MUSTER_FARBE] not in (UNVERAENDERT, KEINE):
            ogs.SetSurfaceForegroundPatternVisible(True)

        # Schnittmuster (Vordergrund)
        if w[SCHNITT_MUSTER] != UNVERAENDERT:
            ogs.SetCutForegroundPatternId(muster(w[SCHNITT_MUSTER]))
        if w[SCHNITT_MUSTER_FARBE] != UNVERAENDERT:
            ogs.SetCutForegroundPatternColor(farbe(w[SCHNITT_MUSTER_FARBE]))
            if w[SCHNITT_MUSTER_FARBE] != KEINE and w[SCHNITT_MUSTER] \
                    == UNVERAENDERT and id_wert(
                        ogs.CutForegroundPatternId) in (None, -1):
                ogs.SetCutForegroundPatternId(volle_fuellung(doc))
        if w[SCHNITT_MUSTER] != UNVERAENDERT \
                or w[SCHNITT_MUSTER_FARBE] not in (UNVERAENDERT, KEINE):
            ogs.SetCutForegroundPatternVisible(True)

        if w[TRANSPARENZ] != UNVERAENDERT:
            ogs.SetSurfaceTransparency(0 if w[TRANSPARENZ] == KEINE
                                       else int(w[TRANSPARENZ]))
        if w[HALBTON] != UNVERAENDERT:
            ogs.SetHalftone(w[HALBTON] is True)


# ---------------------------------------------------------------------------
# Revit-Daten
# ---------------------------------------------------------------------------

def fuellmuster(doc):
    """[(Name, Id-Zahlenwert)] der Zeichnungs-Füllmuster, Vollfüllung zuerst."""
    eintraege = []
    for element in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            muster = element.GetFillPattern()
            if muster.Target != FillPatternTarget.Drafting:
                continue
            eintraege.append((0 if muster.IsSolidFill else 1,
                              element.Name, id_wert(element.Id)))
        except Exception:
            continue
    eintraege.sort(key=lambda e: (e[0], e[1].lower()))
    return [(name, wert) for _v, name, wert in eintraege]


def volle_fuellung(doc):
    for element in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            muster = element.GetFillPattern()
            if muster.IsSolidFill and muster.Target == FillPatternTarget.Drafting:
                return element.Id
        except Exception:
            continue
    return ElementId.InvalidElementId


def _farbe_aus(revit_farbe):
    try:
        if revit_farbe is not None and revit_farbe.IsValid:
            return (revit_farbe.Red, revit_farbe.Green, revit_farbe.Blue)
    except Exception:
        pass
    return None


def farbe_waehlen(besitzer, vorgabe=None):
    """Revit-Farbdialog (wie in Sichtbarkeit/Grafiken). Rückgabe (r,g,b)/None.

    Ersatzweg WinForms-ColorDialog, falls der Revit-Dialog nicht verfügbar ist.
    """
    try:
        from Autodesk.Revit.UI import ColorSelectionDialog, \
            ItemSelectionDialogResult
        dialog = ColorSelectionDialog()
        if vorgabe:
            dialog.OriginalColor = Color(vorgabe[0], vorgabe[1], vorgabe[2])
        if dialog.Show() != ItemSelectionDialogResult.Confirmed:
            return None
        return _farbe_aus(dialog.SelectedColor)
    except ImportError:
        pass
    import clr
    clr.AddReference("System.Windows.Forms")
    clr.AddReference("System.Drawing")
    from System.Windows.Forms import ColorDialog, DialogResult
    from System.Drawing import Color as FarbeGdi
    dialog = ColorDialog()
    dialog.FullOpen = True
    if vorgabe:
        dialog.Color = FarbeGdi.FromArgb(vorgabe[0], vorgabe[1], vorgabe[2])
    if dialog.ShowDialog() != DialogResult.OK:
        return None
    return (dialog.Color.R, dialog.Color.G, dialog.Color.B)


# ---------------------------------------------------------------------------
# Dialog
# ---------------------------------------------------------------------------

_TEXTE = {
    "t0": (u"Überschreibungen dieses Filters vorher zurücksetzen",
           u"Reset this filter's overrides first",
           u"Restablecer antes las modificaciones de este filtro"),
    "t1": (u"Abbrechen", u"Cancel", u"Cancelar"),
    "t2": (u"Leere Felder bleiben unverändert. × entfernt eine "
           u"Überschreibung.",
           u"Empty fields stay unchanged. × removes an override.",
           u"Los campos vacíos no cambian. × quita una modificación."),
}

_XAML = u"""
<Window %s Title="{titel}" Width="600" SizeToContent="Height"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12">
  <StackPanel Margin="14">
    <TextBlock x:Name="hinweis" TextWrapping="Wrap" Margin="0,0,0,10"/>
    <Border BorderBrush="#ABADB3" BorderThickness="1" Padding="8">
      <Grid x:Name="raster">
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
      </Grid>
    </Border>
    <TextBlock Text="{{t2}}" Foreground="#666" TextWrapping="Wrap"
               Margin="0,6,0,0"/>
    <CheckBox x:Name="zuruecksetzen" Content="{{t0}}" Margin="0,10,0,0"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"
                Margin="0,14,0,0">
      <Button x:Name="ok" Content="OK" Width="90" Margin="0,0,8,0"/>
      <Button Content="{{t1}}" Width="90" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>""" % dlg.XMLNS


def _pinsel(rgb):
    from System.Windows.Media import Color as FarbeWpf
    pinsel = SolidColorBrush(FarbeWpf.FromRgb(rgb[0], rgb[1], rgb[2]))
    pinsel.Freeze()
    return pinsel


def frage_darstellung(besitzer, doc, titel, hinweis, vorbelegung=None,
                      sichtbarkeit_vorgabe=True):
    """Dialog "Sichtbarkeit und Grafiken". Rückgabe Darstellung oder None.

    vorbelegung   OverrideGraphicSettings, deren Werte angezeigt werden
                  (nur zur Orientierung - geschrieben wird nur Geändertes)
    sichtbarkeit_vorgabe  True/False/None für die Vorauswahl
    """
    fenster = dlg.lade_xaml(uebersetze_xaml(
        _XAML.replace(u"{titel}", dlg._escape(titel)), _TEXTE))
    dlg.setze_besitzer(fenster, besitzer=besitzer)
    fenster.FindName("hinweis").Text = hinweis
    raster = fenster.FindName("raster")
    darstellung = Darstellung()
    muster_liste = fuellmuster(doc)
    zeile = [0]

    def neue_zeile():
        raster.RowDefinitions.Add(RowDefinition())
        zeile[0] += 1
        return zeile[0] - 1

    def setze(element, reihe, spalte, spannweite=1):
        Grid.SetRow(element, reihe)
        Grid.SetColumn(element, spalte)
        if spannweite > 1:
            Grid.SetColumnSpan(element, spannweite)
        raster.Children.Add(element)

    def beschriftung(text, reihe, fett=False, einzug=0.0):
        block = TextBlock()
        block.Text = text
        block.VerticalAlignment = VerticalAlignment.Center
        block.Margin = Thickness(einzug, 4.0, 12.0, 4.0)
        if fett:
            block.FontWeight = FontWeights.Bold
            block.Margin = Thickness(0.0, 8.0, 0.0, 2.0)
        setze(block, reihe, 0, 3 if fett else 1)

    def auswahl(schluessel, eintraege, aktuell=None):
        """ComboBox: [(Text, Wert)] - erster Eintrag leer = unverändert."""
        box = ComboBox()
        box.Margin = Thickness(0.0, 2.0, 6.0, 2.0)
        leer = ComboBoxItem()
        leer.Content = u""
        leer.Tag = UNVERAENDERT
        box.Items.Add(leer)
        keine = ComboBoxItem()
        keine.Content = t(u"× keine Überschreibung",
                          u"× no override", u"× sin modificación")
        keine.Tag = KEINE
        box.Items.Add(keine)
        for text, wert in eintraege:
            item = ComboBoxItem()
            item.Content = text
            item.Tag = wert
            box.Items.Add(item)
        box.SelectedIndex = 0
        if aktuell is not None:
            box.ToolTip = t(u"Aktuell: %s", u"Current: %s",
                            u"Actual: %s") % aktuell

        def geaendert(sender, args):
            if sender.SelectedItem is not None:
                darstellung.werte[schluessel] = sender.SelectedItem.Tag
        box.SelectionChanged += dlg.sicher(lambda: fenster, geaendert)
        return box

    def farbfeld(schluessel, aktuell=None):
        leiste = StackPanel()
        leiste.Orientation = Orientation.Horizontal
        leiste.Margin = Thickness(0.0, 2.0, 6.0, 2.0)
        knopf = Button()
        knopf.MinWidth = 120.0
        knopf.Padding = Thickness(4.0, 1.0, 4.0, 1.0)
        inhalt = StackPanel()
        inhalt.Orientation = Orientation.Horizontal
        feld = Border()
        feld.Width = 26.0
        feld.Height = 12.0
        feld.BorderThickness = Thickness(1.0)
        feld.BorderBrush = _pinsel((110, 110, 110))
        feld.Margin = Thickness(0.0, 0.0, 6.0, 0.0)
        text = TextBlock()
        inhalt.Children.Add(feld)
        inhalt.Children.Add(text)
        knopf.Content = inhalt
        weg = Button()
        weg.Content = u"×"
        weg.Width = 22.0
        weg.Margin = Thickness(2.0, 0.0, 0.0, 0.0)
        weg.ToolTip = t(u"Keine Überschreibung", u"No override",
                        u"Sin modificación")

        def zeige():
            wert = darstellung.werte[schluessel]
            if wert == UNVERAENDERT:
                # Aktuelle Überschreibung blass anzeigen (nur zur Orientierung)
                feld.Visibility = (Visibility.Visible if aktuell
                                   else Visibility.Collapsed)
                if aktuell:
                    feld.Background = _pinsel(aktuell)
                feld.Opacity = 0.45
                text.Text = u"%d-%d-%d" % aktuell if aktuell else u""
                text.Foreground = _pinsel((140, 140, 140))
                knopf.ToolTip = (t(u"Aktuell RGB %d-%d-%d - unverändert",
                                   u"Current RGB %d-%d-%d - unchanged",
                                   u"Actual RGB %d-%d-%d - sin cambios")
                                 % aktuell) if aktuell else None
            elif wert == KEINE:
                feld.Visibility = Visibility.Collapsed
                text.Text = t(u"× keine Überschreibung",
                              u"× no override", u"× sin modificación")
                text.Foreground = _pinsel((192, 57, 43))
            else:
                feld.Visibility = Visibility.Visible
                feld.Opacity = 1.0
                feld.Background = _pinsel(wert)
                text.Text = u"%d-%d-%d" % wert
                text.Foreground = _pinsel((0, 0, 0))

        def waehlen(sender, args):
            vorher = darstellung.werte[schluessel]
            neu = farbe_waehlen(fenster, vorher if isinstance(vorher, tuple)
                                else aktuell)
            if neu is not None:
                darstellung.werte[schluessel] = neu
                zeige()

        def entfernen(sender, args):
            darstellung.werte[schluessel] = KEINE
            zeige()

        knopf.Click += dlg.sicher(lambda: fenster, waehlen)
        weg.Click += dlg.sicher(lambda: fenster, entfernen)
        leiste.Children.Add(knopf)
        leiste.Children.Add(weg)
        zeige()
        return leiste

    # Aktuelle Werte (nur Anzeige)
    ogs = vorbelegung

    def aktuelle_farbe(name):
        return _farbe_aus(getattr(ogs, name, None)) if ogs is not None else None

    def aktueller_wert(name, leer=-1):
        if ogs is None:
            return None
        try:
            wert = getattr(ogs, name)
            wert = id_wert(wert) if hasattr(wert, "Value") else wert
            return None if wert in (leer, None) else wert
        except Exception:
            return None

    muster_namen = dict((w, n) for n, w in muster_liste)

    def muster_text(name):
        wert = aktueller_wert(name)
        return muster_namen.get(wert) if wert is not None else None

    staerken = [(u"%d" % n, n) for n in range(1, 17)]
    transparenzen = [(u"%d %%" % n, n) for n in range(0, 101, 10)]

    # Sichtbarkeit
    reihe = neue_zeile()
    beschriftung(t(u"Sichtbarkeit", u"Visibility", u"Visibilidad"), reihe)
    sicht = ComboBox()
    sicht.Margin = Thickness(0.0, 2.0, 6.0, 2.0)
    for text, wert in ((t(u"Sichtbar", u"Visible", u"Visible"), True),
                       (t(u"Unsichtbar", u"Hidden", u"Oculto"), False),
                       (t(u"nicht ändern", u"do not change", u"no cambiar"),
                        UNVERAENDERT)):
        item = ComboBoxItem()
        item.Content = text
        item.Tag = wert
        sicht.Items.Add(item)
    sicht.SelectedIndex = {True: 0, False: 1}.get(sichtbarkeit_vorgabe, 2)
    darstellung.werte[SICHTBAR] = sicht.SelectedItem.Tag

    def bei_sicht(sender, args):
        darstellung.werte[SICHTBAR] = sender.SelectedItem.Tag
    sicht.SelectionChanged += dlg.sicher(lambda: fenster, bei_sicht)
    setze(sicht, reihe, 1)

    for kopf, linie_f, linie_s, muster_f, muster_m, praefix in (
            (t(u"Projektion/Oberfläche", u"Projection/Surface",
               u"Proyección/Superficie"),
             PROJ_LINIE_FARBE, PROJ_LINIE_STAERKE, OBERFL_MUSTER_FARBE,
             OBERFL_MUSTER, ("ProjectionLine", "SurfaceForegroundPattern")),
            (t(u"Schnitt", u"Cut", u"Corte"),
             SCHNITT_LINIE_FARBE, SCHNITT_LINIE_STAERKE, SCHNITT_MUSTER_FARBE,
             SCHNITT_MUSTER, ("CutLine", "CutForegroundPattern"))):
        beschriftung(kopf, neue_zeile(), fett=True)

        reihe = neue_zeile()
        beschriftung(t(u"Linien: Farbe / Stärke", u"Lines: color / weight",
                       u"Líneas: color / grosor"), reihe, einzug=10.0)
        setze(farbfeld(linie_f, aktuelle_farbe(praefix[0] + "Color")),
              reihe, 1)
        setze(auswahl(linie_s, staerken,
                      aktueller_wert(praefix[0] + "Weight")), reihe, 2)

        reihe = neue_zeile()
        beschriftung(t(u"Muster: Farbe / Muster", u"Pattern: color / pattern",
                       u"Trama: color / trama"), reihe, einzug=10.0)
        setze(farbfeld(muster_f, aktuelle_farbe(praefix[1] + "Color")),
              reihe, 1)
        setze(auswahl(muster_m, muster_liste, muster_text(praefix[1] + "Id")),
              reihe, 2)

        if linie_f == PROJ_LINIE_FARBE:
            reihe = neue_zeile()
            beschriftung(t(u"Transparenz", u"Transparency", u"Transparencia"),
                         reihe, einzug=10.0)
            aktuell = aktueller_wert("Transparency", leer=0)
            setze(auswahl(TRANSPARENZ, transparenzen,
                          u"%s %%" % aktuell if aktuell else None), reihe, 1)

    reihe = neue_zeile()
    beschriftung(t(u"Halbton", u"Halftone", u"Medio tono"), reihe)
    halbton = auswahl(HALBTON, [(t(u"Ja", u"Yes", u"Sí"), True),
                                (t(u"Nein", u"No", u"No"), False)])
    halbton.Items.RemoveAt(1)          # "keine Überschreibung" = Nein
    setze(halbton, reihe, 1)

    ergebnis = {"ok": False}

    def bei_ok(sender, args):
        darstellung.zuruecksetzen = bool(
            fenster.FindName("zuruecksetzen").IsChecked)
        ergebnis["ok"] = True
        fenster.DialogResult = True

    fenster.FindName("ok").Click += dlg.sicher(lambda: fenster, bei_ok)
    fenster.ShowDialog()
    return darstellung if ergebnis["ok"] else None
