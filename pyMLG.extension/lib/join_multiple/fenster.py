# -*- coding: utf-8 -*-
"""Hauptfenster von JoinMultiple (WPF), angelehnt an "JoinMultiple".

    ┌ Kategorie ─────────── Elemente ─ Priorität ┐ Alle  Keine
    │ ☑ Wände                    12     [300]    │ Bereich: ● Auswahl ○ Ansicht ○ Modell
    │ ☑ Geschossdecken            4     [200]    │ Verbinden: ☑ nur bei Schnitt ...
    │                                            │ Priorität: □ aus Parameter [____]
    │                                            │ 16 Elemente in der Auswahl
    └────────────────────────────────────────────┘ [Verbindung lösen] [Verbinden]

Höhere Priorität schneidet niedrigere. Prioritäten und Haken werden je
Kategorie in %LOCALAPPDATA%\\pyMLG\\JoinMultiple.json gemerkt.

Das Fenster ist modal: Die Revit-Auswahl wird beim Öffnen übernommen.
"""

import io
import json
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import GridLength, GridUnitType, HorizontalAlignment, \
    Thickness, VerticalAlignment  # noqa: E402
from System.Windows.Controls import CheckBox, ColumnDefinition, Control, \
    Grid, ListBoxItem, TextBlock, TextBox  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from join_multiple import logik as lg  # noqa: E402
from join_multiple import revit as rv  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402

TITEL = u"JoinMultiple"


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


ROT = _pinsel(192, 57, 43)

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(),
                               "JoinMultiple_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "JoinMultiple.json")

BEREICH_TEXT = {
    rv.AUSWAHL: t(u"Elemente in der Auswahl", u"elements in selection", u"elementos en la selección"),
    rv.ANSICHT: t(u"Elemente in der aktuellen Ansicht", u"elements in active view", u"elementos en la vista activa"),
    rv.MODELL: t(u"Elemente im ganzen Modell", u"elements in the whole model", u"elementos en todo el modelo"),
}

BREITE_ANZAHL = 80.0
BREITE_PRIO = 90.0

XAML_TEXTE = {
    "t0": (u"Höhere Priorität schneidet niedrigere. Nur markierte Kategorien werden bearbeitet.",
           u"Higher priority cuts lower. Only checked categories are processed.",
           u"La prioridad mayor corta a la menor. Solo se procesan las categorías marcadas."),
    "t1": (u"Kategorie",
           u"Category",
           u"Categoría"),
    "t2": (u"Elemente",
           u"Elements",
           u"Elementos"),
    "t3": (u"Priorität",
           u"Priority",
           u"Prioridad"),
    "t4": (u"Alle markieren",
           u"Check all",
           u"Marcar todo"),
    "t5": (u"Keine",
           u"None",
           u"Ninguno"),
    "t6": (u"Bereich",
           u"Scope",
           u"Ámbito"),
    "t7": (u"Ausgewählte Elemente",
           u"Selected elements",
           u"Elementos seleccionados"),
    "t8": (u"Elemente in aktueller Ansicht",
           u"Elements in active view",
           u"Elementos en la vista activa"),
    "t9": (u"Alle Modellelemente",
           u"All model elements",
           u"Todos los elementos del modelo"),
    "t10": (u"Sonst nur Kategorien, die sich mit 'Geometrie verbinden' zuverlässig bearbeiten lassen (Wände, Decken, Dächer, Stützen, Tragwerk, Fundamente, Allgemeines Modell ...)",
           u"Otherwise only categories that work reliably with 'Join Geometry' (walls, floors, roofs, columns, framing, foundations, generic models ...)",
           u"Si no, solo categorías que funcionan bien con 'Unir geometría' (muros, suelos, cubiertas, pilares, armazón, cimentaciones, modelos genéricos ...)"),
    "t11": (u"Alle Modellkategorien anzeigen",
           u"Show all model categories",
           u"Mostrar todas las categorías de modelo"),
    "t12": (u"Verbinden",
           u"Join",
           u"Unir"),
    "t13": (u"Nur Elemente verbinden, deren Volumen sich überschneiden. Aus: auch Elemente, die sich nur berühren oder deren Umrisse sich überlagern.",
           u"Only join elements whose volumes overlap. Off: also elements that only touch or whose bounding boxes overlap.",
           u"Unir solo elementos cuyos volúmenes se solapan. Desactivado: también elementos que solo se tocan o cuyos contornos se superponen."),
    "t14": (u"Nur wenn sich Elemente schneiden",
           u"Only if elements intersect",
           u"Solo si los elementos se intersecan"),
    "t15": (u"Auch innerhalb derselben Kategorie",
           u"Also within the same category",
           u"También dentro de la misma categoría"),
    "t16": (u"Bereits verbundene Elemente: Schnittreihenfolge nach Priorität umkehren, falls nötig",
           u"Already joined elements: switch the cut order by priority if needed",
           u"Elementos ya unidos: invertir el orden de corte según la prioridad si es necesario"),
    "t17": (u"Bereits verbundene nach Priorität anpassen",
           u"Adjust already joined by priority",
           u"Ajustar los ya unidos por prioridad"),
    "t18": (u"Zahlenwert eines Exemplar- oder Typparameters. Elemente ohne Wert nutzen die Priorität ihrer Kategorie.",
           u"Numeric value of an instance or type parameter. Elements without a value use the priority of their category.",
           u"Valor numérico de un parámetro de ejemplar o de tipo. Los elementos sin valor usan la prioridad de su categoría."),
    "t19": (u"Priorität aus Parameter:",
           u"Priority from parameter:",
           u"Prioridad desde parámetro:"),
    "t20": (u"Prioritäten zurücksetzen",
           u"Reset priorities",
           u"Restablecer prioridades"),
    "t21": (u"Alle Kategorien auf die Vorgabewerte setzen (Stützen 500 ... Wände 300, Decken 200 ...)",
           u"Set all categories to the default values (columns 500 ... walls 300, floors 200 ...)",
           u"Poner todas las categorías en los valores por defecto (pilares 500 ... muros 300, suelos 200 ...)"),
    "t22": (u"Bearbeitete Elemente danach auswählen",
           u"Select processed elements afterwards",
           u"Seleccionar después los elementos procesados"),
    "t23": (u"Verbindung lösen",
           u"Unjoin elements",
           u"Desunir elementos"),
    "t24": (u"Elemente verbinden",
           u"Join elements",
           u"Unir elementos"),
    "t25": (u"Schließen",
           u"Close",
           u"Cerrar"),
}

XAML = u"""
<Window %s Title="JoinMultiple (pyMLG)" Width="980" Height="640"
        MinWidth="760" MinHeight="520" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,3"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Margin" Value="0,3"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="270"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Kategorien -->
    <DockPanel>
      <TextBlock DockPanel.Dock="Top" Foreground="#555" Margin="0,0,0,6"
                 TextWrapping="Wrap"
                 Text="{{t0}}"/>
      <Border DockPanel.Dock="Top" BorderBrush="#ABADB3"
              BorderThickness="1,1,1,0" Background="#F3F3F3">
        <Grid Margin="6,4,24,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="80"/>
            <ColumnDefinition Width="90"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="{{t1}}" FontWeight="SemiBold"/>
          <TextBlock Grid.Column="1" Text="{{t2}}" FontWeight="SemiBold"
                     HorizontalAlignment="Right" Margin="0,0,12,0"/>
          <TextBlock Grid.Column="2" Text="{{t3}}" FontWeight="SemiBold"/>
        </Grid>
      </Border>
      <ListBox x:Name="liste" HorizontalContentAlignment="Stretch"
               ScrollViewer.VerticalScrollBarVisibility="Visible"/>
    </DockPanel>

    <!-- Optionen -->
    <DockPanel Grid.Column="2" LastChildFill="False">
      <StackPanel DockPanel.Dock="Top">
        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
          <Button x:Name="alle" Content="{{t4}}" Padding="8,3"
                  Margin="0,0,6,0"/>
          <Button x:Name="keine" Content="{{t5}}" Padding="8,3"/>
        </StackPanel>

        <GroupBox Header="{{t6}}">
          <StackPanel>
            <RadioButton x:Name="bereich_auswahl"
                         Content="{{t7}}"/>
            <RadioButton x:Name="bereich_ansicht"
                         Content="{{t8}}"/>
            <RadioButton x:Name="bereich_modell" Content="{{t9}}"/>
            <CheckBox x:Name="alle_kategorien"
                      ToolTip="{{t10}}">
              <TextBlock Text="{{t11}}"
                         TextWrapping="Wrap"/>
            </CheckBox>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="{{t12}}">
          <StackPanel>
            <CheckBox x:Name="nur_bei_schnitt"
                      ToolTip="{{t13}}">
              <TextBlock Text="{{t14}}"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="gleiche_kategorie">
              <TextBlock Text="{{t15}}"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="bestehende_anpassen"
                      ToolTip="{{t16}}">
              <TextBlock Text="{{t17}}"
                         TextWrapping="Wrap"/>
            </CheckBox>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="{{t3}}">
          <StackPanel>
            <CheckBox x:Name="prio_parameter"
                      ToolTip="{{t18}}">
              <TextBlock Text="{{t19}}"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <TextBox x:Name="parametername" Padding="3" Margin="18,2,0,6"/>
            <Button x:Name="zuruecksetzen" Padding="8,4"
                    Content="{{t20}}"
                    ToolTip="{{t21}}"/>
          </StackPanel>
        </GroupBox>

        <CheckBox x:Name="auswaehlen" Margin="2,0,0,0">
          <TextBlock Text="{{t22}}"
                     TextWrapping="Wrap"/>
        </CheckBox>
      </StackPanel>

      <StackPanel DockPanel.Dock="Bottom">
        <TextBlock x:Name="anzahl" FontSize="22" FontWeight="SemiBold"/>
        <TextBlock x:Name="anzahl_text" Foreground="#555" Margin="0,0,0,10"
                   TextWrapping="Wrap"/>
        <Button x:Name="loesen" Content="{{t23}}" Padding="8,6"
                Margin="0,0,0,6"/>
        <Button x:Name="verbinden" Content="{{t24}}"
                FontWeight="Bold" Padding="8,9"/>
      </StackPanel>
    </DockPanel>

    <!-- Fußzeile -->
    <DockPanel Grid.Row="1" Grid.ColumnSpan="3" Margin="0,10,0,0">
      <Button x:Name="schliessen" Content="{{t25}}" Width="100"
              DockPanel.Dock="Right" IsCancel="True"/>
      <TextBlock x:Name="status" VerticalAlignment="Center" Foreground="#555"
                 TextTrimming="CharacterEllipsis"/>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


# ---------------------------------------------------------------------------
# Hilfen
# ---------------------------------------------------------------------------

def schreibe_fehlerprotokoll(spur):
    try:
        with io.open(FEHLERPROTOKOLL, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
        return FEHLERPROTOKOLL
    except Exception:
        return None


def meldung(besitzer, text, warnung=False):
    dlg.meldung(besitzer, text, titel=TITEL, warnung=warnung)


def sicher(besitzer_liefern, funktion):
    """Ereignishandler, der Ausnahmen abfängt und verständlich anzeigt -
    eine Ausnahme im Handler würde sonst ShowDialog() beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except Exception as fehler:
            pfad = schreibe_fehlerprotokoll(traceback.format_exc())
            try:
                meldung(besitzer_liefern(), u"%s%s" % (
                    dlg.fehlertext(fehler),
                    t(u"\n\nTechnische Details: %s", u"\n\nTechnical details: %s", u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
                    warnung=True)
            except Exception:
                pass
    return handler


def lade_einstellungen():
    try:
        with io.open(EINSTELLUNGEN, "r", encoding="utf-8") as datei:
            return json.load(datei)
    except Exception:
        return {}


def speichere_einstellungen(werte):
    try:
        with io.open(EINSTELLUNGEN, "w", encoding="utf-8") as datei:
            datei.write(json.dumps(werte, ensure_ascii=False, indent=2))
    except Exception:
        pass


def _spalte(breite=None):
    spalte = ColumnDefinition()
    spalte.Width = (GridLength(breite) if breite
                    else GridLength(1.0, GridUnitType.Star))
    return spalte


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------

class Kategorie(object):
    def __init__(self, schluessel, name):
        self.schluessel = schluessel      # Id-Wert als Text (JSON-Schlüssel)
        self.name = name
        self.elemente = []
        self.box = None
        self.prio_feld = None


class JoinMultipleFenster(object):

    def __init__(self, uiapp, uidoc):
        self.uiapp = uiapp
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.kategorien = []           # [Kategorie], sortiert nach Name
        self.vorgaben = dict((str(k), lg.standard_prioritaet(n))
                             for k, (_b, n)
                             in rv.standard_kategorien().items())

        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("liste", "bereich_auswahl", "bereich_ansicht",
                     "bereich_modell", "alle_kategorien", "nur_bei_schnitt",
                     "gleiche_kategorie", "bestehende_anpassen",
                     "prio_parameter", "parametername", "auswaehlen",
                     "anzahl", "anzahl_text", "status"):
            setattr(self, name, f.FindName(name))

        einst = self.einst = lade_einstellungen()
        self.prioritaeten = dict(einst.get("prioritaeten", {}))
        self.abgewaehlt = set(einst.get("abgewaehlt", []))
        self.alle_kategorien.IsChecked = einst.get("alle_kategorien", False)
        self.nur_bei_schnitt.IsChecked = einst.get("nur_bei_schnitt", True)
        self.gleiche_kategorie.IsChecked = einst.get("gleiche_kategorie",
                                                     True)
        self.bestehende_anpassen.IsChecked = einst.get(
            "bestehende_anpassen", True)
        self.prio_parameter.IsChecked = einst.get("prio_parameter", False)
        self.parametername.Text = einst.get("parametername", u"")
        self.auswaehlen.IsChecked = einst.get("auswaehlen", False)

        # Mit Auswahl startet der Bereich "Auswahl", sonst die Ansicht
        hat_auswahl = self.uidoc.Selection.GetElementIds().Count > 0
        self.bereich_auswahl.IsEnabled = hat_auswahl
        if hat_auswahl:
            self.bereich_auswahl.IsChecked = True
        elif einst.get("bereich") == rv.MODELL:
            self.bereich_modell.IsChecked = True
        else:
            self.bereich_ansicht.IsChecked = True
        if not hat_auswahl:
            self.bereich_auswahl.Content = t(u"Ausgewählte Elemente (keine)", u"Selected elements (none)", u"Elementos seleccionados (ninguno)")

        self._verdrahte()
        self.lade()

    # --- Verdrahtung ------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        def klick(name, funktion):
            f.FindName(name).Click += s(lambda _s, _a: funktion())

        klick("alle", lambda: self.markiere(True))
        klick("keine", lambda: self.markiere(False))
        klick("zuruecksetzen", self.setze_zurueck)
        klick(t("verbinden", u"join", u"unir"), self.verbinde)
        klick("loesen", self.loese)
        klick("schliessen", f.Close)
        for box in (self.bereich_auswahl, self.bereich_ansicht,
                    self.bereich_modell):
            box.Checked += s(lambda _s, _a: self.lade())
        self.alle_kategorien.Checked += s(lambda _s, _a: self.lade())
        self.alle_kategorien.Unchecked += s(lambda _s, _a: self.lade())
        f.Closing += s(lambda _s, _a: self._speichere())

    def _speichere(self):
        self._uebernimm_felder()
        speichere_einstellungen({
            "prioritaeten": self.prioritaeten,
            "abgewaehlt": sorted(self.abgewaehlt),
            "bereich": self.bereich(),
            "alle_kategorien": bool(self.alle_kategorien.IsChecked),
            "nur_bei_schnitt": bool(self.nur_bei_schnitt.IsChecked),
            "gleiche_kategorie": bool(self.gleiche_kategorie.IsChecked),
            "bestehende_anpassen": bool(self.bestehende_anpassen.IsChecked),
            "prio_parameter": bool(self.prio_parameter.IsChecked),
            "parametername": self.parametername.Text or u"",
            "auswaehlen": bool(self.auswaehlen.IsChecked),
        })

    # --- Liste ------------------------------------------------------------
    def bereich(self):
        if self.bereich_auswahl.IsChecked:
            return rv.AUSWAHL
        if self.bereich_modell.IsChecked:
            return rv.MODELL
        return rv.ANSICHT

    def prioritaet_text(self, schluessel):
        if schluessel in self.prioritaeten:
            return u"%s" % self.prioritaeten[schluessel]
        return u"%s" % self.vorgaben.get(schluessel, 0)

    def lade(self):
        """Elemente des Bereichs sammeln und nach Kategorie auflisten."""
        self._uebernimm_felder()
        elemente = rv.sammle(self.doc, self.uidoc, self.bereich(),
                             bool(self.alle_kategorien.IsChecked))
        nach_schluessel = {}
        for element in elemente:
            schluessel = str(rv.kategorie_schluessel(element))
            kategorie = nach_schluessel.get(schluessel)
            if kategorie is None:
                kategorie = nach_schluessel[schluessel] = Kategorie(
                    schluessel, element.Category.Name)
            kategorie.elemente.append(element)
        self.kategorien = sorted(nach_schluessel.values(),
                                 key=lambda k: k.name.casefold())
        self.zeichne()

    def zeichne(self):
        self.liste.Items.Clear()
        for kategorie in self.kategorien:
            zeile = Grid()
            for breite in (None, BREITE_ANZAHL, BREITE_PRIO):
                zeile.ColumnDefinitions.Add(_spalte(breite))

            box = CheckBox()
            box.Content = kategorie.name
            box.IsChecked = kategorie.schluessel not in self.abgewaehlt
            box.VerticalAlignment = VerticalAlignment.Center
            box.Click += sicher(lambda: self.fenster,
                                self._haken_handler(kategorie))
            zeile.Children.Add(box)

            anzahl = TextBlock()
            anzahl.Text = u"%d" % len(kategorie.elemente)
            anzahl.HorizontalAlignment = HorizontalAlignment.Right
            anzahl.VerticalAlignment = VerticalAlignment.Center
            anzahl.Margin = Thickness(0.0, 0.0, 12.0, 0.0)
            Grid.SetColumn(anzahl, 1)
            zeile.Children.Add(anzahl)

            feld = TextBox()
            feld.Text = self.prioritaet_text(kategorie.schluessel)
            feld.Width = 70.0
            feld.Padding = Thickness(2.0, 1.0, 2.0, 1.0)
            feld.HorizontalAlignment = HorizontalAlignment.Left
            feld.TextChanged += sicher(lambda: self.fenster,
                                       self._prio_handler(kategorie))
            Grid.SetColumn(feld, 2)
            zeile.Children.Add(feld)

            item = ListBoxItem()
            item.Content = zeile
            item.Padding = Thickness(4.0, 2.0, 4.0, 2.0)
            kategorie.box, kategorie.prio_feld = box, feld
            self.liste.Items.Add(item)
        self.aktualisiere_anzahl()

    def _haken_handler(self, kategorie):
        def handler(sender, _args):
            if sender.IsChecked:
                self.abgewaehlt.discard(kategorie.schluessel)
            else:
                self.abgewaehlt.add(kategorie.schluessel)
            self.aktualisiere_anzahl()
        return handler

    def _prio_handler(self, kategorie):
        def handler(sender, _args):
            try:
                lg.lies_prioritaet(sender.Text)
            except ValueError as fehler:
                sender.Foreground = ROT
                sender.ToolTip = u"%s" % fehler
                return
            sender.ClearValue(Control.ForegroundProperty)
            sender.ToolTip = None
        return handler

    def _uebernimm_felder(self):
        """Gültige Prioritäten aus den Eingabefeldern merken."""
        for kategorie in self.kategorien:
            if kategorie.prio_feld is None:
                continue
            try:
                wert = lg.lies_prioritaet(kategorie.prio_feld.Text)
            except ValueError:
                continue
            if wert == self.vorgaben.get(kategorie.schluessel, 0):
                self.prioritaeten.pop(kategorie.schluessel, None)
            else:
                self.prioritaeten[kategorie.schluessel] = wert

    def markierte(self):
        return [k for k in self.kategorien
                if k.schluessel not in self.abgewaehlt]

    def aktualisiere_anzahl(self):
        elemente = sum(len(k.elemente) for k in self.markierte())
        self.anzahl.Text = u"%d" % elemente
        self.anzahl_text.Text = BEREICH_TEXT[self.bereich()] + (
            u"" if len(self.markierte()) == len(self.kategorien)
            else t(u" (markierte Kategorien)", u" (checked categories)", u" (categorías marcadas)"))

    def markiere(self, zustand):
        for kategorie in self.kategorien:
            kategorie.box.IsChecked = zustand
            if zustand:
                self.abgewaehlt.discard(kategorie.schluessel)
            else:
                self.abgewaehlt.add(kategorie.schluessel)
        self.aktualisiere_anzahl()

    def setze_zurueck(self):
        if not dlg.frage(self.fenster, t(u"Die Prioritäten aller Kategorien "
                         u"auf die Vorgabewerte zurücksetzen?", u"Reset the priorities of all categories to the default values?", u"¿Restablecer las prioridades de todas las categorías a los valores por defecto?"), titel=TITEL):
            return
        self.prioritaeten = {}
        for kategorie in self.kategorien:
            kategorie.prio_feld.Text = u"%s" % self.vorgaben.get(
                kategorie.schluessel, 0)
        self.status.Text = t(u"Prioritäten zurückgesetzt.", u"Priorities reset.", u"Prioridades restablecidas.")

    # --- Ausführen --------------------------------------------------------
    def _vorbereiten(self, aktion):
        """Markierte Elemente und Prioritätsfunktion oder None."""
        kategorien = self.markierte()
        elemente = [e for k in kategorien for e in k.elemente]
        if not elemente:
            meldung(self.fenster, t(u"Keine Elemente in markierten "
                    u"Kategorien.", u"No elements in checked categories.", u"No hay elementos en las categorías marcadas."))
            return None
        prios = {}
        for kategorie in kategorien:
            try:
                prios[kategorie.schluessel] = lg.lies_prioritaet(
                    kategorie.prio_feld.Text)
            except ValueError as fehler:
                meldung(self.fenster, u"%s: %s" % (kategorie.name, fehler),
                        warnung=True)
                return None
        self._uebernimm_felder()

        parameter = None
        if self.prio_parameter.IsChecked:
            parameter = (self.parametername.Text or u"").strip()
            if not parameter:
                meldung(self.fenster, t(u"Bitte den Namen des Parameters für "
                        u"die Priorität eingeben.", u"Please enter the name of the priority parameter.", u"Introduzca el nombre del parámetro de prioridad."), warnung=True)
                return None

        def prioritaet(element):
            if parameter:
                wert = rv.parameter_zahl(self.doc, element, parameter)
                if wert is not None:
                    return wert
            return prios[str(rv.kategorie_schluessel(element))]

        if self.bereich() == rv.MODELL and len(elemente) > 2000:
            if not dlg.frage(self.fenster, t(u"%d Elemente im ganzen Modell "
                             u"%s? Das kann einige Minuten dauern.", u"%d elements in the whole model - %s? This can take a few minutes.", u"%d elementos en todo el modelo - ¿%s? Puede tardar unos minutos.")
                             % (len(elemente), aktion), titel=TITEL):
                return None
        return elemente, prioritaet

    def verbinde(self):
        vorbereitet = self._vorbereiten(t(u"verbinden", u"join", u"unir"))
        if vorbereitet is None:
            return
        elemente, prioritaet = vorbereitet
        ergebnis = rv.verbinde(
            self.doc, elemente, prioritaet,
            nur_bei_schnitt=bool(self.nur_bei_schnitt.IsChecked),
            gleiche_kategorie=bool(self.gleiche_kategorie.IsChecked),
            bestehende_anpassen=bool(self.bestehende_anpassen.IsChecked))
        n = ergebnis.anzahl
        zeilen = [t(u"Neu verbunden: %d", u"Newly joined: %d", u"Unidos nuevos: %d") % n[lg.NEU],
                  t(u"Bereits verbunden: %d", u"Already joined: %d", u"Ya unidos: %d") % n[lg.BEREITS]]
        if self.nur_bei_schnitt.IsChecked:
            zeilen.append(t(u"Übersprungen, schneiden sich nicht: %d", u"Skipped, do not intersect: %d", u"Omitidos, no se intersecan: %d")
                          % n[lg.UEBERSPRUNGEN])
        zeilen += [u"",
                   t(u"Schnittreihenfolge nach Priorität:", u"Cut order by priority:", u"Orden de corte por prioridad:"),
                   t(u"  umgekehrt: %d", u"  switched: %d", u"  invertidos: %d") % n[lg.UMGEDREHT],
                   t(u"  stimmte bereits: %d", u"  already correct: %d", u"  ya correctos: %d") % n[lg.RICHTIG]]
        if n[lg.GLEICH]:
            zeilen.append(t(u"  gleiche Priorität, unverändert: %d", u"  same priority, unchanged: %d", u"  misma prioridad, sin cambios: %d")
                          % n[lg.GLEICH])
        if n[lg.NICHT_GEPRUEFT]:
            zeilen.append(
                t(u"  NICHT geprüft: %d bereits verbundene Paare - dafür "
                u"'Bereits verbundene nach Priorität anpassen' einschalten", u"  NOT checked: %d already joined pairs - enable 'Adjust already joined by priority'", u"  NO comprobados: %d pares ya unidos - active 'Ajustar los ya unidos por prioridad'")
                % n[lg.NICHT_GEPRUEFT])
        if n[lg.ABGELEHNT]:
            zeilen.append(t(u"  von Revit nicht übernommen: %d", u"  not accepted by Revit: %d", u"  no aceptados por Revit: %d")
                          % n[lg.ABGELEHNT])
            zeilen.extend(u"    " + text for text in ergebnis.abgelehnt[:10])
        if ergebnis.falsch_nach_speichern:
            zeilen.append(t(u"  nach dem Speichern FALSCH: %d", u"  WRONG after saving: %d", u"  INCORRECTOS después de guardar: %d")
                          % len(ergebnis.falsch_nach_speichern))
            zeilen.extend(u"    " + text
                          for text in ergebnis.falsch_nach_speichern[:10])
        else:
            zeilen.append(t(u"  nach dem Speichern kontrolliert: alles richtig", u"  checked after saving: all correct", u"  comprobado después de guardar: todo correcto"))
        zeilen += [u"", t(u"Protokoll: %s", u"Log: %s", u"Registro: %s") % rv.PROTOKOLL]
        self._abschluss(elemente, ergebnis, zeilen,
                        t(u"%d verbunden, %d umgekehrt", u"%d joined, %d switched", u"%d unidos, %d invertidos") % (
                            n[lg.NEU], n[lg.UMGEDREHT]))

    def loese(self):
        vorbereitet = self._vorbereiten(t(u"lösen", u"unjoin", u"desunir"))
        if vorbereitet is None:
            return
        elemente, _prioritaet = vorbereitet
        ergebnis = rv.loese(
            self.doc, elemente,
            gleiche_kategorie=bool(self.gleiche_kategorie.IsChecked))
        self._abschluss(elemente, ergebnis,
                        [t(u"Verbindungen gelöst: %d", u"Joins removed: %d", u"Uniones eliminadas: %d") % ergebnis.geloest],
                        t(u"%d Verbindungen gelöst", u"%d joins removed", u"%d uniones eliminadas") % ergebnis.geloest)

    def _abschluss(self, elemente, ergebnis, zeilen, kurz, maximal=10):
        if self.auswaehlen.IsChecked and ergebnis.bearbeitet:
            rv.waehle_aus(self.uidoc, elemente, ergebnis.bearbeitet)
        if ergebnis.fehler:
            zeilen.append(t(u"\nFehlgeschlagen: %d", u"\nFailed: %d", u"\nFallidos: %d") % len(ergebnis.fehler))
            zeilen.extend(u"  " + text for text in ergebnis.fehler[:maximal])
            if len(ergebnis.fehler) > maximal:
                zeilen.append(t(u"  ... und %d weitere", u"  ... and %d more", u"  ... y %d más")
                              % (len(ergebnis.fehler) - maximal))
        self.status.Text = kurz + (t(u", %d fehlgeschlagen", u", %d failed", u", %d fallidos")
                                   % len(ergebnis.fehler)
                                   if ergebnis.fehler else u"")
        meldung(self.fenster, u"\n".join(zeilen),
                warnung=bool(ergebnis.fehler
                             or getattr(ergebnis, "abgelehnt", None)))

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp, uidoc):
    JoinMultipleFenster(uiapp, uidoc).zeige()
