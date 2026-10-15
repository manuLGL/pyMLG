# -*- coding: utf-8 -*-
"""Hauptfenster von Elements2D (WPF).

    ┌ 2D-Elemente im Projekt ───────────────┐ Gruppieren: ● Kategorie ○ Ansicht
    │ ☑ Alle                           812  │ Eingrenzen: ● alle ○ auf Plänen ○ nicht
    │   ☑ Detailgruppen                 14  │             ☑ direkt auf Plänen
    │     ☑ Detailgruppe                14  │ 812 markiert
    │       ☑ Treppe_Detail             14  │ [In Revit auswählen]
    │         ☑ Treppe_Detail [123] · EG 1  │ [Markierte löschen…]
    └───────────────────────────────────────┘

Wie FilterMore: Die Markierung ist eine Menge von Element-Ids, der Baum wird
erst beim Aufklappen befüllt. Doppelklick auf ein Element öffnet seine
Ansicht und zoomt darauf.
"""

import io
import json
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import FontWeights, Thickness, \
    VerticalAlignment  # noqa: E402
from System.Windows.Controls import CheckBox, Orientation, StackPanel, \
    TextBlock, TreeViewItem  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from elemente_2d import logik as lg  # noqa: E402
from elemente_2d import revit as rv  # noqa: E402
from filter_manager import dialoge as dlg  # noqa: E402
from filter_more import baum as bm  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402

TITEL = t(u"2D-Elemente", u"2D elements", u"Elementos 2D")


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
BLAU = _pinsel(40, 90, 160)

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(), "Elements2D_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "Elements2D.json")

OPTIONEN = ("nach_ansicht", "mit_plaenen")

# Optionsfeld -> Eingrenzung nach Platzierung
PLATZIERUNG = (
    ("platz_alle", lg.ALLE),
    ("platz_platziert", lg.PLATZIERTE),
    ("platz_unplatziert", lg.UNPLATZIERTE),
)

XAML_TEXTE = {
    "t0": (u"Aufklappen", u"Expand", u"Expandir"),
    "t1": (u"Bis zur dritten Ebene aufklappen",
           u"Expand down to the third level",
           u"Expandir hasta el tercer nivel"),
    "t2": (u"Zuklappen", u"Collapse", u"Contraer"),
    "t3": (u"Alle markieren", u"Check all", u"Marcar todo"),
    "t4": (u"Keine", u"None", u"Ninguno"),
    "t5": (u"Umkehren", u"Invert", u"Invertir"),
    "t6": (u"Gruppieren nach", u"Group by", u"Agrupar por"),
    "t7": (u"Kategorie > Familie > Typ", u"Category > family > type",
           u"Categoría > familia > tipo"),
    "t8": (u"Ansichtsart > Ansicht > Kategorie",
           u"View type > view > category",
           u"Tipo de vista > vista > categoría"),
    "t9": (u"Eingrenzen", u"Narrow down", u"Acotar"),
    "t10": (u"Nur Ansichten, die auf keinem Plan liegen",
            u"Only views not placed on any sheet",
            u"Solo vistas que no están en ningún plano"),
    "t11": (u"Zum Aufräumen: Inhalte von Ansichten, die nirgends gedruckt "
            u"werden",
            u"For cleaning up: contents of views that are never printed",
            u"Para limpiar: contenido de vistas que nunca se imprimen"),
    "t21": (u"Alle Ansichten", u"All views", u"Todas las vistas"),
    "t22": (u"Nur Ansichten, die auf einem Plan liegen",
            u"Only views placed on a sheet",
            u"Solo vistas colocadas en un plano"),
    "t24": (u"Pläne wählen…", u"Choose sheets…", u"Elegir planos…"),
    "t25": (u"Nur Ansichten auf bestimmten Plänen - und die Elemente direkt "
            u"auf diesen Plänen",
            u"Only views on certain sheets - and the elements placed directly "
            u"on those sheets",
            u"Solo vistas de ciertos planos y los elementos colocados "
            u"directamente en esos planos"),
    "t23": (u"Was gedruckt wird - bei abhängigen Ansichten zählt auch die "
            u"Hauptansicht",
            u"What gets printed - for dependent views the primary view "
            u"counts too",
            u"Lo que se imprime; en vistas dependientes cuenta también la "
            u"vista principal"),
    "t12": (u"Elemente direkt auf Plänen", u"Elements placed on sheets",
            u"Elementos colocados en planos"),
    "t13": (u"Texte, Linien usw., die direkt auf einem Plan liegen "
            u"(Planköpfe und Ansichtsfenster sind nie dabei)",
            u"Texts, lines etc. placed directly on a sheet (title blocks "
            u"and viewports are never included)",
            u"Textos, líneas, etc. colocados directamente en un plano "
            u"(cajetines y ventanas de vista nunca se incluyen)"),
    "t14": (u"In Revit auswählen", u"Select in Revit",
            u"Seleccionar en Revit"),
    "t15": (u"Markierte Elemente in Revit auswählen und schließen",
            u"Select the checked elements in Revit and close",
            u"Seleccionar en Revit los elementos marcados y cerrar"),
    "t16": (u"Markierte löschen…", u"Delete checked…",
            u"Eliminar marcados…"),
    "t17": (u"Markierte Elemente im ganzen Projekt löschen - mit Strg+Z "
            u"rückgängig zu machen",
            u"Delete the checked elements in the whole project - undo with "
            u"Ctrl+Z",
            u"Eliminar los elementos marcados en todo el proyecto; se "
            u"deshace con Ctrl+Z"),
    "t18": (u"Schließen", u"Close", u"Cerrar"),
    "t19": (u"2D-Elemente im Projekt", u"2D elements in the project",
            u"Elementos 2D del proyecto"),
    "t20": (u"Doppelklick auf ein Element: Ansicht öffnen und zoomen",
            u"Double-click an element: open its view and zoom",
            u"Doble clic en un elemento: abrir su vista y hacer zoom"),
}

XAML = u"""
<Window %s Title="Elements2D (pyMLG)" Width="1040" Height="820"
        MinWidth="760" MinHeight="560" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,2"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Margin" Value="0,2"/>
    </Style>
    <Style x:Key="klein" TargetType="Button">
      <Setter Property="Padding" Value="6,2"/>
      <Setter Property="Margin" Value="0,0,4,4"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="260"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Baum -->
    <DockPanel>
      <TextBlock DockPanel.Dock="Top" Text="{{t19}}" FontWeight="SemiBold"
                 Margin="0,0,0,4"/>
      <WrapPanel DockPanel.Dock="Top">
        <Button x:Name="aufklappen" Content="{{t0}}" Style="{StaticResource klein}"
                ToolTip="{{t1}}"/>
        <Button x:Name="zuklappen" Content="{{t2}}" Style="{StaticResource klein}"/>
        <Button x:Name="alle" Content="{{t3}}" Style="{StaticResource klein}"/>
        <Button x:Name="keine" Content="{{t4}}" Style="{StaticResource klein}"/>
        <Button x:Name="umkehren" Content="{{t5}}" Style="{StaticResource klein}"/>
      </WrapPanel>
      <TreeView x:Name="baum" Padding="2,4" ToolTip="{{t20}}"/>
    </DockPanel>

    <!-- Optionen -->
    <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
      <StackPanel>
        <GroupBox Header="{{t6}}">
          <StackPanel>
            <RadioButton x:Name="nach_kategorie" Content="{{t7}}"/>
            <RadioButton x:Name="nach_ansicht" Content="{{t8}}"/>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="{{t9}}">
          <StackPanel>
            <RadioButton x:Name="platz_alle" Content="{{t21}}"/>
            <RadioButton x:Name="platz_platziert" ToolTip="{{t23}}">
              <TextBlock Text="{{t22}}" TextWrapping="Wrap"/>
            </RadioButton>
            <DockPanel Margin="18,2,0,4">
              <Button x:Name="plaene_waehlen" Content="{{t24}}"
                      Padding="6,2" ToolTip="{{t25}}"/>
              <TextBlock x:Name="plaene_text" Foreground="#555"
                         VerticalAlignment="Center" Margin="8,0,0,0"
                         TextTrimming="CharacterEllipsis"/>
            </DockPanel>
            <RadioButton x:Name="platz_unplatziert" ToolTip="{{t11}}">
              <TextBlock Text="{{t10}}" TextWrapping="Wrap"/>
            </RadioButton>
            <CheckBox x:Name="mit_plaenen" ToolTip="{{t13}}" Margin="0,6,0,2">
              <TextBlock Text="{{t12}}" TextWrapping="Wrap"/>
            </CheckBox>
          </StackPanel>
        </GroupBox>

        <TextBlock x:Name="anzahl" FontSize="22" FontWeight="SemiBold"
                   Margin="0,6,0,0"/>
        <TextBlock x:Name="anzahl_text" Foreground="#555" TextWrapping="Wrap"
                   Margin="0,0,0,8"/>
        <Button x:Name="uebernehmen" Content="{{t14}}"
                FontWeight="Bold" Padding="8,9" Margin="0,0,0,6"
                ToolTip="{{t15}}"/>
        <Button x:Name="loeschen" Content="{{t16}}" Padding="8,9"
                Foreground="#B03030" ToolTip="{{t17}}"/>
      </StackPanel>
    </ScrollViewer>

    <!-- Fußzeile -->
    <DockPanel Grid.Row="1" Grid.ColumnSpan="3" Margin="0,10,0,0">
      <Button x:Name="schliessen" Content="{{t18}}" Width="100"
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
                    t(u"\n\nTechnische Details: %s",
                      u"\n\nTechnical details: %s",
                      u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
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


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------

class Elemente2DFenster(object):

    def __init__(self, uiapp, uidoc):
        self.uiapp = uiapp
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.alle_eintraege = []    # alle 2D-Elemente des Projekts
        self.elemente = {}          # Id -> Element (alle)
        self.eintraege = {}         # Id -> Eintrag (nach Eingrenzung)
        self.markiert = set()       # Ids
        self.wurzel = None
        self.ui_knoten = []         # Knoten mit Steuerelementen
        self.plaene = None          # gewählte Plan-Ids, None = alle Pläne

        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("baum", "nach_kategorie", "anzahl", "anzahl_text",
                     "plaene_waehlen", "plaene_text", "status") \
                + OPTIONEN + tuple(n for n, _p in PLATZIERUNG):
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        self.nach_ansicht.IsChecked = bool(einst.get("nach_ansicht", False))
        self.nach_kategorie.IsChecked = not self.nach_ansicht.IsChecked
        platzierung = einst.get("platzierung", lg.ALLE)
        for name, wert in PLATZIERUNG:
            getattr(self, name).IsChecked = wert == platzierung
        if self.platzierung() is None:
            self.platz_alle.IsChecked = True
        # Vor dem Start ausgewählte Pläne grenzen gleich ein
        self.plaene = rv.gewaehlte_plaene(uidoc)
        if self.plaene is not None:
            self.platz_platziert.IsChecked = True
        self.mit_plaenen.IsChecked = bool(einst.get("mit_plaenen", True))
        self._passe_optionen_an()

        self._verdrahte()
        self.lade()

    # --- Verdrahtung ------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        def klick(name, funktion):
            f.FindName(name).Click += s(lambda _s, _a: funktion())

        klick("aufklappen", lambda: self.klappe(True))
        klick("zuklappen", lambda: self.klappe(False))
        klick("alle", lambda: self.setze_markierung(set(self.eintraege)))
        klick("keine", lambda: self.setze_markierung(set()))
        klick("umkehren", lambda: self.setze_markierung(
            set(self.eintraege) - self.markiert))
        klick("uebernehmen", self.uebernehme)
        klick("loeschen", self.loesche)
        klick("plaene_waehlen", self.waehle_plaene)
        klick("schliessen", f.Close)
        for box in (self.nach_kategorie, self.nach_ansicht):
            box.Checked += s(lambda _s, _a: self.baue_baum())
        for name, _wert in PLATZIERUNG:
            getattr(self, name).Checked += s(lambda _s, _a: self.grenze_ein())
        self.mit_plaenen.Checked += s(lambda _s, _a: self.grenze_ein())
        self.mit_plaenen.Unchecked += s(lambda _s, _a: self.grenze_ein())
        f.Closing += s(lambda _s, _a: self._speichere())

    def _speichere(self):
        werte = dict((n, bool(getattr(self, n).IsChecked)) for n in OPTIONEN)
        werte["platzierung"] = self.platzierung() or lg.ALLE
        speichere_einstellungen(werte)

    def _passe_optionen_an(self):
        # Elemente auf Plänen liegen immer "auf einem Plan"
        self.mit_plaenen.IsEnabled = not self.platz_unplatziert.IsChecked
        self._zeige_plaene()

    def platzierung(self):
        for name, wert in PLATZIERUNG:
            if getattr(self, name).IsChecked:
                return wert
        return None

    def gruppierung(self):
        return lg.NACH_ANSICHT if self.nach_ansicht.IsChecked \
            else lg.NACH_KATEGORIE

    # --- Laden und Baum ---------------------------------------------------
    def lade(self):
        self.alle_eintraege, self.elemente = rv.sammle(self.doc)
        self.markiert = set()
        self.grenze_ein()

    def grenze_ein(self):
        self._passe_optionen_an()
        eintraege = lg.grenze_ein(
            self.alle_eintraege,
            platzierung=self.platzierung() or lg.ALLE,
            plaene=self.plaene,
            mit_plaenen=bool(self.mit_plaenen.IsChecked))
        self.eintraege = dict((e.element_id, e) for e in eintraege)
        self.markiert &= set(self.eintraege)
        self.baue_baum()
        self.status.Text = t(
            u"%d 2D-Elemente im Projekt, %d davon angezeigt.",
            u"%d 2D elements in the project, %d of them shown.",
            u"%d elementos 2D en el proyecto, %d mostrados.") % (
            len(self.alle_eintraege), len(self.eintraege))

    def baue_baum(self):
        self.wurzel = bm.baue(
            lg.datensaetze(self.eintraege.values(), self.gruppierung()),
            t(u"Alle", u"All", u"Todo"))
        self.baum.Items.Clear()
        self.ui_knoten = []
        item = self._item(self.wurzel)
        self.baum.Items.Add(item)
        item.IsExpanded = True
        self.aktualisiere()

    def _item(self, knoten):
        item = TreeViewItem()
        kopf = StackPanel()
        kopf.Orientation = Orientation.Horizontal

        box = CheckBox()
        box.VerticalAlignment = VerticalAlignment.Center
        box.Margin = Thickness(0.0, 1.0, 5.0, 1.0)
        box.Click += sicher(lambda: self.fenster, self._klick(knoten))
        kopf.Children.Add(box)

        name = TextBlock()
        name.Text = knoten.name
        name.VerticalAlignment = VerticalAlignment.Center
        if knoten.ebene == bm.ELEMENT:
            name.Foreground = GRAU
        elif knoten.ebene in (bm.WURZEL, bm.KATEGORIE):
            name.FontWeight = FontWeights.SemiBold
        kopf.Children.Add(name)

        zaehler = TextBlock()
        zaehler.Margin = Thickness(10.0, 0.0, 0.0, 0.0)
        zaehler.Foreground = BLAU
        zaehler.VerticalAlignment = VerticalAlignment.Center
        if knoten.ebene != bm.ELEMENT:
            kopf.Children.Add(zaehler)

        item.Header = kopf
        knoten.item, knoten.box, knoten.zaehler = item, box, zaehler
        knoten.geladen = False
        if knoten.kinder:
            item.Items.Add(u"…")     # Platzhalter, damit der Pfeil erscheint
            item.Expanded += sicher(lambda: self.fenster,
                                    self._aufklappen(knoten))
        if knoten.ebene == bm.ELEMENT:
            item.MouseDoubleClick += sicher(lambda: self.fenster,
                                            self._doppelklick(knoten))
        self.ui_knoten.append(knoten)
        self._zeige_zustand(knoten)
        return item

    def _aufklappen(self, knoten):
        def handler(_sender, args):
            # Das Ereignis steigt von Unterknoten auf - nur das eigene zählt
            if not args.OriginalSource.Equals(knoten.item) or knoten.geladen:
                return
            knoten.geladen = True
            knoten.item.Items.Clear()
            for kind in knoten.kinder:
                knoten.item.Items.Add(self._item(kind))
        return handler

    def _klick(self, knoten):
        def handler(sender, _args):
            if sender.IsChecked:
                self.markiert.update(knoten.ids)
            else:
                self.markiert.difference_update(knoten.ids)
            self.aktualisiere()
        return handler

    def _doppelklick(self, knoten):
        def handler(_sender, args):
            args.Handled = True
            element = self.elemente.get(knoten.element_id)
            if element is None or not element.IsValidObject:
                return
            rv.zeige(self.uidoc, element)
            self.status.Text = t(u"Gezeigt: %s", u"Shown: %s",
                                 u"Mostrado: %s") % knoten.name
        return handler

    def _zeige_zustand(self, knoten):
        zustand, text = bm.anzeige(knoten, self.markiert)
        knoten.box.IsChecked = zustand
        knoten.zaehler.Text = text

    def aktualisiere(self):
        for knoten in self.ui_knoten:
            self._zeige_zustand(knoten)
        self.anzahl.Text = u"%d" % len(self.markiert)
        markierte = [self.eintraege[i] for i in self.markiert]
        self.anzahl_text.Text = t(
            u"von %d Elementen markiert, in %d Ansichten",
            u"of %d elements checked, in %d views",
            u"de %d elementos marcados, en %d vistas") % (
            len(self.eintraege), lg.anzahl_ansichten(markierte))

    def setze_markierung(self, ids):
        self.markiert = set(ids)
        self.aktualisiere()

    def klappe(self, auf):
        """Aufklappen bis zur dritten Ebene bzw. alles unter 'Alle'
        zuklappen."""
        if self.wurzel is None:
            return

        def auf_bis_typ(knoten):
            if knoten.item is None or knoten.ebene >= bm.TYP:
                return
            knoten.item.IsExpanded = True     # lädt die Kinder
            for kind in knoten.kinder:
                auf_bis_typ(kind)

        if auf:
            auf_bis_typ(self.wurzel)
        else:
            for knoten in self.ui_knoten:
                if knoten.ebene != bm.WURZEL and knoten.item is not None:
                    knoten.item.IsExpanded = False

    # --- Pläne ------------------------------------------------------------
    def waehle_plaene(self):
        alle = rv.plaene(self.doc)
        if not alle:
            meldung(self.fenster, t(u"Das Projekt hat keine Pläne.",
                                    u"The project has no sheets.",
                                    u"El proyecto no tiene planos."))
            return
        gewaehlt = dlg.waehle(
            self.fenster, TITEL,
            t(u"Nur 2D-Elemente der Ansichten auf diesen Plänen zeigen:",
              u"Only show 2D elements of the views on these sheets:",
              u"Mostrar solo elementos 2D de las vistas de estos planos:"),
            alle, mehrfach=True,
            vorauswahl=self.plaene if self.plaene is not None
            else [w for _n, w in alle])
        if gewaehlt is None:
            return
        self.plaene = None if len(gewaehlt) == len(alle) else set(gewaehlt)
        if self.platz_platziert.IsChecked:
            self.grenze_ein()
        else:
            self.platz_platziert.IsChecked = True     # grenzt selbst ein

    def _zeige_plaene(self):
        if self.plaene is None:
            self.plaene_text.Text = t(u"alle Pläne", u"all sheets",
                                      u"todos los planos")
        else:
            self.plaene_text.Text = t(u"%d Pläne gewählt",
                                      u"%d sheets chosen",
                                      u"%d planos elegidos") % len(self.plaene)

    # --- Auswählen und Löschen --------------------------------------------
    def _pruefe_markierung(self):
        if not self.markiert:
            meldung(self.fenster, t(u"Es ist kein Element markiert.",
                                    u"No element is checked.",
                                    u"No hay ningún elemento marcado."))
            return False
        return True

    def uebernehme(self):
        if not self._pruefe_markierung():
            return
        rv.waehle_aus(self.uidoc, self.markiert)
        self.fenster.Close()

    def loesche(self):
        if not self._pruefe_markierung():
            return
        markierte = [self.eintraege[i] for i in self.markiert]
        if not dlg.frage(self.fenster, t(
                u"%d 2D-Elemente in %d Ansichten löschen?\n\n"
                u"Abhängige Elemente (z.B. Beschriftungen dieser Elemente) "
                u"werden mitgelöscht. Rückgängig mit Strg+Z.",
                u"Delete %d 2D elements in %d views?\n\n"
                u"Dependent elements (e.g. tags of these elements) are "
                u"deleted too. Undo with Ctrl+Z.",
                u"¿Eliminar %d elementos 2D en %d vistas?\n\n"
                u"Los elementos dependientes (p. ej. etiquetas de estos "
                u"elementos) también se eliminan. Se deshace con Ctrl+Z.")
                % (len(markierte), lg.anzahl_ansichten(markierte)),
                titel=TITEL, warnung=True):
            return
        ids = self._sichere_ids(set(self.markiert))
        if not ids:
            self.status.Text = t(u"Nichts gelöscht.", u"Nothing deleted.",
                                 u"No se ha eliminado nada.")
            return
        geloescht, fehler = rv.loesche(self.doc, sorted(ids))
        self.lade()
        text = t(u"%d Elemente gelöscht (inklusive abhängiger).",
                 u"%d elements deleted (including dependent ones).",
                 u"%d elementos eliminados (incluidos los dependientes).") \
            % len(geloescht)
        self.status.Text = text
        if fehler:
            zeilen = [u"• %s%s" % (u"[%d] " % wert if wert else u"", grund)
                      for wert, grund in fehler[:15]]
            if len(fehler) > 15:
                zeilen.append(t(u"… und %d weitere", u"… and %d more",
                                u"… y %d más") % (len(fehler) - 15))
            meldung(self.fenster, text + t(
                u"\n\nNicht gelöscht:\n", u"\n\nNot deleted:\n",
                u"\n\nNo eliminados:\n") + u"\n".join(zeilen), warnung=True)

    def _sichere_ids(self, ids):
        """Probelauf vor dem Löschen: Würde Revit dabei ein Modell-Element
        (3D, Raster, Ebene ...) mitlöschen, werden die verursachenden
        markierten Elemente nach Rückfrage ausgelassen. Rückgabe: Ids, die
        gefahrlos gelöscht werden können, oder leer."""
        gefaehrdet, ursachen = rv.pruefe_loeschen(self.doc, sorted(ids))
        if not gefaehrdet:
            return ids
        zeilen = [u"• " + text for _w, text in gefaehrdet[:15]]
        if len(gefaehrdet) > 15:
            zeilen.append(t(u"… und %d weitere", u"… and %d more",
                            u"… y %d más") % (len(gefaehrdet) - 15))
        liste = u"\n".join(zeilen)
        rest = ids - ursachen
        if not ursachen or not rest:
            meldung(self.fenster, t(
                u"Abgebrochen - nichts wurde gelöscht.\n\nRevit würde dabei "
                u"%d Modell-Elemente mitlöschen:\n\n%s",
                u"Cancelled - nothing was deleted.\n\nRevit would also "
                u"delete %d model elements:\n\n%s",
                u"Cancelado; no se ha eliminado nada.\n\nRevit eliminaría "
                u"también %d elementos del modelo:\n\n%s")
                % (len(gefaehrdet), liste), warnung=True)
            return set()
        if not dlg.frage(self.fenster, t(
                u"Revit würde %d Modell-Elemente mitlöschen:\n\n%s\n\n"
                u"Sie hängen an %d der markierten Elemente. Diese auslassen "
                u"und nur die übrigen %d löschen?",
                u"Revit would also delete %d model elements:\n\n%s\n\n"
                u"They depend on %d of the checked elements. Skip those and "
                u"delete only the other %d?",
                u"Revit eliminaría también %d elementos del modelo:\n\n%s"
                u"\n\nDependen de %d de los elementos marcados. ¿Omitirlos "
                u"y eliminar solo los otros %d?")
                % (len(gefaehrdet), liste, len(ursachen), len(rest)),
                titel=TITEL, warnung=True):
            return set()
        # Zweiter Probelauf: GetDependentElements kennt nicht jede Kette
        if rv.pruefe_loeschen(self.doc, sorted(rest))[0]:
            meldung(self.fenster, t(
                u"Abgebrochen - nichts wurde gelöscht. Auch ohne diese "
                u"Elemente würde Revit Modell-Elemente mitlöschen.",
                u"Cancelled - nothing was deleted. Even without these "
                u"elements Revit would delete model elements.",
                u"Cancelado; no se ha eliminado nada. Incluso sin estos "
                u"elementos Revit eliminaría elementos del modelo."),
                warnung=True)
            return set()
        return rest

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp, uidoc):
    Elemente2DFenster(uiapp, uidoc).zeige()
