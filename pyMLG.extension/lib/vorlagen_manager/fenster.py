# -*- coding: utf-8 -*-
"""Hauptfenster des Ansichtsvorlagen-Managers (WPF).

Aufgebaut wie der native Dialog "Ansichtsvorlagen", aber mit Suche und
Mehrfachauswahl: Links die Vorlagen (mehrere gleichzeitig markierbar),
rechts deren Parameter.

    ┌ Ansichtsvorlagen ───────┬ Parameter ───────────────────────────────┐
    │ Suchen: [_____]         │ Suchen: [____]  □ nur eingeschlossene    │
    │ in: [Name ▾]            │ Parameter      │ Wert       │ Einschl.   │
    │ Disziplin:   [<alle> ▾] │ Ansichtsmaßstab│ 1 : 100    │    ☑       │
    │ Ansichtstyp: [<alle> ▾] │ Detailgrad     │ Grob       │    ☑       │
    │ ┌─────────────────────┐ │ Filter (V/G)   │ 4 Filter   │    ☑       │
    │ │ 360_ACOTADO     (6) │ │ …                                        │
    │ └─────────────────────┘ │ [Alle][Keine][Kopieren…][Holen…]         │
    │ Neu… Dupl. Umb. Löschen │                                          │
    └─────────────────────────┴──────────────────────────────────────────┘
    [Auf Ansichten anwenden…] [Zuweisung lösen…]        [OK] [Abbrechen]

"Auf andere Vorlagen kopieren…" und "Von anderer Vorlage holen…" sind
dieselbe Sache aus beiden Richtungen: Quelle und Ziele festlegen, dann in
vorlagen_manager.uebertragen einzeln abhaken, was wandern soll.

Sind mehrere Vorlagen markiert, zeigt die Tabelle die Parameter, die alle
gemeinsam haben. Gleiche Werte stehen im Klartext, unterschiedliche als
"<verschieden>"; eine Eingabe schreibt in alle markierten Vorlagen.

Einstellungen, die kein einfacher Wert sind (Kategorien, Filter,
Bearbeitungsbereiche, Ansichtsbereich), stehen als Knopf in der Wertspalte
und öffnen ihren Editor in vorlagen_manager.bloecke - auch der schreibt in
alle markierten Vorlagen. Was die API gar nicht freigibt (Modelldarstellung,
Schatten, Skizzenlinien, Beleuchtung, Fotobelichtung), bleibt grau und
lässt sich nur übertragen.

Transaktionen: Das Fenster läuft in einer TransactionGroup, jede Änderung
ist eine eigene Transaktion darin. Anders als der Filter-Manager wird sofort
geschrieben - es gibt kein "Anwenden". OK -> Assimilate (ein
Rückgängig-Schritt), Abbrechen -> RollBack, verwirft also alles.

CPython-Besonderheiten siehe filter_manager/dialoge.py.
"""

from Autodesk.Revit.DB import (
    ElementId,
    Transaction,
    TransactionGroup,
    TransactionStatus,
)
from System import Action
from System.Windows import (
    GridLength,
    HorizontalAlignment,
    TextTrimming,
    Thickness,
    VerticalAlignment,
)
from System.Windows.Controls import (
    Border,
    Button,
    CheckBox,
    ComboBox,
    ComboBoxItem,
    Grid,
    ListBoxItem,
    RowDefinition,
    TextBlock,
    TextBox,
)
from System.Windows.Input import Key
from System.Windows.Media import Color, SolidColorBrush
from System.Windows.Threading import DispatcherPriority

from filter_manager.aufloesung import id_liste, id_wert
from filter_manager import dialoge as dlg
from mlg_sprache import t, uebersetze_xaml
from vorlagen_manager import bloecke as bl
from vorlagen_manager import logik as lg
from vorlagen_manager import revit as rv
from vorlagen_manager import uebertragen as ue

FACHLICH = (rv.VorlagenFehler,)

TITEL = t(u"Ansichtsvorlagen-Manager", u"View Template Manager",
          u"Gestor de plantillas de vista")

SUCHE_NAME = 0
SUCHE_WERTE = 1


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem ladenden Thread und WPF
    # verweigert ihn in Fenstern anderer Threads
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
ZEILE_HELL = _pinsel(248, 249, 250)

VERSCHIEDEN_TEXT = t(u"<verschieden>", u"<varies>", u"<varía>")
ALLE_TEXT = t(u"<alle>", u"<all>", u"<todos>")

XAML_TEXTE = {
    "t0": (u"Ansichtsvorlagen-Manager (pyMLG)",
           u"View Template Manager (pyMLG)",
           u"Gestor de plantillas de vista (pyMLG)"),
    "t1": (u"Ansichtsvorlagen",
           u"View templates",
           u"Plantillas de vista"),
    "t2": (u"Suchen:",
           u"Search:",
           u"Buscar:"),
    "t3": (u"Mehrere Wörter: alle müssen vorkommen",
           u"Several words: all must occur",
           u"Varias palabras: deben aparecer todas"),
    "t4": (u"Name der Vorlage",
           u"Template name",
           u"Nombre de la plantilla"),
    "t5": (u"Name und Parameterwerte",
           u"Name and parameter values",
           u"Nombre y valores de parámetros"),
    "t6": (u"Disziplin:",
           u"Discipline:",
           u"Disciplina:"),
    "t7": (u"Ansichtstyp:",
           u"View type:",
           u"Tipo de vista:"),
    "t8": (u"Neu aus Ansicht…",
           u"New from view…",
           u"Nueva desde vista…"),
    "t9": (u"Duplizieren",
           u"Duplicate",
           u"Duplicar"),
    "t10": (u"Umbenennen…",
            u"Rename…",
            u"Renombrar…"),
    "t11": (u"Löschen",
            u"Delete",
            u"Eliminar"),
    "t12": (u"Parameter",
            u"Parameters",
            u"Parámetros"),
    "t13": (u"Nur eingeschlossene",
            u"Included only",
            u"Solo incluidos"),
    "t14": (u"Wert",
            u"Value",
            u"Valor"),
    "t15": (u"Einschließen",
            u"Include",
            u"Incluir"),
    "t16": (u"Alle einschließen",
            u"Include all",
            u"Incluir todos"),
    "t17": (u"Keine einschließen",
            u"Include none",
            u"No incluir ninguno"),
    "t18": (u"Von anderer Vorlage holen…",
            u"Get from another template…",
            u"Traer de otra plantilla…"),
    "t23": (u"Auf andere Vorlagen kopieren…",
            u"Copy to other templates…",
            u"Copiar a otras plantillas…"),
    "t19": (u"Auf Ansichten anwenden…",
            u"Apply to views…",
            u"Aplicar a vistas…"),
    "t20": (u"Zuweisung lösen…",
            u"Remove assignment…",
            u"Quitar asignación…"),
    "t21": (u"Abbrechen",
            u"Cancel",
            u"Cancelar"),
    "t22": (u"Abbrechen verwirft alle Änderungen dieser Sitzung.",
            u"Cancel discards all changes of this session.",
            u"Cancelar descarta todos los cambios de esta sesión."),
}

XAML = u"""
<Window %s Title="{{t0}}" Width="1180" Height="720"
        MinWidth="880" MinHeight="480" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,3"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="1*" MinWidth="260"/>
      <ColumnDefinition Width="2.2*" MinWidth="470"/>
    </Grid.ColumnDefinitions>

    <!-- Vorlagen -->
    <GroupBox Header="{{t1}}" Grid.Column="0">
      <DockPanel>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Text="{{t2}}" VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <TextBox x:Name="suche" Grid.Column="1" Padding="3"
                   ToolTip="{{t3}}"/>
          <TextBlock Grid.Row="1" Text="in:" VerticalAlignment="Center"
                     Margin="0,4,6,0"/>
          <ComboBox x:Name="suchmodus" Grid.Row="1" Grid.Column="1"
                    Margin="0,4,0,0" SelectedIndex="0">
            <ComboBoxItem Content="{{t4}}"/>
            <ComboBoxItem Content="{{t5}}"/>
          </ComboBox>
          <TextBlock Grid.Row="2" Text="{{t6}}" VerticalAlignment="Center"
                     Margin="0,4,6,0"/>
          <ComboBox x:Name="disziplin" Grid.Row="2" Grid.Column="1"
                    Margin="0,4,0,0"/>
          <TextBlock Grid.Row="3" Text="{{t7}}" VerticalAlignment="Center"
                     Margin="0,4,6,0"/>
          <ComboBox x:Name="typfilter" Grid.Row="3" Grid.Column="1"
                    Margin="0,4,0,0"/>
        </Grid>
        <WrapPanel DockPanel.Dock="Bottom" Margin="0,2,0,0">
          <Button x:Name="neu" Content="{{t8}}" Margin="0,4,6,0"/>
          <Button x:Name="duplizieren" Content="{{t9}}" Margin="0,4,6,0"/>
          <Button x:Name="umbenennen" Content="{{t10}}" Margin="0,4,6,0"/>
          <Button x:Name="loeschen" Content="{{t11}}" Margin="0,4,6,0"/>
        </WrapPanel>
        <TextBlock x:Name="vorlagenzaehler" DockPanel.Dock="Bottom"
                   Foreground="#666" Margin="0,4,0,0"/>
        <ListBox x:Name="vorlagenliste" SelectionMode="Extended"/>
      </DockPanel>
    </GroupBox>

    <!-- Parameter -->
    <GroupBox Header="{{t12}}" Grid.Column="1" Margin="0">
      <DockPanel>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,4">
          <TextBlock Text="{{t2}}" VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <CheckBox x:Name="nurenthalten" DockPanel.Dock="Right"
                    Content="{{t13}}" VerticalAlignment="Center"
                    Margin="8,0,0,0"/>
          <TextBox x:Name="psuche" Padding="3"/>
        </DockPanel>
        <TextBlock x:Name="parameterhinweis" DockPanel.Dock="Top"
                   TextWrapping="Wrap" Foreground="#666" Margin="0,0,0,6"/>
        <WrapPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
          <Button x:Name="enthalten_alle" Content="{{t16}}" Margin="0,2,6,0"/>
          <Button x:Name="enthalten_keine" Content="{{t17}}" Margin="0,2,6,0"/>
          <Button x:Name="kopieren" Content="{{t23}}" Margin="0,2,6,0"
                  FontWeight="Bold"/>
          <Button x:Name="uebertragen" Content="{{t18}}" Margin="0,2,6,0"/>
        </WrapPanel>
        <!-- Rechter Rand: Platz fuer die Bildlaufleiste der Tabelle -->
        <Grid x:Name="kopf" DockPanel.Dock="Top" Margin="0,0,18,2">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="1.4*"/>
            <ColumnDefinition Width="1.6*"/>
            <ColumnDefinition Width="100"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="{{t12}}" FontWeight="Bold" Margin="4,0"/>
          <TextBlock Grid.Column="1" Text="{{t14}}" FontWeight="Bold"
                     Margin="4,0"/>
          <TextBlock Grid.Column="2" Text="{{t15}}" FontWeight="Bold"
                     HorizontalAlignment="Center"/>
        </Grid>
        <Border BorderBrush="#ABADB3" BorderThickness="1">
          <ScrollViewer x:Name="tabellescroll"
                        VerticalScrollBarVisibility="Visible"
                        HorizontalScrollBarVisibility="Disabled">
            <Grid x:Name="tabelle">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.4*"/>
                <ColumnDefinition Width="1.6*"/>
                <ColumnDefinition Width="100"/>
              </Grid.ColumnDefinitions>
            </Grid>
          </ScrollViewer>
        </Border>
      </DockPanel>
    </GroupBox>

    <!-- Fusszeile -->
    <DockPanel Grid.Row="1" Grid.ColumnSpan="2" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="OK" Width="90"/>
        <Button x:Name="abbrechen" Content="{{t21}}" Width="90"
                IsCancel="True" Margin="0" ToolTip="{{t22}}"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal">
        <Button x:Name="anwenden_auf" Content="{{t19}}"/>
        <Button x:Name="zuweisung_loesen" Content="{{t20}}"/>
        <TextBlock x:Name="status" VerticalAlignment="Center"
                   Foreground="#666" Margin="8,0"
                   TextTrimming="CharacterEllipsis"/>
      </StackPanel>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


class VorlagenFenster(object):

    def __init__(self, uiapp, doc):
        self.uiapp = uiapp
        self.doc = doc
        self.modell = rv.Modell(doc)
        self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, handle=uiapp.MainWindowHandle)
        self.c = self.fenster.FindName

        self.vorlagen = []           # alle Vorlagen (View)
        self.nach_id = {}            # Id-Wert -> View
        self.verwendung = {}         # Id-Wert -> Anzahl Ansichten
        self.ausgewaehlt = []        # Id-Werte der markierten Vorlagen
        self.zeilen = {}             # pid -> [Eintrag je markierter Vorlage]
        self.reihenfolge = []        # pids in Anzeigereihenfolge
        self.suchtexte = {}          # Id-Wert -> Suchtext mit Werten
        self.sitzung_geaendert = False
        self.uebernehmen = False
        self._still = False
        self._neubau_geplant = False
        self._verdrahte()

    # ------------------------------------------------------------------
    # Grundgerüst
    # ------------------------------------------------------------------

    def _h(self, funktion):
        return dlg.sicher(lambda: self.fenster, funktion, FACHLICH)

    def _verdrahte(self):
        c = self.c
        c("suche").TextChanged += self._h(lambda s, a: self._fuelle_liste())
        c("suchmodus").SelectionChanged += self._h(self._bei_suchmodus)
        c("disziplin").SelectionChanged += self._h(
            lambda s, a: self._fuelle_liste())
        c("typfilter").SelectionChanged += self._h(
            lambda s, a: self._fuelle_liste())
        c("vorlagenliste").SelectionChanged += self._h(self._bei_auswahl)
        c("neu").Click += self._h(self._neu)
        c("duplizieren").Click += self._h(self._duplizieren)
        c("umbenennen").Click += self._h(self._umbenennen)
        c("loeschen").Click += self._h(self._loeschen)
        c("psuche").TextChanged += self._h(lambda s, a: self._baue_tabelle())
        c("nurenthalten").Click += self._h(lambda s, a: self._baue_tabelle())
        c("enthalten_alle").Click += self._h(
            lambda s, a: self._alle_haken(True))
        c("enthalten_keine").Click += self._h(
            lambda s, a: self._alle_haken(False))
        c("uebertragen").Click += self._h(self._uebertragen)
        c("kopieren").Click += self._h(self._kopieren)
        c("anwenden_auf").Click += self._h(self._auf_ansichten)
        c("zuweisung_loesen").Click += self._h(self._zuweisung_loesen)
        c("ok").Click += self._h(self._ok)
        self.fenster.Closing += self._h(self._beim_schliessen)

    def zeige(self):
        gruppe = TransactionGroup(self.doc, t(u"pyMLG Ansichtsvorlagen",
                                              u"pyMLG View Templates",
                                              u"pyMLG Plantillas de vista"))
        gruppe.Start()
        try:
            self._lade()
            self._fuelle_filter()
            self._fuelle_liste()
            self._zeige_auswahl()
            self.fenster.ShowDialog()
        finally:
            if self.uebernehmen and self.sitzung_geaendert:
                gruppe.Assimilate()
            else:
                gruppe.RollBack()

    def _transaktion(self, name, funktion):
        """Führt funktion() in einer eigenen Transaktion aus."""
        transaktion = Transaction(self.doc, name)
        transaktion.Start()
        try:
            ergebnis = funktion()
            if transaktion.Commit() != TransactionStatus.Committed:
                raise rv.VorlagenFehler(
                    t(u"Revit hat die Änderung \"%s\" nicht übernommen.",
                      u"Revit did not accept the change \"%s\".",
                      u"Revit no aceptó el cambio \"%s\".") % name)
        except Exception:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            raise
        self.sitzung_geaendert = True
        return ergebnis

    # ------------------------------------------------------------------
    # Daten laden
    # ------------------------------------------------------------------

    def _lade(self):
        self.vorlagen = self.modell.vorlagen()
        self.nach_id = dict((id_wert(v.Id), v) for v in self.vorlagen)
        self.verwendung = self.modell.verwendung()
        self.suchtexte = {}

    def _disziplin_wert(self, vorlage):
        try:
            return int(vorlage.Discipline)
        except Exception:
            return None

    def _fuelle_filter(self):
        """Auswahlfelder für Disziplin und Ansichtstyp aus den Vorlagen."""
        beschriftung = dict(
            (wert, text) for wert, text
            in rv.AUFZAEHLUNGEN.get("VIEW_DISCIPLINE", []))
        disziplinen, typen = {}, {}
        for vorlage in self.vorlagen:
            wert = self._disziplin_wert(vorlage)
            if wert is not None:
                disziplinen.setdefault(
                    wert, beschriftung.get(wert, u"%d" % wert))
            typen.setdefault(u"%s" % vorlage.ViewType,
                             rv.ansichtstyp(vorlage))
        self._still = True
        try:
            self._fuelle_box("disziplin", ALLE_TEXT, sorted(
                disziplinen.items(), key=lambda e: lg.natuerlich(e[1])))
            self._fuelle_box("typfilter", ALLE_TEXT, sorted(
                typen.items(), key=lambda e: lg.natuerlich(e[1])))
        finally:
            self._still = False

    def _fuelle_box(self, name, erster, eintraege):
        box = self.c(name)
        box.Items.Clear()
        kopf = ComboBoxItem()
        kopf.Content = erster
        kopf.Tag = None
        box.Items.Add(kopf)
        for wert, text in eintraege:
            item = ComboBoxItem()
            item.Content = text
            item.Tag = wert
            box.Items.Add(item)
        box.SelectedIndex = 0

    def _suchtext(self, vorlage):
        """Name und alle Parameterwerte einer Vorlage, klein geschrieben."""
        schluessel = id_wert(vorlage.Id)
        vorhanden = self.suchtexte.get(schluessel)
        if vorhanden is None:
            teile = [vorlage.Name]
            try:
                for eintrag in self.modell.eintraege(
                        vorlage, mit_bloecken=False).values():
                    teile.append(eintrag.name)
                    teile.append(eintrag.text)
            except Exception:
                pass
            vorhanden = u" ".join(teile).lower()
            self.suchtexte[schluessel] = vorhanden
        return vorhanden

    # ------------------------------------------------------------------
    # Vorlagenliste
    # ------------------------------------------------------------------

    def _bei_suchmodus(self, sender, args):
        if self._still:
            return
        if self.c("suchmodus").SelectedIndex == SUCHE_WERTE:
            # Werte aller Vorlagen einlesen - das dauert einen Moment
            self.c("status").Text = t(u"Parameterwerte werden gelesen…",
                                      u"Reading parameter values…",
                                      u"Leyendo valores de parámetros…")
            for vorlage in self.vorlagen:
                self._suchtext(vorlage)
            self.c("status").Text = u""
        self._fuelle_liste()

    def _gewaehlter_filter(self, name):
        box = self.c(name)
        item = box.SelectedItem
        return None if item is None else item.Tag

    def _fuelle_liste(self):
        if self._still:
            return
        woerter = self.c("suche").Text.lower().split()
        mit_werten = self.c("suchmodus").SelectedIndex == SUCHE_WERTE
        disziplin = self._gewaehlter_filter("disziplin")
        typ = self._gewaehlter_filter("typfilter")
        liste = self.c("vorlagenliste")

        passend = []
        for vorlage in self.vorlagen:
            if disziplin is not None and \
                    self._disziplin_wert(vorlage) != disziplin:
                continue
            if typ is not None and u"%s" % vorlage.ViewType != typ:
                continue
            text = (self._suchtext(vorlage) if mit_werten
                    else vorlage.Name.lower())
            if lg.passt(text, woerter):
                passend.append(vorlage)

        self._still = True
        try:
            liste.Items.Clear()
            for vorlage in passend:
                schluessel = id_wert(vorlage.Id)
                anzahl = self.verwendung.get(schluessel, 0)
                item = ListBoxItem()
                item.Content = u"%s   (%d)" % (vorlage.Name, anzahl)
                item.Tag = schluessel
                item.ToolTip = self._tooltip(vorlage, anzahl)
                liste.Items.Add(item)
                if schluessel in self.ausgewaehlt:
                    item.IsSelected = True
        finally:
            self._still = False
        self.c("vorlagenzaehler").Text = t(
            u"%d von %d Vorlagen", u"%d of %d templates",
            u"%d de %d plantillas") % (len(passend), len(self.vorlagen))

    def _tooltip(self, vorlage, anzahl):
        zeilen = [t(u"Ansichtstyp: %s", u"View type: %s", u"Tipo de vista: %s")
                  % rv.ansichtstyp(vorlage),
                  t(u"Ansichten mit dieser Vorlage: %d",
                    u"Views using this template: %d",
                    u"Vistas con esta plantilla: %d") % anzahl]
        if anzahl:
            namen = [a.Name for a in self.modell.ansichten_mit(vorlage)[:20]]
            zeilen.extend(u"• " + name for name in namen)
            if anzahl > 20:
                zeilen.append(t(u"… und %d weitere", u"… and %d more",
                                u"… y %d más") % (anzahl - 20))
        return u"\n".join(zeilen)

    def _markierte_ids(self):
        return [item.Tag for item in self.c("vorlagenliste").SelectedItems]

    def _markierte(self):
        return [self.nach_id[i] for i in self.ausgewaehlt if i in self.nach_id]

    def _bei_auswahl(self, sender, args):
        if self._still:
            return
        neu = self._markierte_ids()
        if not neu or sorted(neu) == sorted(self.ausgewaehlt):
            # Leere Markierung entsteht auch beim Neuaufbau der Liste
            return
        self.ausgewaehlt = neu
        self._zeige_auswahl()

    def _waehle(self, ids):
        self.ausgewaehlt = list(ids)
        self.c("suche").Text = u""
        self._fuelle_liste()
        self._zeige_auswahl()

    # ------------------------------------------------------------------
    # Parametertabelle
    # ------------------------------------------------------------------

    def _zeige_auswahl(self):
        markierte = self._markierte()
        anzahl = len(markierte)
        for name in ("umbenennen", "anwenden_auf", "zuweisung_loesen",
                     "kopieren"):
            self.c(name).IsEnabled = anzahl == 1
        for name in ("duplizieren", "loeschen", "enthalten_alle",
                     "enthalten_keine", "uebertragen"):
            self.c(name).IsEnabled = anzahl > 0

        self.zeilen = {}
        self.reihenfolge = []
        if markierte:
            # Die Zusammenfassung der Blöcke zählt Kategorien durch - das
            # lohnt nur für die erste Vorlage, deren Stand die Editoren
            # ohnehin zeigen
            je_vorlage = [self.modell.eintraege(v, mit_bloecken=nummer == 0)
                          for nummer, v in enumerate(markierte)]
            gemeinsam = set(je_vorlage[0])
            for weitere in je_vorlage[1:]:
                gemeinsam &= set(weitere)
            for pid in gemeinsam:
                self.zeilen[pid] = [e[pid] for e in je_vorlage]
            self.reihenfolge = sorted(
                gemeinsam, key=lambda p: self.zeilen[p][0].sortierung())

        hinweis = self.c("parameterhinweis")
        if not markierte:
            hinweis.Text = t(u"Links eine oder mehrere Vorlagen auswählen.",
                             u"Select one or more templates on the left.",
                             u"Seleccione una o varias plantillas a la "
                             u"izquierda.")
        elif anzahl == 1:
            hinweis.Text = t(u"\"%s\" – %d Parameter.",
                             u"\"%s\" – %d parameters.",
                             u"\"%s\" – %d parámetros.") % (
                markierte[0].Name, len(self.reihenfolge))
        else:
            hinweis.Text = t(
                u"%d Vorlagen markiert – gezeigt werden die %d Parameter, die "
                u"alle gemeinsam haben. Eine Eingabe schreibt in alle.",
                u"%d templates selected – showing the %d parameters they all "
                u"have in common. An entry writes to all of them.",
                u"%d plantillas seleccionadas: se muestran los %d parámetros "
                u"comunes. Una entrada se escribe en todas.") % (
                anzahl, len(self.reihenfolge))
        self._baue_tabelle()
        self._zeige_status()

    def _zeige_status(self):
        markierte = self._markierte()
        if len(markierte) != 1:
            self.c("status").Text = u""
            return
        anzahl = self.verwendung.get(id_wert(markierte[0].Id), 0)
        self.c("status").Text = t(
            u"Ansichten mit dieser Vorlage: %d",
            u"Views using this template: %d",
            u"Vistas con esta plantilla: %d") % anzahl

    def _sichtbare_pids(self):
        woerter = self.c("psuche").Text.lower().split()
        nur_enthalten = bool(self.c("nurenthalten").IsChecked)
        sichtbar = []
        for pid in self.reihenfolge:
            eintraege = self.zeilen[pid]
            if nur_enthalten and not any(e.enthalten for e in eintraege):
                continue
            if woerter and not lg.passt(eintraege[0].suchtext(), woerter):
                continue
            sichtbar.append(pid)
        return sichtbar

    def _baue_tabelle(self):
        tabelle = self.c("tabelle")
        scroll = self.c("tabellescroll")
        position = scroll.VerticalOffset
        tabelle.Children.Clear()
        tabelle.RowDefinitions.Clear()

        self._still = True
        try:
            for nummer, pid in enumerate(self._sichtbare_pids()):
                zeile = RowDefinition()
                zeile.Height = GridLength.Auto
                tabelle.RowDefinitions.Add(zeile)
                self._baue_zeile(tabelle, nummer, self.zeilen[pid])
        finally:
            self._still = False
        self.fenster.Dispatcher.InvokeAsync(
            Action(lambda: scroll.ScrollToVerticalOffset(position)),
            DispatcherPriority.Background)

    def _baue_zeile(self, tabelle, nummer, eintraege):
        erster = eintraege[0]
        if nummer % 2:
            hintergrund = Border()
            hintergrund.Background = ZEILE_HELL
            Grid.SetRow(hintergrund, nummer)
            Grid.SetColumnSpan(hintergrund, 3)
            tabelle.Children.Add(hintergrund)

        name = TextBlock()
        name.Text = erster.name
        name.Margin = Thickness(4.0, 4.0, 4.0, 4.0)
        name.TextTrimming = TextTrimming.CharacterEllipsis
        name.ToolTip = erster.name
        name.VerticalAlignment = VerticalAlignment.Center
        Grid.SetRow(name, nummer)
        tabelle.Children.Add(name)

        wert = self._wert_element(eintraege)
        Grid.SetRow(wert, nummer)
        Grid.SetColumn(wert, 1)
        tabelle.Children.Add(wert)

        haken = CheckBox()
        haken.HorizontalAlignment = HorizontalAlignment.Center
        haken.VerticalAlignment = VerticalAlignment.Center
        haken.Margin = Thickness(0.0, 4.0, 0.0, 4.0)
        zustand = lg.vereine_flagge([e.enthalten for e in eintraege])
        haken.IsThreeState = zustand is None
        haken.IsChecked = zustand
        haken.Tag = erster.pid
        haken.ToolTip = t(u"Parameter wird von der Vorlage gesteuert",
                          u"Parameter is controlled by the template",
                          u"El parámetro lo controla la plantilla")
        haken.Click += self._h(self._bei_haken)
        Grid.SetRow(haken, nummer)
        Grid.SetColumn(haken, 2)
        tabelle.Children.Add(haken)

    def _wert_element(self, eintraege):
        erster = eintraege[0]
        gemeinsam = lg.vereine([e.schluessel for e in eintraege])
        gemischt = gemeinsam is lg.VERSCHIEDEN

        if erster.art == rv.ART_BLOCK:
            if bl.editor_fuer(erster.bip_name) is not None:
                return self._blockknopf(erster)
            # Bei mehreren Vorlagen steht hier "Bearbeiten…": ob die
            # Einstellungen gleich sind, lässt sich nicht ablesen
            block = TextBlock()
            block.Text = erster.text
            block.Foreground = GRAU
            block.Margin = Thickness(4.0, 4.0, 4.0, 4.0)
            block.VerticalAlignment = VerticalAlignment.Center
            block.TextTrimming = TextTrimming.CharacterEllipsis
            block.ToolTip = t(
                u"Diese Einstellung gibt die Revit-API nicht zum Bearbeiten "
                u"frei – über \"Aus Vorlage übertragen…\" oder den nativen "
                u"Dialog ändern.",
                u"The Revit API does not expose this setting for editing – "
                u"change it via \"Transfer from template…\" or the native "
                u"dialog.",
                u"La API de Revit no permite editar este ajuste: use "
                u"\"Transferir desde plantilla…\" o el cuadro de diálogo "
                u"nativo.")
            return block

        if erster.art in (rv.ART_JANEIN, rv.ART_LISTE, rv.ART_ELEMENT):
            return self._auswahlfeld(eintraege, gemischt)
        return self._textfeld(eintraege, gemischt)

    def _auswahlfeld(self, eintraege, gemischt):
        erster = eintraege[0]
        box = ComboBox()
        box.Margin = Thickness(4.0, 2.0, 4.0, 2.0)
        box.Tag = erster.pid
        if gemischt:
            item = ComboBoxItem()
            item.Content = VERSCHIEDEN_TEXT
            item.Tag = None
            box.Items.Add(item)
        gewaehlt = None
        for text, wert in self.modell.optionen(erster, self.vorlagen):
            item = ComboBoxItem()
            item.Content = text
            item.Tag = wert
            box.Items.Add(item)
            if not gemischt and wert == erster.schluessel:
                gewaehlt = box.Items.Count - 1
        if gemischt:
            box.SelectedIndex = 0
        else:
            if gewaehlt is None:
                # Der aktuelle Wert steht in keiner Liste - lieber anzeigen
                # als stillschweigend den ersten Eintrag vorgaukeln
                item = ComboBoxItem()
                item.Content = erster.text
                item.Tag = erster.schluessel
                box.Items.Add(item)
                gewaehlt = box.Items.Count - 1
            box.SelectedIndex = gewaehlt
        box.SelectionChanged += self._h(self._bei_auswahlfeld)
        return box

    def _blockknopf(self, eintrag):
        """Knopf, der den Editor öffnet - beschriftet mit dem Stand."""
        knopf = Button()
        knopf.Content = eintrag.text
        knopf.Tag = eintrag.pid
        knopf.Margin = Thickness(4.0, 2.0, 4.0, 2.0)
        knopf.Padding = Thickness(6.0, 1.0, 6.0, 1.0)
        knopf.HorizontalContentAlignment = HorizontalAlignment.Left
        knopf.ToolTip = t(u"Öffnet den Editor – die Änderungen gehen an alle "
                          u"markierten Vorlagen.",
                          u"Opens the editor – the changes go to all selected "
                          u"templates.",
                          u"Abre el editor: los cambios van a todas las "
                          u"plantillas seleccionadas.")
        knopf.Click += self._h(self._bei_block)
        return knopf

    def _textfeld(self, eintraege, gemischt):
        erster = eintraege[0]
        feld = TextBox()
        feld.Margin = Thickness(4.0, 2.0, 4.0, 2.0)
        feld.Padding = Thickness(2.0, 1.0, 2.0, 1.0)
        feld.Text = VERSCHIEDEN_TEXT if gemischt else erster.text
        feld.Tag = (erster.pid, feld.Text)
        if erster.zusatz:
            feld.ToolTip = erster.zusatz
        if gemischt:
            feld.Foreground = GRAU
        feld.LostFocus += self._h(self._bei_textfeld)
        feld.KeyDown += self._h(self._bei_taste)
        return feld

    # ------------------------------------------------------------------
    # Änderungen schreiben
    # ------------------------------------------------------------------

    def _bei_haken(self, sender, args):
        if self._still:
            return
        pid = sender.Tag
        # Dritter Zustand ist nur die Anzeige gemischter Haken - beim
        # Klicken gilt an/aus
        if sender.IsChecked is None:
            sender.IsChecked = True
        an = bool(sender.IsChecked)
        sender.IsThreeState = False
        self._setze_haken([pid], an)

    def _alle_haken(self, an):
        pids = self._sichtbare_pids()
        if not pids:
            return
        anzahl_vorlagen = len(self._markierte())
        frage = (t(u"%d angezeigte Parameter in %d Vorlage(n) einschließen?",
                   u"Include %d shown parameters in %d template(s)?",
                   u"¿Incluir %d parámetros mostrados en %d plantilla(s)?")
                 if an else
                 t(u"%d angezeigte Parameter in %d Vorlage(n) nicht mehr "
                   u"einschließen? Die Ansichten steuern diese Werte dann "
                   u"wieder selbst.",
                   u"Stop including %d shown parameters in %d template(s)? "
                   u"The views then control these values again themselves.",
                   u"¿Dejar de incluir %d parámetros mostrados en %d "
                   u"plantilla(s)? Las vistas volverán a controlar esos "
                   u"valores."))
        if not dlg.frage(self.fenster, frage % (len(pids), anzahl_vorlagen),
                         titel=TITEL, warnung=not an):
            return
        self._setze_haken(pids, an)

    def _setze_haken(self, pids, an):
        markierte = self._markierte()
        name = (t(u"Parameter einschließen", u"Include parameters",
                  u"Incluir parámetros") if an
                else t(u"Parameter nicht einschließen",
                       u"Exclude parameters", u"No incluir parámetros"))

        def ausfuehren():
            for vorlage in markierte:
                self.modell.setze_enthalten(vorlage, pids, an)

        self._transaktion(name, ausfuehren)
        self.modell.vergiss()
        self._nach_aenderung()

    def _bei_block(self, sender, args):
        """Editor eines nicht lesbaren Blocks öffnen und Ergebnis schreiben."""
        eintraege = self.zeilen.get(sender.Tag)
        markierte = self._markierte()
        if not eintraege or not markierte:
            return
        editor = bl.editor_fuer(eintraege[0].bip_name)
        if editor is None:
            return
        aenderung = editor(self.fenster, self.doc, markierte)
        if aenderung is None:
            return

        def ausfuehren():
            for vorlage in markierte:
                aenderung.anwenden(vorlage)

        try:
            self._transaktion(t(u"%s ändern", u"Change %s", u"Cambiar %s")
                              % aenderung.beschreibung, ausfuehren)
        finally:
            self.modell.vergiss()
            self.suchtexte = {}
            self._nach_aenderung()
        dlg.meldung(self.fenster, t(
            u"%d Einstellung(en) in %d Vorlage(n) geschrieben.",
            u"%d setting(s) written in %d template(s).",
            u"%d ajuste(s) escrito(s) en %d plantilla(s).") % (
                aenderung.anzahl, len(markierte)), titel=TITEL)

    def _bei_auswahlfeld(self, sender, args):
        if self._still or sender.SelectedItem is None:
            return
        wert = sender.SelectedItem.Tag
        if wert is None:          # "<verschieden>" gewählt: nichts tun
            return
        self._schreibe(sender.Tag, wert)

    def _bei_taste(self, sender, args):
        if args.Key == Key.Enter:
            self._bei_textfeld(sender, args)

    def _bei_textfeld(self, sender, args):
        if self._still:
            return
        pid, urtext = sender.Tag
        if sender.Text == urtext:
            return
        # Merken, damit ein späteres LostFocus des inzwischen ersetzten
        # Feldes nicht ein zweites Mal schreibt
        sender.Tag = (pid, sender.Text)
        self._schreibe(pid, sender.Text)

    def _schreibe(self, pid, wert):
        eintraege = self.zeilen.get(pid)
        if not eintraege:
            return
        name = t(u"Vorlagenparameter ändern", u"Change template parameter",
                 u"Cambiar parámetro de plantilla")

        def ausfuehren():
            for eintrag in eintraege:
                self.modell.setze_wert(eintrag, wert)

        try:
            self._transaktion(name, ausfuehren)
        finally:
            # Auch nach einem Fehler neu einlesen, damit die Tabelle nicht
            # einen Wert zeigt, den Revit gar nicht angenommen hat
            self.modell.vergiss()
            self.suchtexte = {}
            self._nach_aenderung()

    def _nach_aenderung(self):
        """Tabelle und Liste nach einer Änderung auffrischen."""
        if self._neubau_geplant:
            return
        self._neubau_geplant = True

        def ausfuehren():
            self._neubau_geplant = False
            try:
                self.verwendung = self.modell.verwendung()
                self._zeige_auswahl()
            except Exception as fehler:
                dlg.zeige_fehler(self.fenster, fehler, FACHLICH, titel=TITEL)

        self.fenster.Dispatcher.InvokeAsync(Action(ausfuehren),
                                            DispatcherPriority.Background)

    # ------------------------------------------------------------------
    # Vorlagen verwalten
    # ------------------------------------------------------------------

    def _pruefe(self, name, alter_name=None):
        return lg.pruefe_name(self.modell.namen(), name, alter_name)

    def _neu(self, sender, args):
        ansichten = self.modell.vorlagenfaehige_ansichten()
        if not ansichten:
            raise rv.VorlagenFehler(
                t(u"Das Projekt enthält keine Ansicht, aus der sich eine "
                  u"Vorlage erzeugen lässt.",
                  u"The project has no view a template could be created from.",
                  u"El proyecto no tiene ninguna vista a partir de la cual "
                  u"crear una plantilla."))
        eintraege = [(u"%s: %s" % (rv.ansichtstyp(a), a.Name), id_wert(a.Id))
                     for a in ansichten]
        eintraege.sort(key=lambda e: lg.natuerlich(e[0]))
        gewaehlt = dlg.waehle(
            self.fenster, t(u"Neue Vorlage aus Ansicht",
                            u"New template from view",
                            u"Nueva plantilla desde vista"),
            t(u"Die Ansicht wählen, deren Einstellungen die neue Vorlage "
              u"übernehmen soll.",
              u"Choose the view whose settings the new template should take.",
              u"Elija la vista cuyos ajustes tomará la plantilla nueva."),
            eintraege)
        if gewaehlt is None:
            return
        ansicht = self.doc.GetElement(ElementId(gewaehlt))
        name = dlg.frage_text(
            self.fenster, t(u"Neue Vorlage", u"New template",
                            u"Plantilla nueva"),
            t(u"Name der neuen Vorlage:", u"Name of the new template:",
              u"Nombre de la plantilla nueva:"),
            lg.freier_name(self.modell.namen(), ansicht.Name),
            pruefen=self._pruefe)
        if name is None:
            return
        vorlage = self._transaktion(
            t(u"Ansichtsvorlage erstellen", u"Create view template",
              u"Crear plantilla de vista"),
            lambda: self.modell.aus_ansicht(ansicht, name))
        self.modell.vergiss()
        self._lade()
        self._waehle([id_wert(vorlage.Id)])

    def _duplizieren(self, sender, args):
        markierte = self._markierte()
        if not markierte:
            return
        if len(markierte) == 1:
            name = dlg.frage_text(
                self.fenster, t(u"Vorlage duplizieren", u"Duplicate template",
                                u"Duplicar plantilla"),
                t(u"Name der Kopie von \"%s\":",
                  u"Name of the copy of \"%s\":",
                  u"Nombre de la copia de \"%s\":") % markierte[0].Name,
                lg.freier_name(self.modell.namen(),
                               markierte[0].Name + t(u" - Kopie", u" - Copy",
                                                     u" - Copia")),
                pruefen=self._pruefe)
            if name is None:
                return
            namen = [name]
        else:
            if not dlg.frage(self.fenster, t(
                    u"%d Vorlagen duplizieren? Die Kopien erhalten den Zusatz "
                    u"\" - Kopie\".",
                    u"Duplicate %d templates? The copies get the suffix "
                    u"\" - Copy\".",
                    u"¿Duplicar %d plantillas? Las copias reciben el sufijo "
                    u"\" - Copia\".") % len(markierte), titel=TITEL):
                return
            namen = None

        def ausfuehren():
            kopien = []
            vergeben = list(self.modell.namen())
            for vorlage in markierte:
                name = namen[0] if namen else lg.freier_name(
                    vergeben, vorlage.Name + t(u" - Kopie", u" - Copy",
                                               u" - Copia"))
                vergeben.append(name)
                kopien.append(self.modell.dupliziere(vorlage, name))
            return kopien

        kopien = self._transaktion(t(u"Ansichtsvorlage duplizieren",
                                     u"Duplicate view template",
                                     u"Duplicar plantilla de vista"),
                                   ausfuehren)
        self.modell.vergiss()
        self._lade()
        self._waehle([id_wert(k.Id) for k in kopien])

    def _umbenennen(self, sender, args):
        markierte = self._markierte()
        if len(markierte) != 1:
            return
        vorlage = markierte[0]
        alt = vorlage.Name
        name = dlg.frage_text(
            self.fenster, t(u"Vorlage umbenennen", u"Rename template",
                            u"Renombrar plantilla"),
            t(u"Neuer Name für \"%s\":", u"New name for \"%s\":",
              u"Nombre nuevo para \"%s\":") % alt,
            alt, pruefen=lambda n: self._pruefe(n, alt))
        if name is None or name == alt:
            return

        def ausfuehren():
            vorlage.Name = name

        self._transaktion(t(u"Ansichtsvorlage umbenennen",
                            u"Rename view template",
                            u"Renombrar plantilla de vista"), ausfuehren)
        self._lade()
        self._fuelle_liste()
        self._zeige_auswahl()

    def _loeschen(self, sender, args):
        markierte = self._markierte()
        if not markierte:
            return
        betroffen = sum(self.verwendung.get(id_wert(v.Id), 0)
                        for v in markierte)
        namen = u"\n".join(u"• " + v.Name for v in markierte[:15])
        if len(markierte) > 15:
            namen += t(u"\n… und %d weitere", u"\n… and %d more",
                       u"\n… y %d más") % (len(markierte) - 15)
        frage = t(u"%d Ansichtsvorlage(n) löschen?\n\n%s",
                  u"Delete %d view template(s)?\n\n%s",
                  u"¿Eliminar %d plantilla(s) de vista?\n\n%s") % (
            len(markierte), namen)
        if betroffen:
            frage += t(
                u"\n\n%d Ansicht(en) verlieren dadurch ihre Vorlage; ihre "
                u"Einstellungen bleiben, sind danach aber frei änderbar.",
                u"\n\n%d view(s) will lose their template; their settings "
                u"remain but become editable again.",
                u"\n\n%d vista(s) perderán su plantilla; sus ajustes se "
                u"mantienen, pero vuelven a ser editables.") % betroffen
        if not dlg.frage(self.fenster, frage, titel=TITEL, warnung=True):
            return
        ids = [v.Id for v in markierte]

        def ausfuehren():
            self.doc.Delete(id_liste(ids))

        self._transaktion(t(u"Ansichtsvorlagen löschen",
                            u"Delete view templates",
                            u"Eliminar plantillas de vista"), ausfuehren)
        self.modell.vergiss()
        self._lade()
        self._waehle([])

    # ------------------------------------------------------------------
    # Übertragen
    # ------------------------------------------------------------------

    def _quellenliste(self, ausser):
        """[(Text, Id)] aller Vorlagen ausser den angegebenen."""
        gesperrt = set(id_wert(v.Id) for v in ausser)
        eintraege = [(u"%s: %s" % (rv.ansichtstyp(v), v.Name), id_wert(v.Id))
                     for v in self.vorlagen if id_wert(v.Id) not in gesperrt]
        eintraege.sort(key=lambda e: lg.natuerlich(e[0]))
        return eintraege

    def _uebertragen(self, sender, args):
        """Einstellungen einer anderen Vorlage in die markierten holen."""
        ziele = self._markierte()
        if not ziele:
            return
        quellen = self._quellenliste(ziele)
        if not quellen:
            raise rv.VorlagenFehler(
                t(u"Es gibt keine weitere Vorlage, aus der geholt werden "
                  u"könnte. Bitte die Quelle nicht mitmarkieren.",
                  u"There is no other template to get settings from. Please "
                  u"do not select the source as a target.",
                  u"No hay otra plantilla de la que traer ajustes. No "
                  u"seleccione la fuente como destino."))
        gewaehlt = dlg.waehle(
            self.fenster, t(u"Von anderer Vorlage holen",
                            u"Get from another template",
                            u"Traer de otra plantilla"),
            t(u"Quellvorlage wählen. Im nächsten Schritt wird abgehakt, was "
              u"davon in die %d markierte(n) Vorlage(n) wandert.",
              u"Choose the source template. In the next step you tick what "
              u"of it goes into the %d selected template(s).",
              u"Elija la plantilla de origen. En el paso siguiente se marca "
              u"qué pasa a las %d plantillas seleccionadas.") % len(ziele),
            quellen)
        if gewaehlt is None:
            return
        quelle = self.nach_id.get(gewaehlt)
        if quelle is not None:
            self._fuehre_uebertragung(quelle, ziele)

    def _kopieren(self, sender, args):
        """Einstellungen der markierten Vorlage auf andere kopieren."""
        markierte = self._markierte()
        if len(markierte) != 1:
            return
        quelle = markierte[0]
        moeglich = self._quellenliste([quelle])
        if not moeglich:
            raise rv.VorlagenFehler(
                t(u"Das Projekt hat keine zweite Ansichtsvorlage.",
                  u"The project has no second view template.",
                  u"El proyecto no tiene una segunda plantilla de vista."))
        gewaehlt = dlg.waehle(
            self.fenster, t(u"Auf andere Vorlagen kopieren",
                            u"Copy to other templates",
                            u"Copiar a otras plantillas"),
            t(u"Zielvorlagen wählen. Im nächsten Schritt wird abgehakt, was "
              u"von \"%s\" auf sie übertragen wird.",
              u"Choose the target templates. In the next step you tick what "
              u"gets transferred from \"%s\" to them.",
              u"Elija las plantillas de destino. En el paso siguiente se "
              u"marca qué se transfiere de \"%s\" a ellas.") % quelle.Name,
            moeglich, mehrfach=True)
        if not gewaehlt:
            return
        ziele = [self.nach_id[i] for i in gewaehlt if i in self.nach_id]
        if ziele:
            self._fuehre_uebertragung(quelle, ziele)

    def _fuehre_uebertragung(self, quelle, ziele):
        """Posten abhaken lassen und auf alle Zielvorlagen schreiben."""
        posten = ue.waehle_posten(self.fenster, self.modell, quelle, ziele)
        if not posten:
            return

        def ausfuehren():
            return ue.uebertrage(self.modell, quelle, ziele, posten)

        try:
            geschrieben, fehler = self._transaktion(
                t(u"Einstellungen übertragen", u"Transfer settings",
                  u"Transferir ajustes"), ausfuehren)
        finally:
            self.modell.vergiss()
            self.suchtexte = {}
        self._bericht(geschrieben, fehler, t(
            u"%d Einstellung(en) von \"%s\" auf %d Vorlage(n) übertragen.",
            u"%d setting(s) transferred from \"%s\" to %d template(s).",
            u"%d ajuste(s) transferido(s) de \"%s\" a %d plantilla(s).") % (
                geschrieben, quelle.Name, len(ziele)))
        self._zeige_auswahl()

    # ------------------------------------------------------------------
    # Ansichten
    # ------------------------------------------------------------------

    def _auf_ansichten(self, sender, args):
        markierte = self._markierte()
        if len(markierte) != 1:
            return
        vorlage = markierte[0]
        ansichten = self.modell.zuweisbare_ansichten(vorlage)
        if not ansichten:
            raise rv.VorlagenFehler(
                t(u"Keine Ansicht im Projekt nimmt die Vorlage \"%s\" an.",
                  u"No view in the project accepts the template \"%s\".",
                  u"Ninguna vista del proyecto admite la plantilla \"%s\".")
                % vorlage.Name)
        schon = set(id_wert(a.Id)
                    for a in self.modell.ansichten_mit(vorlage))
        eintraege = []
        for ansicht in ansichten:
            zusatz = u""
            try:
                laufende = id_wert(ansicht.ViewTemplateId)
                if laufende not in (None, -1) and laufende != id_wert(
                        vorlage.Id):
                    andere = self.nach_id.get(laufende)
                    if andere is not None:
                        zusatz = t(u"   [jetzt: %s]", u"   [currently: %s]",
                                   u"   [ahora: %s]") % andere.Name
            except Exception:
                pass
            eintraege.append((u"%s: %s%s" % (rv.ansichtstyp(ansicht),
                                             ansicht.Name, zusatz),
                              id_wert(ansicht.Id)))
        eintraege.sort(key=lambda e: lg.natuerlich(e[0]))
        gewaehlt = dlg.waehle(
            self.fenster, t(u"Vorlage auf Ansichten anwenden",
                            u"Apply template to views",
                            u"Aplicar plantilla a vistas"),
            t(u"Ansichten wählen, denen \"%s\" zugewiesen wird. Bereits "
              u"zugewiesene Ansichten sind vormarkiert; eine vorhandene "
              u"andere Vorlage wird ersetzt.",
              u"Choose the views to assign \"%s\" to. Views that already use "
              u"it are pre-checked; an existing other template is replaced.",
              u"Elija las vistas a las que se asignará \"%s\". Las vistas que "
              u"ya la usan están premarcadas; se sustituye otra plantilla "
              u"existente.") % vorlage.Name,
            eintraege, mehrfach=True, vorauswahl=schon)
        if not gewaehlt:
            return
        ziele = [self.doc.GetElement(ElementId(i)) for i in gewaehlt]
        erfolgreich, fehler = self._transaktion(
            t(u"Ansichtsvorlage zuweisen", u"Assign view template",
              u"Asignar plantilla de vista"),
            lambda: self.modell.weise_zu(ziele, vorlage))
        self.verwendung = self.modell.verwendung()
        self._fuelle_liste()
        self._bericht(erfolgreich, fehler, t(
            u"%d Ansicht(en) verwenden jetzt \"%s\".",
            u"%d view(s) now use \"%s\".",
            u"%d vista(s) usan ahora \"%s\".") % (erfolgreich, vorlage.Name))
        self._zeige_status()

    def _zuweisung_loesen(self, sender, args):
        markierte = self._markierte()
        if len(markierte) != 1:
            return
        vorlage = markierte[0]
        ansichten = self.modell.ansichten_mit(vorlage)
        if not ansichten:
            dlg.meldung(self.fenster, t(
                u"Keine Ansicht verwendet \"%s\".",
                u"No view uses \"%s\".",
                u"Ninguna vista usa \"%s\".") % vorlage.Name, titel=TITEL)
            return
        eintraege = [(u"%s: %s" % (rv.ansichtstyp(a), a.Name), id_wert(a.Id))
                     for a in ansichten]
        eintraege.sort(key=lambda e: lg.natuerlich(e[0]))
        gewaehlt = dlg.waehle(
            self.fenster, t(u"Zuweisung lösen", u"Remove assignment",
                            u"Quitar asignación"),
            t(u"Ansichten wählen, die \"%s\" nicht mehr verwenden sollen. Die "
              u"Einstellungen bleiben erhalten und sind danach frei änderbar.",
              u"Choose the views that should no longer use \"%s\". Their "
              u"settings remain and become editable again.",
              u"Elija las vistas que ya no deban usar \"%s\". Sus ajustes se "
              u"mantienen y vuelven a ser editables.") % vorlage.Name,
            eintraege, mehrfach=True,
            vorauswahl=[e[1] for e in eintraege])
        if not gewaehlt:
            return
        ziele = [self.doc.GetElement(ElementId(i)) for i in gewaehlt]
        erfolgreich, fehler = self._transaktion(
            t(u"Zuweisung der Ansichtsvorlage lösen",
              u"Remove view template assignment",
              u"Quitar asignación de plantilla de vista"),
            lambda: self.modell.loese_zuweisung(ziele))
        self.verwendung = self.modell.verwendung()
        self._fuelle_liste()
        self._bericht(erfolgreich, fehler, t(
            u"%d Ansicht(en) verwenden \"%s\" nicht mehr.",
            u"%d view(s) no longer use \"%s\".",
            u"%d vista(s) ya no usan \"%s\".") % (erfolgreich, vorlage.Name))
        self._zeige_status()

    def _bericht(self, erfolgreich, fehler, kopfzeile):
        text = kopfzeile
        if fehler:
            text += t(u"\n\nNicht möglich (%d):\n",
                      u"\n\nNot possible (%d):\n",
                      u"\n\nNo es posible (%d):\n") % len(fehler)
            text += u"\n".join(u"• " + u" / ".join(f) for f in fehler[:20])
            if len(fehler) > 20:
                text += t(u"\n… und %d weitere", u"\n… and %d more",
                          u"\n… y %d más") % (len(fehler) - 20)
        dlg.meldung(self.fenster, text, titel=TITEL, warnung=bool(fehler))

    # ------------------------------------------------------------------
    # Schliessen
    # ------------------------------------------------------------------

    def _ok(self, sender, args):
        self.uebernehmen = True
        self.fenster.Close()

    def _beim_schliessen(self, sender, args):
        if self.uebernehmen or not self.sitzung_geaendert:
            return
        if not dlg.frage(self.fenster, t(
                u"Alle Änderungen dieser Sitzung verwerfen?",
                u"Discard all changes of this session?",
                u"¿Descartar todos los cambios de esta sesión?"),
                titel=TITEL, warnung=True):
            args.Cancel = True


def starte(uiapp, doc):
    VorlagenFenster(uiapp, doc).zeige()
