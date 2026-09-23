# -*- coding: utf-8 -*-
"""Editoren für die Einstellungen, die nicht als Parameterwert lesbar sind.

Im nativen Dialog steht dort nur "Bearbeiten…": Sichtbarkeit und
Grafiküberschreibungen der Kategorien, die Filter der Vorlage, die
Bearbeitungsbereiche und der Ansichtsbereich.

Jeder Editor arbeitet nach demselben Muster:

    1. Er zeigt den Stand der ersten markierten Vorlage.
    2. Er merkt sich nur die *Änderungen*, nicht den ganzen Zustand.
    3. Er gibt eine Aenderung zurück, die das Fenster in einer Transaktion
       auf alle markierten Vorlagen anwendet.

Dadurch bleibt beim Bearbeiten mehrerer Vorlagen alles unangetastet, was
nicht angefasst wurde: Wer nur "Halbton" für Wände setzt, überschreibt in
den anderen Vorlagen nicht deren Linienfarben.

Die Grafiküberschreibungen selbst kommen aus filter_manager.darstellung -
derselbe Dialog wie beim Filter-Manager, mit "unverändert" als Vorgabe je
Einstellung. Mehrere Überschreibungen nacheinander werden gesammelt und in
der Reihenfolge angewendet, in der sie gesetzt wurden.
"""

from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    FilteredWorksetCollector,
    OverrideGraphicSettings,
    ParameterFilterElement,
    PlanViewPlane,
    ViewPlan,
    WorksetId,
    WorksetKind,
    WorksetVisibility,
)
from System.Windows import (
    GridLength,
    GridUnitType,
    TextTrimming,
    Thickness,
    VerticalAlignment,
    Visibility,
)
from System.Windows.Controls import (
    Button,
    CheckBox,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    ListBoxItem,
    TextBlock,
    TextBox,
)
from System.Windows.Media import Color, SolidColorBrush

from filter_manager import darstellung as ds
from filter_manager import dialoge as dlg
from filter_manager.aufloesung import id_wert
from mlg_sprache import t, uebersetze_xaml
from vorlagen_manager import logik as lg
from vorlagen_manager.revit import KATEGORIE_BLOECKE, VorlagenFehler

UNVERAENDERT = t(u"unverändert", u"unchanged", u"sin cambios")

# Überschriften der Kategorieblöcke (Typen stehen in revit.KATEGORIE_BLOECKE)
UEBERSCHRIFTEN = {
    "VIS_GRAPHICS_MODEL": t(u"Modellkategorien", u"Model categories",
                            u"Categorías de modelo"),
    "VIS_GRAPHICS_ANNOTATION": t(u"Beschriftungskategorien",
                                 u"Annotation categories",
                                 u"Categorías de anotación"),
    "VIS_GRAPHICS_ANALYTICAL_MODEL": t(u"Analytische Kategorien",
                                       u"Analytical categories",
                                       u"Categorías analíticas"),
}


def _pinsel(r, g, b):
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
BLAU = _pinsel(20, 90, 170)


class Aenderung(object):
    """Was ein Editor geändert hat - anwendbar auf beliebig viele Vorlagen."""

    def __init__(self, beschreibung, funktion, anzahl):
        self.beschreibung = beschreibung
        self._funktion = funktion
        self.anzahl = anzahl          # Zahl der geänderten Einträge

    def anwenden(self, vorlage):
        self._funktion(vorlage)


# ---------------------------------------------------------------------------
# Gemeinsames Fenster: Hinweis, Suche, Liste, Knopfleiste
# ---------------------------------------------------------------------------

_XAML_TEXTE = {
    "t0": (u"Suchen:", u"Search:", u"Buscar:"),
    "t1": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

_XAML = u"""
<Window %s Title="{titel}" Width="880" Height="660" MinWidth="620"
        MinHeight="420" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,3"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
    </Style>
  </Window.Resources>
  <DockPanel Margin="12">
    <TextBlock x:Name="hinweis" DockPanel.Dock="Top" TextWrapping="Wrap"
               Margin="0,0,0,8"/>
    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
      <TextBlock Text="{{t0}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <StackPanel x:Name="schalter" DockPanel.Dock="Right"
                  Orientation="Horizontal" Margin="8,0,0,0"/>
      <TextBox x:Name="suche" Padding="3"/>
    </DockPanel>
    <DockPanel DockPanel.Dock="Bottom" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="OK" Width="90"/>
        <Button Content="{{t1}}" Width="90" IsCancel="True" Margin="0"/>
      </StackPanel>
      <StackPanel x:Name="knoepfe" Orientation="Horizontal"/>
    </DockPanel>
    <TextBlock x:Name="zaehler" DockPanel.Dock="Bottom" Foreground="#666"
               Margin="0,6,0,0"/>
    <ListBox x:Name="liste" SelectionMode="Extended"
             HorizontalContentAlignment="Stretch"/>
  </DockPanel>
</Window>"""


def _escape(text):
    return (text.replace(u"&", u"&amp;").replace(u"<", u"&lt;")
            .replace(u">", u"&gt;").replace(u'"', u"&quot;"))


class Zeile(object):
    """Eine Listenzeile mit eigenem Änderungszustand."""

    def __init__(self, item, suchtext, schluessel):
        self.item = item
        self.suchtext = suchtext
        self.schluessel = schluessel
        self.geaendert = False
        self.haken = None            # Kontrollkästchen der Zeile
        self.anzeige = None          # TextBlock für die Überschreibungen
        self.grundtext = u"–"        # Stand der ersten Vorlage
        self.vergleich = None        # TextBlock für den Zielvergleich
        self.abweichend = None       # True/False/None (noch nicht verglichen)


class Blockfenster(object):
    """Liste mit Suche, Mehrfachauswahl und frei bestückbarer Knopfleiste."""

    def __init__(self, besitzer, titel, hinweis):
        self.titel = titel
        self.fenster = dlg.lade_xaml(uebersetze_xaml(
            (_XAML % dlg.XMLNS).replace(u"{titel}", _escape(titel)),
            _XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, besitzer=besitzer)
        self.c = self.fenster.FindName
        self.c("hinweis").Text = hinweis
        self.zeilen = []
        self.bestaetigt = False
        self.zaehlwort = t(u"geändert", u"changed", u"cambiados")
        self._zusatzfilter = None
        self.c("suche").TextChanged += self._h(lambda s, a: self.filtern())
        self.c("ok").Click += self._h(self._bei_ok)

    def _h(self, funktion):
        return dlg.sicher(lambda: self.fenster, funktion, (VorlagenFehler,))

    # -- Aufbau ---------------------------------------------------------

    def zeile(self, gitter, suchtext, schluessel):
        """Eine Zeile anhängen. Der Index steckt im Tag des ListBoxItem."""
        item = ListBoxItem()
        item.Content = gitter
        item.Tag = len(self.zeilen)
        self.c("liste").Items.Add(item)
        neue = Zeile(item, suchtext.lower(), schluessel)
        self.zeilen.append(neue)
        return neue

    def knopf(self, text, funktion, tooltip=None):
        knopf = Button()
        knopf.Content = text
        if tooltip:
            knopf.ToolTip = tooltip
        knopf.Click += self._h(funktion)
        self.c("knoepfe").Children.Add(knopf)
        return knopf

    def schalter(self, text, funktion):
        kasten = CheckBox()
        kasten.Content = text
        kasten.VerticalAlignment = VerticalAlignment.Center
        kasten.Margin = Thickness(8.0, 0.0, 0.0, 0.0)
        kasten.Click += self._h(funktion)
        self.c("schalter").Children.Add(kasten)
        return kasten

    # -- Auswahl und Filter ---------------------------------------------

    def markierte(self):
        """Die markierten Zeile-Objekte (über den Index im Tag)."""
        return [self.zeilen[item.Tag]
                for item in self.c("liste").SelectedItems]

    def sichtbare(self):
        return [z for z in self.zeilen
                if z.item.Visibility == Visibility.Visible]

    def setze_zusatzfilter(self, funktion):
        self._zusatzfilter = funktion
        self.filtern()

    def filtern(self):
        woerter = self.c("suche").Text.lower().split()
        for zeile in self.zeilen:
            passt = lg.passt(zeile.suchtext, woerter) and (
                self._zusatzfilter is None or self._zusatzfilter(zeile))
            zeile.item.Visibility = (Visibility.Visible if passt
                                     else Visibility.Collapsed)
        self.zaehle()

    def zaehle(self):
        sichtbar = len(self.sichtbare())
        geaendert = sum(1 for z in self.zeilen if z.geaendert)
        text = t(u"%d von %d angezeigt", u"%d of %d shown",
                 u"%d de %d mostrados") % (sichtbar, len(self.zeilen))
        if geaendert:
            text += u" – %d %s" % (geaendert, self.zaehlwort)
        self.c("zaehler").Text = text

    def hinweis_markieren(self):
        dlg.meldung(self.fenster, t(
            u"Bitte zuerst eine oder mehrere Zeilen markieren.",
            u"Please select one or more rows first.",
            u"Seleccione primero una o varias filas."), titel=self.titel)

    # -- Ablauf ---------------------------------------------------------

    def _bei_ok(self, sender, args):
        self.bestaetigt = True
        self.fenster.DialogResult = True

    def zeige(self):
        self.fenster.Loaded += lambda s, a: self.c("suche").Focus()
        self.zaehle()
        self.fenster.ShowDialog()
        return self.bestaetigt


# ---------------------------------------------------------------------------
# Bausteine für Zeilen
# ---------------------------------------------------------------------------

def _gitter(*breiten):
    """Grid mit den angegebenen Spaltenbreiten ("auto" oder Sternfaktor)."""
    gitter = Grid()
    for breite in breiten:
        spalte = ColumnDefinition()
        spalte.Width = (GridLength.Auto if breite == "auto"
                        else GridLength(float(breite), GridUnitType.Star))
        gitter.ColumnDefinitions.Add(spalte)
    return gitter


def _text(inhalt, spalte, grau=False, einzug=0.0):
    block = TextBlock()
    block.Text = inhalt
    block.VerticalAlignment = VerticalAlignment.Center
    block.Margin = Thickness(4.0 + einzug, 3.0, 4.0, 3.0)
    block.TextTrimming = TextTrimming.CharacterEllipsis
    block.ToolTip = inhalt
    if grau:
        block.Foreground = GRAU
    Grid.SetColumn(block, spalte)
    return block


def _kaestchen(spalte, beschriftung=None):
    kasten = CheckBox()
    kasten.VerticalAlignment = VerticalAlignment.Center
    kasten.Margin = Thickness(4.0, 3.0, 8.0, 3.0)
    if beschriftung:
        kasten.Content = beschriftung
    Grid.SetColumn(kasten, spalte)
    return kasten


def _hinweis(vorlagen, text):
    if len(vorlagen) == 1:
        return text
    return t(u"Die Änderungen gehen an alle %d markierten Vorlagen.\n",
             u"The changes are applied to all %d selected templates.\n",
             u"Los cambios se aplican a las %d plantillas seleccionadas.\n") \
        % len(vorlagen) + text


def ogs_kurz(doc, ogs):
    """Knappe Beschreibung einer Grafiküberschreibung."""
    if ogs is None:
        return u"–"
    teile = []
    try:
        if ogs.Halftone:
            teile.append(t(u"Halbton", u"halftone", u"medio tono"))
        if ogs.Transparency:
            teile.append(t(u"Transparenz %d%%", u"transparency %d%%",
                           u"transparencia %d%%") % ogs.Transparency)
        if ogs.ProjectionLineColor is not None and \
                ogs.ProjectionLineColor.IsValid:
            teile.append(t(u"Linienfarbe", u"line colour", u"color de línea"))
        if ogs.ProjectionLineWeight > 0:
            teile.append(t(u"Linienstärke %d", u"line weight %d",
                           u"grosor de línea %d") % ogs.ProjectionLineWeight)
        if ogs.CutLineColor is not None and ogs.CutLineColor.IsValid:
            teile.append(t(u"Schnittfarbe", u"cut colour", u"color de corte"))
        if ogs.CutLineWeight > 0:
            teile.append(t(u"Schnittstärke %d", u"cut weight %d",
                           u"grosor de corte %d") % ogs.CutLineWeight)
        for muster, name in (
                (ogs.SurfaceForegroundPatternId,
                 t(u"Oberflächenmuster", u"surface pattern",
                   u"patrón de superficie")),
                (ogs.CutForegroundPatternId,
                 t(u"Schnittmuster", u"cut pattern", u"patrón de corte"))):
            if id_wert(muster) not in (None, -1):
                teile.append(name)
    except Exception:
        return u"?"
    return u", ".join(teile) if teile else u"–"


def _geplant_text(zeile, darstellungen):
    """Was aus der Überschreibungsspalte nach dem Anwenden wird."""
    if not darstellungen:
        return zeile.grundtext
    return u"» " + (darstellungen[-1].zusammenfassung() or UNVERAENDERT)


def _frage_ueberschreibungen(fenster, doc, anzahl, vorbelegung):
    return ds.frage_darstellung(
        fenster.fenster, doc,
        t(u"Grafiküberschreibungen", u"Graphic overrides",
          u"Modificaciones gráficas"),
        t(u"Überschreibungen für %d Zeile(n). Was auf \"unverändert\" steht, "
          u"bleibt in jeder Vorlage wie es ist.",
          u"Overrides for %d row(s). Anything left on \"unchanged\" stays as "
          u"it is in every template.",
          u"Modificaciones para %d fila(s). Lo que quede en \"sin cambios\" "
          u"se mantiene en cada plantilla.") % anzahl, vorbelegung)


# ---------------------------------------------------------------------------
# Kategorien (V/G)
# ---------------------------------------------------------------------------

def _steuerbar(kategorie, vorlage, streng):
    try:
        if not streng:
            return True
        return kategorie.get_AllowsVisibilityControl(vorlage)
    except Exception:
        return not streng


def _sammle_kategorien(doc, vorlage, typ, streng):
    oberste = []
    for kategorie in doc.Settings.Categories:
        try:
            if kategorie.CategoryType != typ or not kategorie.IsVisibleInUI:
                continue
        except Exception:
            continue
        if not _steuerbar(kategorie, vorlage, streng):
            continue
        unter = []
        try:
            for kind in kategorie.SubCategories:
                if _steuerbar(kind, vorlage, streng):
                    unter.append((u"%s" % kind.Name, kind, 20.0))
        except Exception:
            pass
        unter.sort(key=lambda e: lg.natuerlich(e[0]))
        oberste.append((u"%s" % kategorie.Name, kategorie, 0.0, unter))
    oberste.sort(key=lambda e: lg.natuerlich(e[0]))
    ergebnis = []
    for name, kategorie, einzug, unter in oberste:
        ergebnis.append((name, kategorie, einzug))
        ergebnis.extend(unter)
    return ergebnis


def kategorien(doc, vorlage, typ):
    """[(Name, Category, Einzug)] der steuerbaren Kategorien eines Typs.

    AllowsVisibilityControl() antwortet nicht bei jeder Vorlagenart
    brauchbar. Kommt damit nichts heraus, wird ohne diese Prüfung gesammelt
    - lieber eine Zeile zu viel als ein leerer Dialog.
    """
    gefunden = _sammle_kategorien(doc, vorlage, typ, True)
    return gefunden or _sammle_kategorien(doc, vorlage, typ, False)


def kategorien_bearbeiten(besitzer, doc, vorlagen, bip_name):
    """V/G-Kategorien bearbeiten. Rückgabe: Aenderung oder None."""
    typ = KATEGORIE_BLOECKE[bip_name]
    ueberschrift = UEBERSCHRIFTEN[bip_name]
    erste = vorlagen[0]
    fenster = Blockfenster(
        besitzer,
        t(u"%s (Sichtbarkeit/Grafiken)", u"%s (visibility/graphics)",
          u"%s (visibilidad/gráficos)") % ueberschrift,
        _hinweis(vorlagen, t(
            u"Haken = sichtbar. Für Grafiküberschreibungen eine oder mehrere "
            u"Zeilen markieren und \"Überschreibungen…\" wählen. Angezeigt "
            u"wird der Stand von \"%s\"; geschrieben wird nur, was hier "
            u"geändert wird.",
            u"Check = visible. For graphic overrides select one or more rows "
            u"and choose \"Overrides…\". Shown is the state of \"%s\"; only "
            u"what is changed here gets written.",
            u"Marca = visible. Para modificaciones gráficas seleccione una o "
            u"varias filas y elija \"Modificaciones…\". Se muestra el estado "
            u"de \"%s\"; solo se escribe lo que se cambie aquí.")
            % erste.Name))

    aenderungen = {}   # Kategorie-Id -> {"sichtbar", "darstellungen"}

    def zustand(zeile):
        return aenderungen.setdefault(
            zeile.schluessel, {"sichtbar": None, "darstellungen": []})

    def auffrischen(zeile):
        eintrag = aenderungen.get(zeile.schluessel) or {}
        zeile.geaendert = bool(eintrag.get("sichtbar") is not None
                               or eintrag.get("darstellungen"))
        zeile.anzeige.Text = _geplant_text(zeile,
                                           eintrag.get("darstellungen"))
        zeile.anzeige.Foreground = BLAU if zeile.geaendert else GRAU
        fenster.zaehle()

    for name, kategorie, einzug in kategorien(doc, erste, typ):
        gitter = _gitter("auto", "1.3", "1")
        haken = _kaestchen(0)
        try:
            haken.IsChecked = not erste.GetCategoryHidden(kategorie.Id)
        except Exception:
            haken.IsChecked = True
        try:
            haken.IsEnabled = erste.CanCategoryBeHidden(kategorie.Id)
        except Exception:
            pass
        gitter.Children.Add(haken)
        gitter.Children.Add(_text(name, 1, einzug=einzug))
        try:
            kurz = ogs_kurz(doc, erste.GetCategoryOverrides(kategorie.Id))
        except Exception:
            kurz = u"–"
        anzeige = _text(kurz, 2, grau=True)
        gitter.Children.Add(anzeige)

        zeile = fenster.zeile(gitter, name, id_wert(kategorie.Id))
        zeile.haken = haken
        zeile.anzeige = anzeige
        zeile.grundtext = kurz

        def bei_haken(sender, args, zeile=zeile):
            zustand(zeile)["sichtbar"] = bool(sender.IsChecked)
            auffrischen(zeile)

        haken.Click += fenster._h(bei_haken)

    if fenster.zeilen and not any(z.haken.IsEnabled for z in fenster.zeilen):
        # CanCategoryBeHidden() war für keine Zeile zu gebrauchen - dann
        # lieber alle freigeben; ein untauglicher Aufruf wird beim
        # Schreiben ohnehin abgefangen
        for zeile in fenster.zeilen:
            zeile.haken.IsEnabled = True

    nur_geaendert = fenster.schalter(
        t(u"nur geänderte", u"changed only", u"solo cambiados"), None)
    nur_geaendert.Click += fenster._h(
        lambda s, a: fenster.setze_zusatzfilter(
            (lambda z: z.geaendert) if nur_geaendert.IsChecked else None))

    def setze_sichtbar(an):
        ziel = fenster.markierte() or fenster.sichtbare()
        for zeile in ziel:
            if not zeile.haken.IsEnabled:
                continue
            zeile.haken.IsChecked = an
            zustand(zeile)["sichtbar"] = an
            auffrischen(zeile)

    def bei_ueberschreibungen(sender, args):
        ziel = fenster.markierte()
        if not ziel:
            fenster.hinweis_markieren()
            return
        vorbelegung = None
        if len(ziel) == 1:
            try:
                vorbelegung = erste.GetCategoryOverrides(
                    ElementId(ziel[0].schluessel))
            except Exception:
                vorbelegung = None
        darstellung = _frage_ueberschreibungen(fenster, doc, len(ziel),
                                               vorbelegung)
        if darstellung is None:
            return
        for zeile in ziel:
            eintrag = zustand(zeile)
            if darstellung.sichtbarkeit() is not None:
                eintrag["sichtbar"] = darstellung.sichtbarkeit()
                if zeile.haken.IsEnabled:
                    zeile.haken.IsChecked = darstellung.sichtbarkeit()
            if darstellung.aendert_grafik():
                eintrag["darstellungen"].append(darstellung)
            auffrischen(zeile)

    def bei_zuruecksetzen(sender, args):
        ziel = fenster.markierte()
        if not ziel:
            fenster.hinweis_markieren()
            return
        leer = ds.Darstellung()
        leer.zuruecksetzen = True
        for zeile in ziel:
            zustand(zeile)["darstellungen"].append(leer)
            auffrischen(zeile)

    fenster.knopf(t(u"Sichtbar", u"Visible", u"Visible"),
                  lambda s, a: setze_sichtbar(True))
    fenster.knopf(t(u"Ausblenden", u"Hide", u"Ocultar"),
                  lambda s, a: setze_sichtbar(False))
    fenster.knopf(t(u"Überschreibungen…", u"Overrides…", u"Modificaciones…"),
                  bei_ueberschreibungen)
    fenster.knopf(t(u"Zurücksetzen", u"Reset", u"Restablecer"),
                  bei_zuruecksetzen)

    if not fenster.zeige():
        return None
    echte = dict((k, v) for k, v in aenderungen.items()
                 if v["sichtbar"] is not None or v["darstellungen"])
    if not echte:
        return None

    def anwenden(vorlage):
        for kid, eintrag in echte.items():
            kategorie_id = ElementId(kid)
            if eintrag["sichtbar"] is not None:
                try:
                    vorlage.SetCategoryHidden(kategorie_id,
                                              not eintrag["sichtbar"])
                except Exception:
                    pass
            for darstellung in eintrag["darstellungen"]:
                try:
                    ogs = (OverrideGraphicSettings()
                           if darstellung.zuruecksetzen
                           else OverrideGraphicSettings(
                               vorlage.GetCategoryOverrides(kategorie_id)))
                    darstellung.uebertrage(ogs, doc)
                    vorlage.SetCategoryOverrides(kategorie_id, ogs)
                except Exception:
                    pass

    return Aenderung(ueberschrift, anwenden, len(echte))


# ---------------------------------------------------------------------------
# Filter (V/G)
# ---------------------------------------------------------------------------

def filter_bearbeiten(besitzer, doc, vorlagen):
    """Filter der Vorlage bearbeiten. Rückgabe: Aenderung oder None."""
    erste = vorlagen[0]
    fenster = Blockfenster(
        besitzer,
        t(u"Filter (Sichtbarkeit/Grafiken)", u"Filters (visibility/graphics)",
          u"Filtros (visibilidad/gráficos)"),
        _hinweis(vorlagen, t(
            u"Erster Haken = Filter ist in der Vorlage angewendet, zweiter = "
            u"die gefilterten Elemente sind sichtbar. Angezeigt wird der "
            u"Stand von \"%s\".",
            u"First check = filter is applied in the template, second = the "
            u"filtered elements are visible. Shown is the state of \"%s\".",
            u"Primera marca = el filtro está aplicado en la plantilla, "
            u"segunda = los elementos filtrados son visibles. Se muestra el "
            u"estado de \"%s\".") % erste.Name))

    try:
        angewendet = set(id_wert(i) for i in erste.GetFilters())
    except Exception:
        angewendet = set()
    aenderungen = {}   # Filter-Id -> {"an", "sichtbar", "darstellungen"}

    def zustand(zeile):
        return aenderungen.setdefault(
            zeile.schluessel,
            {"an": None, "sichtbar": None, "darstellungen": []})

    def auffrischen(zeile):
        eintrag = aenderungen.get(zeile.schluessel) or {}
        zeile.geaendert = bool(eintrag.get("an") is not None
                               or eintrag.get("sichtbar") is not None
                               or eintrag.get("darstellungen"))
        zeile.anzeige.Text = _geplant_text(zeile,
                                           eintrag.get("darstellungen"))
        zeile.anzeige.Foreground = BLAU if zeile.geaendert else GRAU
        fenster.zaehle()

    elemente = list(FilteredElementCollector(doc).OfClass(
        ParameterFilterElement))
    elemente.sort(key=lambda f: lg.natuerlich(f.Name))
    for element in elemente:
        fid = id_wert(element.Id)
        gitter = _gitter("auto", "auto", "1.3", "1")
        haken = _kaestchen(0)
        haken.IsChecked = fid in angewendet
        haken.ToolTip = t(u"Filter in der Vorlage anwenden",
                          u"Apply filter in the template",
                          u"Aplicar el filtro en la plantilla")
        gitter.Children.Add(haken)
        sicht = _kaestchen(1, t(u"sichtbar", u"visible", u"visible"))
        try:
            sicht.IsChecked = (erste.GetFilterVisibility(element.Id)
                               if fid in angewendet else True)
        except Exception:
            sicht.IsChecked = True
        gitter.Children.Add(sicht)
        gitter.Children.Add(_text(u"%s" % element.Name, 2))
        try:
            kurz = (ogs_kurz(doc, erste.GetFilterOverrides(element.Id))
                    if fid in angewendet else u"–")
        except Exception:
            kurz = u"–"
        anzeige = _text(kurz, 3, grau=True)
        gitter.Children.Add(anzeige)

        zeile = fenster.zeile(gitter, element.Name, fid)
        zeile.haken = haken
        zeile.anzeige = anzeige
        zeile.grundtext = kurz

        def bei_an(sender, args, zeile=zeile):
            zustand(zeile)["an"] = bool(sender.IsChecked)
            auffrischen(zeile)

        def bei_sicht(sender, args, zeile=zeile):
            zustand(zeile)["sichtbar"] = bool(sender.IsChecked)
            auffrischen(zeile)

        haken.Click += fenster._h(bei_an)
        sicht.Click += fenster._h(bei_sicht)

    def bei_ueberschreibungen(sender, args):
        ziel = fenster.markierte()
        if not ziel:
            fenster.hinweis_markieren()
            return
        vorbelegung = None
        if len(ziel) == 1 and ziel[0].schluessel in angewendet:
            try:
                vorbelegung = erste.GetFilterOverrides(
                    ElementId(ziel[0].schluessel))
            except Exception:
                vorbelegung = None
        darstellung = _frage_ueberschreibungen(fenster, doc, len(ziel),
                                               vorbelegung)
        if darstellung is None:
            return
        for zeile in ziel:
            eintrag = zustand(zeile)
            if darstellung.sichtbarkeit() is not None:
                eintrag["sichtbar"] = darstellung.sichtbarkeit()
                zeile.haken.IsChecked = True
                eintrag["an"] = True
            if darstellung.aendert_grafik():
                eintrag["darstellungen"].append(darstellung)
                zeile.haken.IsChecked = True
                eintrag["an"] = True
            auffrischen(zeile)

    fenster.knopf(t(u"Überschreibungen…", u"Overrides…", u"Modificaciones…"),
                  bei_ueberschreibungen)

    if not fenster.zeige():
        return None
    echte = dict((k, v) for k, v in aenderungen.items()
                 if v["an"] is not None or v["sichtbar"] is not None
                 or v["darstellungen"])
    if not echte:
        return None

    def anwenden(vorlage):
        for fid, eintrag in echte.items():
            filter_id = ElementId(fid)
            try:
                schon = vorlage.IsFilterApplied(filter_id)
                if eintrag["an"] is False:
                    if schon:
                        vorlage.RemoveFilter(filter_id)
                    continue
                if not schon:
                    vorlage.AddFilter(filter_id)
                if eintrag["sichtbar"] is not None:
                    vorlage.SetFilterVisibility(filter_id,
                                                eintrag["sichtbar"])
                for darstellung in eintrag["darstellungen"]:
                    ogs = (OverrideGraphicSettings()
                           if darstellung.zuruecksetzen
                           else OverrideGraphicSettings(
                               vorlage.GetFilterOverrides(filter_id)))
                    darstellung.uebertrage(ogs, doc)
                    vorlage.SetFilterOverrides(filter_id, ogs)
            except Exception:
                pass

    return Aenderung(t(u"Filter", u"Filters", u"Filtros"), anwenden,
                     len(echte))


# ---------------------------------------------------------------------------
# Bearbeitungsbereiche
# ---------------------------------------------------------------------------

WORKSET_WAHL = [
    (WorksetVisibility.Visible, t(u"Sichtbar", u"Visible", u"Visible")),
    (WorksetVisibility.Hidden, t(u"Ausgeblendet", u"Hidden", u"Oculto")),
    (WorksetVisibility.UseGlobalSetting,
     t(u"Globale Einstellung verwenden", u"Use global setting",
       u"Usar valor global")),
]


def worksets_bearbeiten(besitzer, doc, vorlagen):
    """Sichtbarkeit der Bearbeitungsbereiche. Rückgabe: Aenderung oder None."""
    if not doc.IsWorkshared:
        raise VorlagenFehler(t(
            u"Das Projekt ist nicht für die Zusammenarbeit freigegeben - es "
            u"gibt keine Bearbeitungsbereiche.",
            u"The project is not workshared - there are no worksets.",
            u"El proyecto no está compartido: no hay subproyectos."))
    erste = vorlagen[0]
    fenster = Blockfenster(
        besitzer, t(u"Bearbeitungsbereiche", u"Worksets", u"Subproyectos"),
        _hinweis(vorlagen, t(
            u"Sichtbarkeit je Bearbeitungsbereich. Angezeigt wird der Stand "
            u"von \"%s\"; nur geänderte Zeilen werden geschrieben.",
            u"Visibility per workset. Shown is the state of \"%s\"; only "
            u"changed rows get written.",
            u"Visibilidad por subproyecto. Se muestra el estado de \"%s\"; "
            u"solo se escriben las filas cambiadas.") % erste.Name))

    aenderungen = {}   # Workset-Id (int) -> WorksetVisibility

    worksets = list(FilteredWorksetCollector(doc).OfKind(
        WorksetKind.UserWorkset))
    worksets.sort(key=lambda w: lg.natuerlich(w.Name))
    for workset in worksets:
        gitter = _gitter("1.4", "auto")
        gitter.Children.Add(_text(u"%s" % workset.Name, 0))
        box = ComboBox()
        box.Width = 240.0
        box.Margin = Thickness(4.0, 2.0, 4.0, 2.0)
        try:
            aktuell = erste.GetWorksetVisibility(workset.Id)
        except Exception:
            aktuell = None
        for nummer, (wert, beschriftung) in enumerate(WORKSET_WAHL):
            eintrag = ComboBoxItem()
            eintrag.Content = beschriftung
            eintrag.Tag = nummer
            box.Items.Add(eintrag)
            if aktuell is not None and wert == aktuell:
                box.SelectedIndex = nummer
        if box.SelectedIndex < 0:
            box.SelectedIndex = 0
        Grid.SetColumn(box, 1)
        gitter.Children.Add(box)

        zeile = fenster.zeile(gitter, workset.Name,
                              workset.Id.IntegerValue)

        def bei_wahl(sender, args, zeile=zeile):
            if sender.SelectedItem is None:
                return
            aenderungen[zeile.schluessel] = \
                WORKSET_WAHL[sender.SelectedItem.Tag][0]
            zeile.geaendert = True
            fenster.zaehle()

        box.SelectionChanged += fenster._h(bei_wahl)

    if not fenster.zeige() or not aenderungen:
        return None

    def anwenden(vorlage):
        for nummer, sichtbarkeit in aenderungen.items():
            try:
                vorlage.SetWorksetVisibility(WorksetId(nummer), sichtbarkeit)
            except Exception:
                pass

    return Aenderung(t(u"Bearbeitungsbereiche", u"Worksets", u"Subproyectos"),
                     anwenden, len(aenderungen))


# ---------------------------------------------------------------------------
# Ansichtsbereich
# ---------------------------------------------------------------------------

EBENEN = [
    (PlanViewPlane.TopClipPlane, t(u"Obere Begrenzung", u"Top", u"Superior")),
    (PlanViewPlane.CutPlane, t(u"Schnittebene", u"Cut plane",
                               u"Plano de corte")),
    (PlanViewPlane.BottomClipPlane, t(u"Untere Begrenzung", u"Bottom",
                                      u"Inferior")),
    (PlanViewPlane.ViewDepthPlane, t(u"Ansichtstiefe", u"View depth",
                                     u"Profundidad de vista")),
]

# Interne Einheit ist Fuss; eingegeben wird in Millimetern
MM_JE_FUSS = 304.8


def ansichtsbereich_bearbeiten(besitzer, doc, vorlagen):
    """Ansichtsbereich von Grundrissvorlagen. Rückgabe: Aenderung oder None."""
    grundrisse = [v for v in vorlagen if isinstance(v, ViewPlan)]
    if not grundrisse:
        raise VorlagenFehler(t(
            u"Den Ansichtsbereich gibt es nur bei Grundriss- und "
            u"Deckenplanvorlagen.",
            u"The view range only exists for plan view templates.",
            u"El rango de vista solo existe en plantillas de plano."))
    erste = grundrisse[0]
    bereich = erste.GetViewRange()
    fenster = Blockfenster(
        besitzer, t(u"Ansichtsbereich", u"View range", u"Rango de vista"),
        _hinweis(grundrisse, t(
            u"Versatz in Millimetern gegenüber der zugehörigen Ebene. Die "
            u"Ebenen selbst bleiben, wie sie in jeder Vorlage eingestellt "
            u"sind. Angezeigt wird der Stand von \"%s\"; unveränderte Felder "
            u"werden nicht geschrieben.",
            u"Offset in millimetres from the associated level. The levels "
            u"themselves stay as set in each template. Shown is the state of "
            u"\"%s\"; untouched fields are not written.",
            u"Desfase en milímetros respecto al nivel asociado. Los niveles "
            u"se mantienen como están en cada plantilla. Se muestra el estado "
            u"de \"%s\"; los campos sin tocar no se escriben.") % erste.Name))

    felder = []
    for ebene, beschriftung in EBENEN:
        gitter = _gitter("1", "auto")
        gitter.Children.Add(_text(beschriftung, 0))
        feld = TextBox()
        feld.Width = 150.0
        feld.Margin = Thickness(4.0, 2.0, 4.0, 2.0)
        try:
            feld.Text = u"%.1f" % (bereich.GetOffset(ebene) * MM_JE_FUSS)
        except Exception:
            feld.Text = u""
        Grid.SetColumn(feld, 1)
        gitter.Children.Add(feld)
        fenster.zeile(gitter, beschriftung, ebene)
        felder.append((ebene, feld, feld.Text))

    if not fenster.zeige():
        return None
    neue = []
    for ebene, feld, urtext in felder:
        text = feld.Text.strip().replace(u",", u".")
        if not text or text == urtext:
            continue
        try:
            neue.append((ebene, float(text) / MM_JE_FUSS))
        except ValueError:
            raise VorlagenFehler(t(u"\"%s\" ist keine Zahl.",
                                   u"\"%s\" is not a number.",
                                   u"\"%s\" no es un número.") % feld.Text)
    if not neue:
        return None

    def anwenden(vorlage):
        if not isinstance(vorlage, ViewPlan):
            return
        bereich = vorlage.GetViewRange()
        for ebene, versatz in neue:
            bereich.SetOffset(ebene, versatz)
        vorlage.SetViewRange(bereich)

    return Aenderung(t(u"Ansichtsbereich", u"View range", u"Rango de vista"),
                     anwenden, len(neue))


# ---------------------------------------------------------------------------
# Welcher Block hat einen Editor?
# ---------------------------------------------------------------------------

def editor_fuer(bip_name):
    """Funktion (besitzer, doc, vorlagen) -> Aenderung, oder None."""
    if bip_name in UEBERSCHRIFTEN:
        return lambda besitzer, doc, vorlagen: kategorien_bearbeiten(
            besitzer, doc, vorlagen, bip_name)
    return {
        "VIS_GRAPHICS_FILTERS": filter_bearbeiten,
        "VIS_GRAPHICS_WORKSETS": worksets_bearbeiten,
        "PLAN_VIEW_RANGE": ansichtsbereich_bearbeiten,
    }.get(bip_name)
