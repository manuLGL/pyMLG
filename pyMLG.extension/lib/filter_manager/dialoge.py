# -*- coding: utf-8 -*-
"""WPF-Grundlagen und kleine Dialoge - lauffähig unter CPython (pythonnet).

pyrevit.forms ist unter CPython nicht verfügbar (siehe
ScheduleSync.panel/README.md). Die Fenster werden deshalb als XAML-Text mit
System.Windows.Markup.XamlReader.Parse gebaut; Steuerelemente holt man per
FindName, Ereignisse werden mit += verdrahtet. Datenbindung an
Python-Objekte funktioniert unter pythonnet nicht - Listen werden im Code
befüllt.

Wichtig: Eine Ausnahme in einem Ereignishandler beendet ShowDialog() mit
einem nichtssagenden Fehler. Jeder Handler läuft deshalb über sicher().
"""

import io
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")

from System.Windows import (  # noqa: E402
    MessageBox,
    MessageBoxButton,
    MessageBoxImage,
    MessageBoxResult,
    Thickness,
    Visibility,
)
from System.Windows.Controls import (  # noqa: E402
    CheckBox,
    ComboBoxItem,
    ListBoxItem,
)
from System.Windows.Interop import WindowInteropHelper  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402

TITEL = t(u"Filter-Manager", u"Filter Manager", u"Gestor de filtros")

XMLNS = (u'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
         u'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"')


def protokollordner():
    """Fester Ordner für Protokolle.

    Nicht %TEMP%: Revit setzt TEMP pro Sitzung auf einen Unterordner mit
    zufälliger GUID - dort findet man die Dateien kaum wieder.
    """
    basis = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    ordner = os.path.join(basis, "pyMLG")
    try:
        if not os.path.isdir(ordner):
            os.makedirs(ordner)
    except Exception:
        ordner = os.environ.get("TEMP", ".")
    return ordner


FEHLERPROTOKOLL = os.path.join(protokollordner(),
                               "FilterManager_Fehler.log")


def schreibe_fehlerprotokoll(spur):
    try:
        with io.open(FEHLERPROTOKOLL, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
        return FEHLERPROTOKOLL
    except Exception:
        return None


def lade_xaml(text):
    return XamlReader.Parse(text)


def setze_besitzer(fenster, handle=None, besitzer=None):
    """Fenster vor Revit (handle) bzw. vor einem WPF-Fenster anzeigen."""
    try:
        if besitzer is not None:
            fenster.Owner = besitzer
        elif handle is not None:
            WindowInteropHelper(fenster).Owner = handle
    except Exception:
        pass


def fehlertext(fehler):
    text = getattr(fehler, "Message", None) or u"%s" % fehler
    zeilen = [z.strip() for z in text.splitlines() if z.strip()]
    return zeilen[0] if zeilen else fehler.__class__.__name__


def zeige_fehler(besitzer, fehler, fachlich=(), titel=TITEL):
    """Fachliche Fehler nur als Text, alle anderen zusätzlich protokollieren."""
    if isinstance(fehler, tuple(fachlich)):
        meldung(besitzer, u"%s" % fehler, titel=titel, warnung=True)
        return
    pfad = schreibe_fehlerprotokoll(traceback.format_exc())
    meldung(besitzer, u"%s%s" % (
        fehlertext(fehler),
        t(u"\n\nTechnische Details: %s", u"\n\nTechnical details: %s", u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
        titel=titel, warnung=True)


def sicher(besitzer_liefern, funktion, fachlich=()):
    """Ereignishandler, der Ausnahmen abfängt und verständlich anzeigt."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except Exception as fehler:
            try:
                zeige_fehler(besitzer_liefern(), fehler, fachlich)
            except Exception:
                schreibe_fehlerprotokoll(traceback.format_exc())
    return handler


def meldung(besitzer, text, titel=TITEL, warnung=False):
    symbol = MessageBoxImage.Warning if warnung else MessageBoxImage.Information
    if besitzer is not None:
        MessageBox.Show(besitzer, text, titel, MessageBoxButton.OK, symbol)
    else:
        MessageBox.Show(text, titel, MessageBoxButton.OK, symbol)


def frage(besitzer, text, titel=TITEL, warnung=False, abbrechen=False):
    """Ja/Nein(/Abbrechen). Rückgabe True, False oder None (Abbrechen)."""
    knoepfe = (MessageBoxButton.YesNoCancel if abbrechen
               else MessageBoxButton.YesNo)
    symbol = MessageBoxImage.Warning if warnung else MessageBoxImage.Question
    ergebnis = MessageBox.Show(besitzer, text, titel, knoepfe, symbol)
    if ergebnis == MessageBoxResult.Yes:
        return True
    if ergebnis == MessageBoxResult.No:
        return False
    return None


def _escape(text):
    return (text.replace(u"&", u"&amp;").replace(u"<", u"&lt;")
            .replace(u">", u"&gt;").replace(u'"', u"&quot;"))


# ---------------------------------------------------------------------------
# Texteingabe
# ---------------------------------------------------------------------------

_XAML_EINGABE_TEXTE = {
    "t0": (u"Abbrechen",
           u"Cancel",
           u"Cancelar"),
}

_XAML_EINGABE = u"""
<Window %s Title="{titel}" Width="460" SizeToContent="Height"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12">
  <StackPanel Margin="14">
    <TextBlock x:Name="hinweis" TextWrapping="Wrap" Margin="0,0,0,8"/>
    <TextBox x:Name="eingabe" Padding="3"/>
    <TextBlock x:Name="fehler" Foreground="#C0392B" TextWrapping="Wrap"
               Margin="0,6,0,0" Visibility="Collapsed"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"
                Margin="0,14,0,0">
      <Button x:Name="ok" Content="OK" Width="90" Margin="0,0,8,0"
              IsDefault="True"/>
      <Button Content="{{t0}}" Width="90" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>""" % XMLNS


def frage_text(besitzer, titel, hinweis, vorgabe=u"", pruefen=None):
    """Einzeilige Eingabe. pruefen(text) liefert den bereinigten Text oder
    wirft eine Ausnahme, deren Text im Dialog angezeigt wird.
    Rückgabe: Text oder None bei Abbruch."""
    fenster = lade_xaml(uebersetze_xaml(
        _XAML_EINGABE.replace(u"{titel}", _escape(titel)),
        _XAML_EINGABE_TEXTE))
    setze_besitzer(fenster, besitzer=besitzer)
    fenster.FindName("hinweis").Text = hinweis
    eingabe = fenster.FindName("eingabe")
    fehler_text = fenster.FindName("fehler")
    eingabe.Text = vorgabe or u""
    ergebnis = {"text": None}

    def bei_ok(sender, args):
        try:
            text = pruefen(eingabe.Text) if pruefen else eingabe.Text
        except Exception as fehler:
            fehler_text.Text = (u"%s" % fehler.args[0] if getattr(
                fehler, "args", None) else fehlertext(fehler))
            fehler_text.Visibility = Visibility.Visible
            return
        ergebnis["text"] = text
        fenster.DialogResult = True

    fenster.FindName("ok").Click += sicher(lambda: fenster, bei_ok)

    def bei_laden(sender, args):
        eingabe.Focus()
        eingabe.SelectAll()

    fenster.Loaded += bei_laden
    fenster.ShowDialog()
    return ergebnis["text"]


# ---------------------------------------------------------------------------
# Auswahlliste mit Suche
# ---------------------------------------------------------------------------

_XAML_AUSWAHL_TEXTE = {
    "t0": (u"Suchen:",
           u"Search:",
           u"Buscar:"),
    "t1": (u"Alle markieren",
           u"Check all",
           u"Marcar todo"),
    "t2": (u"Keine markieren",
           u"Uncheck all",
           u"Desmarcar todo"),
    "t3": (u"Abbrechen",
           u"Cancel",
           u"Cancelar"),
}

_XAML_AUSWAHL = u"""
<Window %s Title="{titel}" Width="560" Height="620"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="12" MinWidth="380" MinHeight="300">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <TextBlock x:Name="hinweis" TextWrapping="Wrap" Margin="0,0,0,8"/>
    <DockPanel Grid.Row="1" Margin="0,0,0,6">
      <TextBlock Text="{{t0}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="suche" Padding="3"/>
    </DockPanel>
    <ListBox x:Name="liste" Grid.Row="2"/>
    <StackPanel x:Name="optionen" Grid.Row="3" Orientation="Horizontal"
                Margin="0,8,0,0" Visibility="Collapsed">
      <TextBlock x:Name="optionen_text" VerticalAlignment="Center"
                 Margin="0,0,8,0"/>
      <ComboBox x:Name="optionen_liste" MinWidth="220"/>
    </StackPanel>
    <DockPanel Grid.Row="4" Margin="0,12,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="OK" Width="90" Margin="0,0,8,0"/>
        <Button Content="{{t3}}" Width="90" IsCancel="True"/>
      </StackPanel>
      <StackPanel x:Name="markieren" DockPanel.Dock="Left"
                  Orientation="Horizontal">
        <Button x:Name="alle" Content="{{t1}}" Padding="8,2"
                Margin="0,0,6,0"/>
        <Button x:Name="keine" Content="{{t2}}" Padding="8,2"/>
      </StackPanel>
      <TextBlock x:Name="zaehler" VerticalAlignment="Center" Margin="10,0"
                 Foreground="#666" TextTrimming="CharacterEllipsis"/>
    </DockPanel>
  </Grid>
</Window>""" % XMLNS


def waehle(besitzer, titel, hinweis, eintraege, mehrfach=False,
           vorauswahl=(), optionen=None):
    """Auswahl aus [(Text, Wert)] mit Live-Suche.

    mehrfach   Kontrollkästchen; die Markierung bleibt beim Suchen erhalten
    optionen   (Beschriftung, [Texte], Startindex) - zusätzliche Auswahlbox
    Rückgabe:  Wert bzw. Liste der Werte (mehrfach) - mit optionen als
               (Auswahl, Optionsindex); None bei Abbruch.
    """
    fenster = lade_xaml(uebersetze_xaml(
        _XAML_AUSWAHL.replace(u"{titel}", _escape(titel)),
        _XAML_AUSWAHL_TEXTE))
    setze_besitzer(fenster, besitzer=besitzer)
    fenster.FindName("hinweis").Text = hinweis
    suche = fenster.FindName("suche")
    liste = fenster.FindName("liste")
    zaehler = fenster.FindName("zaehler")
    vorauswahl = set(vorauswahl)
    elemente = []          # (Suchtext, Steuerelement, Wert)

    for text, wert in eintraege:
        if mehrfach:
            box = CheckBox()
            box.Content = text
            box.IsChecked = wert in vorauswahl
            box.Margin = Thickness(2.0, 2.0, 2.0, 2.0)
            liste.Items.Add(box)
            elemente.append((text.lower(), box, wert))
        else:
            item = ListBoxItem()
            item.Content = text
            liste.Items.Add(item)
            elemente.append((text.lower(), item, wert))

    if not mehrfach:
        fenster.FindName("markieren").Visibility = Visibility.Collapsed

    def aktualisiere_zaehler():
        sichtbar = sum(1 for _t, e, _w in elemente
                       if e.Visibility == Visibility.Visible)
        if mehrfach:
            markiert = sum(1 for _t, e, _w in elemente if e.IsChecked)
            zaehler.Text = t(u"%d markiert, %d von %d angezeigt", u"%d checked, %d of %d shown", u"%d marcados, %d de %d mostrados") % (
                markiert, sichtbar, len(elemente))
        else:
            zaehler.Text = t(u"%d von %d", u"%d of %d", u"%d de %d") % (sichtbar, len(elemente))

    def bei_suche(sender, args):
        woerter = suche.Text.lower().split()
        for text, element, _wert in elemente:
            passt = all(w in text for w in woerter)
            element.Visibility = (Visibility.Visible if passt
                                  else Visibility.Collapsed)
        aktualisiere_zaehler()

    def markiere(zustand):
        for _text, element, _wert in elemente:
            if element.Visibility == Visibility.Visible:
                element.IsChecked = zustand
        aktualisiere_zaehler()

    ergebnis = {"wert": None, "ok": False}
    options_liste = fenster.FindName("optionen_liste")
    if optionen:
        beschriftung, texte, start = optionen
        fenster.FindName("optionen").Visibility = Visibility.Visible
        fenster.FindName("optionen_text").Text = beschriftung
        for text in texte:
            item = ComboBoxItem()
            item.Content = text
            options_liste.Items.Add(item)
        options_liste.SelectedIndex = start

    def bei_ok(sender, args):
        if mehrfach:
            wert = [w for _t, e, w in elemente if e.IsChecked]
            if not wert:
                meldung(fenster, t(u"Bitte mindestens einen Eintrag markieren.", u"Please check at least one entry.", u"Marque al menos una entrada."),
                        titel=titel)
                return
        else:
            item = liste.SelectedItem
            treffer = [w for _t, e, w in elemente if e is item or e == item]
            if not treffer:
                meldung(fenster, t(u"Bitte einen Eintrag auswählen.", u"Please select an entry.", u"Seleccione una entrada."),
                        titel=titel)
                return
            wert = treffer[0]
        ergebnis["wert"] = wert
        ergebnis["ok"] = True
        fenster.DialogResult = True

    def bei_doppelklick(sender, args):
        if not mehrfach and liste.SelectedItem is not None:
            bei_ok(sender, args)

    besitzer_liefern = lambda: fenster  # noqa: E731
    suche.TextChanged += sicher(besitzer_liefern, bei_suche)
    fenster.FindName("ok").Click += sicher(besitzer_liefern, bei_ok)
    fenster.FindName("alle").Click += sicher(
        besitzer_liefern, lambda s, a: markiere(True))
    fenster.FindName("keine").Click += sicher(
        besitzer_liefern, lambda s, a: markiere(False))
    liste.MouseDoubleClick += sicher(besitzer_liefern, bei_doppelklick)
    for _text, element, _wert in elemente:
        if mehrfach:
            element.Click += sicher(besitzer_liefern,
                                    lambda s, a: aktualisiere_zaehler())

    fenster.Loaded += lambda s, a: suche.Focus()
    aktualisiere_zaehler()
    fenster.ShowDialog()
    if not ergebnis["ok"]:
        return None
    if optionen:
        return ergebnis["wert"], options_liste.SelectedIndex
    return ergebnis["wert"]
