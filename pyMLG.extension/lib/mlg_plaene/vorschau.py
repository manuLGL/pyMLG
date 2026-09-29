# -*- coding: utf-8 -*-
"""Vorschaufenster der Plan-Werkzeuge: der Plan verkleinert, die Ansichten
als Rechtecke mit Namen, per Maus verschiebbar.

    ergebnis = vorschau.zeigen(titel, kopf, flaeche, eintraege, abstand, linien)

eintraege: Liste von Eintrag (feste werden grau gezeigt). Liefert
{kennung: neues Rechteck} der beweglichen Einträge oder None (Abbrechen).
Rechtecke in Plankoordinaten (Fuss), y nach oben - siehe logik. linien:
die Linien des Plankopfs ([[(x, y), ...], ...]), damit er erkennbar ist.

Bedienung: Klicken wählt, Strg+Klick wählt mehrere, Ziehen verschiebt die
Auswahl (rastet an Kanten und Mitten ein), Pfeiltasten schieben um 1 mm,
mit Umschalt um 1 cm. Läuft unter IronPython (pyrevit.forms).
"""

from System.Windows import Thickness, TextWrapping
from System.Windows.Controls import Border, Canvas, TextBlock
from System.Windows.Input import Key, Keyboard, ModifierKeys
from System.Windows import Point
from System.Windows.Media import (Color, DoubleCollection, PointCollection,
                                  SolidColorBrush)
from System.Windows.Shapes import Polyline, Rectangle
from pyrevit import forms

from mlg_plaene import logik as lg
from mlg_sprache import t, uebersetze_xaml

MAX_BREITE, MAX_HOEHE = 920.0, 560.0
EINRASTEN_PX = 6.0
MM = 0.1 / 30.48


def _pinsel(a, r, g, b):
    pinsel = SolidColorBrush(Color.FromArgb(a, r, g, b))
    pinsel.Freeze()
    return pinsel


DUNKEL = _pinsel(255, 68, 84, 106)
WEISS = _pinsel(255, 255, 255, 255)
GRAU_RAND = _pinsel(255, 150, 158, 168)
GRAU_FLAECHE = _pinsel(110, 176, 186, 198)
BLAU_RAND = _pinsel(255, 6, 150, 215)
BLAU_FLAECHE = _pinsel(70, 6, 150, 215)
BLAU_GEWAEHLT = _pinsel(120, 6, 150, 215)
ROT = _pinsel(255, 192, 58, 58)
ORANGE = _pinsel(255, 214, 124, 38)
FLAECHE_RAND = _pinsel(255, 176, 186, 198)
PLANKOPF_LINIE = _pinsel(255, 120, 130, 142)


class Eintrag(object):
    """Ein Rechteck in der Vorschau."""

    def __init__(self, kennung, name, rechteck, beweglich=True):
        self.kennung = kennung
        self.name = name
        self.rechteck = rechteck
        self.beweglich = beweglich


XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{{titel}}" SizeToContent="WidthAndHeight" ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen" FontSize="12">
    <DockPanel Margin="10">
        <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <Button x:Name="btn_auto" Content="{{auto}}" Padding="8,2" Margin="0,0,10,0"/>
            <Button x:Name="btn_alle" Content="{{alle}}" Padding="8,2" Margin="0,0,10,0"/>
            <TextBlock Text="{{buendig}}" VerticalAlignment="Center" Margin="0,0,4,0"/>
            <Button x:Name="btn_links" Content="{{links}}" Padding="6,2" Margin="0,0,2,0"/>
            <Button x:Name="btn_mitte_h" Content="{{mitte_h}}" Padding="6,2" Margin="0,0,2,0"/>
            <Button x:Name="btn_rechts" Content="{{rechts}}" Padding="6,2" Margin="0,0,8,0"/>
            <Button x:Name="btn_oben" Content="{{oben}}" Padding="6,2" Margin="0,0,2,0"/>
            <Button x:Name="btn_mitte_v" Content="{{mitte_v}}" Padding="6,2" Margin="0,0,2,0"/>
            <Button x:Name="btn_unten" Content="{{unten}}" Padding="6,2" Margin="0,0,10,0"/>
            <TextBlock Text="{{verteilen}}" VerticalAlignment="Center" Margin="0,0,4,0"/>
            <Button x:Name="btn_vert_h" Content="&#x2194;" Padding="8,2" Margin="0,0,2,0"/>
            <Button x:Name="btn_vert_v" Content="&#x2195;" Padding="8,2" Margin="0,0,10,0"/>
            <CheckBox x:Name="cb_einrasten" Content="{{einrasten}}" IsChecked="True"
                      VerticalAlignment="Center"/>
        </WrapPanel>
        <TextBlock DockPanel.Dock="Top" Text="{{hilfe}}" Foreground="Gray"
                   Margin="0,0,0,6" TextWrapping="Wrap" MaxWidth="920"
                   HorizontalAlignment="Left"/>
        <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0" LastChildFill="False">
            <TextBlock x:Name="lbl_status" DockPanel.Dock="Left" VerticalAlignment="Center"/>
            <Button x:Name="btn_cancel" DockPanel.Dock="Right" Content="{{abbrechen}}"
                    Width="90" Height="26" IsCancel="True"/>
            <Button x:Name="btn_ok" DockPanel.Dock="Right" Content="{{ok}}"
                    Width="120" Height="26" Margin="0,0,8,0" IsDefault="True"/>
        </DockPanel>
        <Border BorderBrush="#FF9AA3AD" BorderThickness="1">
            <Canvas x:Name="leinwand" Background="#FFEEF0F3" ClipToBounds="True"
                    Focusable="True"/>
        </Border>
    </DockPanel>
</Window>
"""

TEXTE = {
    "titel": (u"Vorschau", u"Preview", u"Vista previa"),
    "auto": (u"Automatisch anordnen", u"Arrange automatically",
             u"Organizar automáticamente"),
    "alle": (u"Alle wählen", u"Select all", u"Seleccionar todo"),
    "buendig": (u"Bündig:", u"Align:", u"Alinear:"),
    "links": (u"links", u"left", u"izq."),
    "mitte_h": (u"Mitte", u"centre", u"centro"),
    "rechts": (u"rechts", u"right", u"dcha."),
    "oben": (u"oben", u"top", u"arriba"),
    "mitte_v": (u"Mitte", u"middle", u"medio"),
    "unten": (u"unten", u"bottom", u"abajo"),
    "verteilen": (u"Verteilen:", u"Distribute:", u"Distribuir:"),
    "einrasten": (u"Einrasten", u"Snap", u"Ajustar"),
    "hilfe": (u"Klicken wählt, Strg+Klick wählt mehrere, Ziehen verschiebt. "
              u"Pfeiltasten: 1 mm, mit Umschalt 1 cm. Bündig bei einer Ansicht: "
              u"an der Zeichenfläche (gestrichelt). Grau = schon auf dem Plan, "
              u"rot = überschneidet sich.",
              u"Click selects, Ctrl+click selects several, drag moves. Arrow "
              u"keys: 1 mm, with Shift 1 cm. Align with one view: to the drawing "
              u"area (dashed). Grey = already on the sheet, red = overlaps.",
              u"Clic selecciona, Ctrl+clic selecciona varias, arrastrar mueve. "
              u"Flechas: 1 mm, con Mayús 1 cm. Alinear con una vista: al área de "
              u"dibujo (discontinua). Gris = ya en el plano, rojo = se solapa."),
    "ok": (u"Übernehmen", u"Apply", u"Aplicar"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}


class VorschauFenster(forms.WPFWindow):
    def __init__(self, titel, kopf, flaeche, eintraege, abstand=0.0, linien=None):
        forms.WPFWindow.__init__(self, uebersetze_xaml(XAML, TEXTE),
                                 literal_string=True)
        self.Title = titel
        self.kopf = kopf
        self.flaeche = flaeche
        self.eintraege = eintraege
        self.abstand = abstand
        self.rechtecke = [e.rechteck for e in eintraege]
        self.beweglich = [i for i, e in enumerate(eintraege) if e.beweglich]
        self.auswahl = set()
        self.ziehen = None
        self.ergebnis = None

        # Ausschnitt: Plankopf und alle Ansichten, mit Luft zum Verschieben
        rahmen = lg.huelle([kopf] + self.rechtecke)
        luft = 0.08 * max(lg.breite(rahmen), lg.hoehe(rahmen))
        self.ausschnitt = (rahmen[0] - luft, rahmen[1] - luft,
                           rahmen[2] + luft, rahmen[3] + luft)
        self.massstab = min(MAX_BREITE / lg.breite(self.ausschnitt),
                            MAX_HOEHE / lg.hoehe(self.ausschnitt))
        self.leinwand.Width = lg.breite(self.ausschnitt) * self.massstab
        self.leinwand.Height = lg.hoehe(self.ausschnitt) * self.massstab

        self._hintergrund(linien or [])
        self.kaesten = [self._kasten(i, e) for i, e in enumerate(eintraege)]

        self.leinwand.MouseLeftButtonDown += self._leer_geklickt
        self.leinwand.MouseMove += self._bewegen
        self.leinwand.MouseLeftButtonUp += self._loslassen
        self.PreviewKeyDown += self._taste
        self.btn_auto.Click += self._auto
        self.btn_alle.Click += self._alle
        for knopf, art in ((self.btn_links, lg.LINKS), (self.btn_mitte_h, lg.MITTE_H),
                           (self.btn_rechts, lg.RECHTS), (self.btn_oben, lg.OBEN),
                           (self.btn_mitte_v, lg.MITTE_V), (self.btn_unten, lg.UNTEN)):
            knopf.Click += (lambda s, e, art=art: self._ausrichten(art))
        self.btn_vert_h.Click += (lambda s, e: self._verteilen(True))
        self.btn_vert_v.Click += (lambda s, e: self._verteilen(False))
        self.btn_ok.Click += self._ok
        self.btn_cancel.Click += (lambda s, e: self.Close())
        self._zeichnen()

    # ---------------------------------------------------------- Zeichnen

    def _setze(self, element, r):
        x0, y0, x1, y1 = self.ausschnitt
        Canvas.SetLeft(element, (r[0] - x0) * self.massstab)
        Canvas.SetTop(element, (y1 - r[3]) * self.massstab)
        element.Width = max(2.0, lg.breite(r) * self.massstab)
        element.Height = max(2.0, lg.hoehe(r) * self.massstab)

    def _punkt(self, x, y):
        return Point((x - self.ausschnitt[0]) * self.massstab,
                     (self.ausschnitt[3] - y) * self.massstab)

    def _hintergrund(self, linien):
        blatt = Rectangle()
        blatt.Fill = WEISS
        blatt.Stroke = None if linien else DUNKEL
        blatt.StrokeThickness = 1.5
        blatt.IsHitTestVisible = False
        self._setze(blatt, self.kopf)
        self.leinwand.Children.Add(blatt)

        # Plankopf mit Rahmen und Schriftfeld, wie er auf dem Plan aussieht
        for linie in linien:
            zug = Polyline()
            punkte = PointCollection()
            for x, y in linie:
                punkte.Add(self._punkt(x, y))
            zug.Points = punkte
            zug.Stroke = PLANKOPF_LINIE
            zug.StrokeThickness = 0.8
            zug.IsHitTestVisible = False
            self.leinwand.Children.Add(zug)

        flaeche = Rectangle()
        flaeche.Stroke = FLAECHE_RAND
        flaeche.StrokeThickness = 1
        striche = DoubleCollection()
        striche.Add(4)
        striche.Add(3)
        flaeche.StrokeDashArray = striche
        flaeche.IsHitTestVisible = False
        self._setze(flaeche, self.flaeche)
        self.leinwand.Children.Add(flaeche)

    def _kasten(self, index, eintrag):
        kasten = Border()
        kasten.BorderThickness = Thickness(1.5)
        text = TextBlock()
        text.Text = eintrag.name
        text.FontSize = 11
        text.Margin = Thickness(3)
        text.TextWrapping = TextWrapping.Wrap
        text.Foreground = DUNKEL
        kasten.Child = text
        kasten.ToolTip = u"{}\n{:.0f} x {:.0f} cm".format(
            eintrag.name, lg.breite(eintrag.rechteck) * 30.48,
            lg.hoehe(eintrag.rechteck) * 30.48)
        if eintrag.beweglich:
            kasten.MouseLeftButtonDown += (
                lambda s, e, i=index: self._gedrueckt(i, e))
        else:
            kasten.IsHitTestVisible = False
        self.leinwand.Children.Add(kasten)
        return kasten

    def _zeichnen(self):
        ueber = lg.ueberschneidungen(self.rechtecke)
        anzahl_ueber = anzahl_aussen = 0
        for i, kasten in enumerate(self.kaesten):
            r = self.rechtecke[i]
            self._setze(kasten, r)
            if not self.eintraege[i].beweglich:
                kasten.Background = GRAU_FLAECHE
                kasten.BorderBrush = ROT if i in ueber else GRAU_RAND
                continue
            aussen = not lg.liegt_in(r, self.flaeche)
            anzahl_ueber += i in ueber
            anzahl_aussen += aussen
            kasten.Background = BLAU_GEWAEHLT if i in self.auswahl else BLAU_FLAECHE
            kasten.BorderThickness = Thickness(3 if i in self.auswahl else 1.5)
            kasten.BorderBrush = ROT if i in ueber else (ORANGE if aussen else BLAU_RAND)
        teile = []
        if anzahl_ueber:
            teile.append(t(u"{} überschneiden sich", u"{} overlap",
                           u"{} se solapan").format(anzahl_ueber))
        if anzahl_aussen:
            teile.append(t(u"{} außerhalb der Zeichenfläche",
                           u"{} outside the drawing area",
                           u"{} fuera del área de dibujo").format(anzahl_aussen))
        self.lbl_status.Text = u" · ".join(teile) or t(
            u"Keine Überschneidungen", u"No overlaps", u"Sin solapes")
        self.lbl_status.Foreground = ROT if anzahl_ueber else DUNKEL

    # ---------------------------------------------------------- Maus / Tasten

    def _gedrueckt(self, index, args):
        strg = (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control
        if strg:
            if index in self.auswahl:
                self.auswahl.discard(index)
            else:
                self.auswahl.add(index)
        elif index not in self.auswahl:
            self.auswahl = set([index])
        if index in self.auswahl:
            self.ziehen = {"start": args.GetPosition(self.leinwand), "haupt": index,
                           "rechtecke": dict((i, self.rechtecke[i]) for i in self.auswahl)}
            self.leinwand.CaptureMouse()
        self.leinwand.Focus()
        args.Handled = True
        self._zeichnen()

    def _leer_geklickt(self, sender, args):
        self.auswahl = set()
        self.leinwand.Focus()
        self._zeichnen()

    def _bewegen(self, sender, args):
        if not self.ziehen:
            return
        punkt = args.GetPosition(self.leinwand)
        start = self.ziehen["start"]
        dx = (punkt.X - start.X) / self.massstab
        dy = -(punkt.Y - start.Y) / self.massstab
        if self.cb_einrasten.IsChecked:
            haupt = lg.verschoben(self.ziehen["rechtecke"][self.ziehen["haupt"]], dx, dy)
            ziele = [self.rechtecke[i] for i in range(len(self.rechtecke))
                     if i not in self.auswahl] + [self.flaeche, self.kopf]
            rx, ry = lg.einrasten(haupt, ziele, EINRASTEN_PX / self.massstab)
            dx, dy = dx + rx, dy + ry
        for i, r in self.ziehen["rechtecke"].items():
            self.rechtecke[i] = lg.verschoben(r, dx, dy)
        self._zeichnen()

    def _loslassen(self, sender, args):
        if self.ziehen:
            self.ziehen = None
            self.leinwand.ReleaseMouseCapture()

    def _taste(self, sender, args):
        schritte = {Key.Left: (-1, 0), Key.Right: (1, 0), Key.Up: (0, 1), Key.Down: (0, -1)}
        if args.Key not in schritte or not self.auswahl:
            return
        weite = 10 * MM if (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift else MM
        sx, sy = schritte[args.Key]
        for i in self.auswahl:
            self.rechtecke[i] = lg.verschoben(self.rechtecke[i], sx * weite, sy * weite)
        args.Handled = True
        self._zeichnen()

    # ---------------------------------------------------------- Knöpfe

    def _gewaehlte(self):
        return sorted(self.auswahl)

    def _ausrichten(self, art):
        gewaehlt = self._gewaehlte()
        if not gewaehlt:
            return
        referenz = self.flaeche if len(gewaehlt) == 1 else None
        neu = lg.ausrichten([self.rechtecke[i] for i in gewaehlt], art, referenz)
        for i, r in zip(gewaehlt, neu):
            self.rechtecke[i] = r
        self._zeichnen()

    def _verteilen(self, waagrecht):
        gewaehlt = self._gewaehlte()
        neu = lg.verteilen([self.rechtecke[i] for i in gewaehlt], waagrecht)
        for i, r in zip(gewaehlt, neu):
            self.rechtecke[i] = r
        self._zeichnen()

    def _auto(self, sender, args):
        fest = [self.rechtecke[i] for i in range(len(self.rechtecke))
                if i not in self.beweglich]
        neu = lg.anordnen(self.flaeche, fest,
                          [self.rechtecke[i] for i in self.beweglich], self.abstand)
        for i, r in zip(self.beweglich, neu):
            self.rechtecke[i] = r
        self._zeichnen()

    def _alle(self, sender, args):
        self.auswahl = set(self.beweglich)
        self._zeichnen()

    def _ok(self, sender, args):
        self.ergebnis = dict((self.eintraege[i].kennung, self.rechtecke[i])
                             for i in self.beweglich)
        self.Close()


def zeigen(titel, kopf, flaeche, eintraege, abstand=0.0, linien=None):
    """Zeigt die Vorschau modal. {kennung: Rechteck} oder None."""
    fenster = VorschauFenster(titel, kopf, flaeche, eintraege, abstand, linien)
    fenster.ShowDialog()
    return fenster.ergebnis
