# -*- coding: utf-8 -*-
"""Das andockbare Panel des Clash Navigators (IronPython, WPF).

    ┌ Bericht ─────────────────────┬ Kollision ──────────────┐
    │ ☐ Tabelle  [Suche]  ◀ ▶ + −  │ Name, Test, Elemente... │
    │ ▼ Neu               202      ├ Bewerten ───────────────┤
    │    Kollision1 · Rohr ↔ Wand  │ Klasse [CLR ▾]          │
    │ ▶ Geprüft           133      │ Kommentar [...]         │
    │                              │ ☑ Status [Geprüft ▾]    │
    │                              │ [Anwenden] [Wie vorher] │
    │                              ├ Rückgängig / Daten /    │
    │                              │ Optionen / Baum / Log   │
    └──────────────────────────────┴─────────────────────────┘

Registriert wird das Panel einmal beim Start von Revit (startup.py der
Extension). Das Panel läuft ausserhalb des API-Kontexts: alles, was Revit
anfasst, geht als Auftrag an ein ExternalEvent. Reine Anzeige (Baum,
Speichern der Bewertungen) läuft direkt.

Ein Klick auf eine Kollision legt die Schnittbox. Mehrere schnelle Klicks
(z.B. Pfeiltasten im Baum) werden zusammengefasst - Revit springt nur zur
zuletzt gewählten.
"""

import io
import os
import time
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Xaml")
clr.AddReference("System.Data")

from System import Double, String  # noqa: E402
from System.Data import DataTable  # noqa: E402
from System.Windows import (CornerRadius, FontWeights,  # noqa: E402
                            MessageBox, Point, TextWrapping,
                            MessageBoxButton, MessageBoxImage, MessageBoxResult,
                            RoutedEventHandler, Thickness, Visibility)
from System.Windows.Controls import (Border, Canvas, ComboBoxItem,  # noqa: E402,E501
                                     GridViewColumnHeader, ListBoxItem,
                                     Orientation, StackPanel, TextBlock,
                                     TreeViewItem)
from System.Windows.Documents import Run  # noqa: E402
from System.Windows.Input import Cursors  # noqa: E402
from System.Windows.Markup import XamlReader  # noqa: E402
from System.Windows.Media import (Color, DoubleCollection,  # noqa: E402
                                  PointCollection, SolidColorBrush)
from System.Windows.Shapes import Ellipse, Polyline, Rectangle  # noqa: E402
from Microsoft.Win32 import OpenFileDialog  # noqa: E402

from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler  # noqa: E402,E501
from pyrevit import forms  # noqa: E402

from mlg_sprache import t, tt, uebersetze_xaml  # noqa: E402

from clash_navigator import PANEL_ID  # noqa: E402
from clash_navigator import bericht as br  # noqa: E402
from clash_navigator import logik as lg  # noqa: E402
from clash_navigator import speicher as sp  # noqa: E402
from clash_navigator import vorrang as vr  # noqa: E402

TITEL = u"Clash Navigator"


def _pinsel(r, g, b):
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
WEISS = _pinsel(255, 255, 255)
RAHMEN = _pinsel(221, 221, 221)
AKZENT = _pinsel(6, 150, 215)
HINDERNIS = _pinsel(190, 190, 190)
DUNKEL = _pinsel(68, 84, 106)
STATUSFARBEN = {
    br.NEU: _pinsel(214, 69, 65),
    br.AKTIV: _pinsel(230, 126, 34),
    br.GEPRUEFT: _pinsel(6, 150, 215),
    br.GENEHMIGT: _pinsel(39, 174, 96),
    br.BEHOBEN: _pinsel(149, 165, 166),
}

KOORDINATEN = (
    (u"auto", (u"Automatisch", u"Automatic", u"Automático")),
    (u"gemeinsam", (u"Gemeinsame Koordinaten", u"Shared coordinates",
                    u"Coordenadas compartidas")),
    (u"intern", (u"Interne Koordinaten", u"Internal coordinates",
                 u"Coordenadas internas")),
)

XMLNS = (u'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" '
         u'xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"')

XAML_TEXTE = {
    "kein_bericht": (u"Kein Bericht geladen", u"No report loaded",
                     u"Ningún informe cargado"),
    "tabelle": (u"Tabelle", u"Table", u"Tabla"),
    "suche_tip": (u"Suche in Name, Test, Ebene, Elementen, IDs und "
                  u"Kommentaren. Mehrere Wörter müssen alle vorkommen.",
                  u"Search name, test, level, elements, ids and comments. "
                  u"All words must match.",
                  u"Busca en nombre, prueba, nivel, elementos, ID y "
                  u"comentarios. Deben aparecer todas las palabras."),
    "zurueck_tip": (u"Vorherige Kollision", u"Previous clash",
                    u"Conflicto anterior"),
    "vor_tip": (u"Nächste Kollision", u"Next clash", u"Conflicto siguiente"),
    "auf": (u"Alle auf", u"Expand", u"Expandir"),
    "zu": (u"Alle zu", u"Collapse", u"Contraer"),
    "hinweis": (u"Klick auf eine Kollision \u2192 Schnittbox (auch in "
                u"Verknüpfungen).",
                u"Click a clash \u2192 section box (links included).",
                u"Clic en un conflicto \u2192 caja de sección (también en "
                u"vínculos)."),
    "sp_name": (u"Name", u"Name", u"Nombre"),
    "sp_status": (u"Status", u"Status", u"Estado"),
    "sp_klasse": (u"Klasse", u"Class", u"Clase"),
    "sp_abstand": (u"Abstand", u"Distance", u"Distancia"),
    "sp_test": (u"Test", u"Test", u"Prueba"),
    "sp_ebene": (u"Ebene", u"Level", u"Nivel"),
    "sp_elemente": (u"Elemente", u"Elements", u"Elementos"),
    "kollision": (u"Kollision", u"Clash", u"Conflicto"),
    "bewerten": (u"Bewerten", u"Review", u"Revisar"),
    "klasse": (u"Klassifizierung", u"Classification", u"Clasificación"),
    "kommentar": (u"Kommentar", u"Comment", u"Comentario"),
    "status_aendern": (u"Status ändern", u"Change status",
                       u"Cambiar estado"),
    "anwenden": (u"Anwenden", u"Apply", u"Aplicar"),
    "wie_vorher": (u"Wie vorher", u"Apply previous", u"Como antes"),
    "wie_vorher_tip": (u"Die zuletzt angewendete Bewertung auf die gewählte "
                       u"Kollision übertragen.",
                       u"Apply the last used review to the selected clash.",
                       u"Aplicar la última revisión al conflicto "
                       u"seleccionado."),
    "loesen": (u"Lösen (Umgehung)", u"Resolve (bypass)",
               u"Resolver (desvío)"),
    "abstand": (u"Sicherheitsabstand:", u"Clearance:",
                u"Distancia de seguridad:"),
    "berechnen": (u"Varianten berechnen", u"Calculate options",
                  u"Calcular variantes"),
    "berechnen_tip": (u"Umgehungen oben, unten und seitlich berechnen, gegen "
                      u"die Umgebung (auch Verknüpfungen) prüfen und die "
                      u"besten 3 zeigen.",
                      u"Calculate bypasses above, below and to the sides, "
                      u"check them against the surroundings (links "
                      u"included) and show the best 3.",
                      u"Calcula desvíos por arriba, abajo y a los lados, los "
                      u"comprueba contra el entorno (también vínculos) y "
                      u"muestra los 3 mejores."),
    "uebernehmen": (u"Übernehmen", u"Apply", u"Aplicar"),
    "uebernehmen_tip": (u"Die gewählte Variante modellieren - Bögen setzt "
                        u"Revit nach den Routing-Einstellungen. Strg+Z in "
                        u"Revit nimmt es zurück.",
                        u"Model the chosen option - Revit places the elbows "
                        u"from the routing preferences. Ctrl+Z in Revit "
                        u"undoes it.",
                        u"Modela la variante elegida; Revit coloca los codos "
                        u"según las preferencias de trazado. Ctrl+Z en "
                        u"Revit lo deshace."),
    "vorschau_weg": (u"Vorschau entfernen", u"Remove preview",
                     u"Quitar vista previa"),
    "vorrang": (u"Vorrang (wer weicht aus?)", u"Priority (who gives way?)",
                u"Prioridad (¿quién se desvía?)"),
    "vorrang_hinweis": (u"Oben bleibt liegen, unten weicht aus. Erkannt wird "
                        u"das Gewerk am Namen des Rohrsystems, sonst an der "
                        u"Systemklassifizierung.",
                        u"Top stays, bottom gives way. The trade is taken "
                        u"from the pipe system name, otherwise from the "
                        u"system classification.",
                        u"Arriba se queda, abajo se desvía. El oficio se "
                        u"reconoce por el nombre del sistema de tuberías o, "
                        u"si no, por su clasificación."),
    "standard": (u"Standard", u"Default", u"Por defecto"),
    "rueckgaengig": (u"Rückgängig", u"Undo", u"Deshacer"),
    "daten": (u"Daten", u"Data", u"Datos"),
    "laden": (u"Bericht laden...", u"Load report...", u"Cargar informe..."),
    "laden_tip": (u"Kollisionsbericht aus Navisworks: XML (empfohlen) oder "
                  u"CSV.",
                  u"Clash report from Navisworks: XML (recommended) or CSV.",
                  u"Informe de conflictos de Navisworks: XML (recomendado) "
                  u"o CSV."),
    "neu_laden": (u"Neu laden", u"Reload", u"Recargar"),
    "neu_laden_tip": (u"Bericht und Bewertungen neu einlesen - z.B. nach "
                      u"einem neuen Export aus Navisworks oder um den Stand "
                      u"der Kollegen zu sehen.",
                      u"Re-read report and reviews - e.g. after a new "
                      u"export from Navisworks or to see your colleagues' "
                      u"changes.",
                      u"Volver a leer informe y revisiones, p. ej. tras una "
                      u"nueva exportación o para ver los cambios del "
                      u"equipo."),
    "optionen": (u"Optionen", u"Options", u"Opciones"),
    "schnitt": (u"Schnittbox = Überschneidung", u"Section box = intersection",
                u"Caja = intersección"),
    "schnitt_tip": (u"An: nur der Bereich, in dem sich die Elemente "
                    u"überschneiden. Aus: beide Elemente ganz.",
                    u"On: only where the elements overlap. Off: both "
                    u"elements entirely.",
                    u"Activado: solo donde se solapan los elementos. "
                    u"Desactivado: ambos elementos completos."),
    "rand": (u"Rand der Schnittbox:", u"Section box offset:",
             u"Margen de la caja:"),
    "nur_offen": (u"Nur offene (Neu, Aktiv)", u"Only open (new, active)",
                  u"Solo abiertos (nuevo, activo)"),
    "nur_aktiv": (u"Nur aktives Modell", u"Active model only",
                  u"Solo el modelo activo"),
    "nur_aktiv_tip": (u"Nur Kollisionen, an denen ein Element des geöffneten "
                      u"Modells beteiligt ist (Vergleich über den "
                      u"Dateinamen aus Navisworks).",
                      u"Only clashes involving an element of the open model "
                      u"(matched by the file name from Navisworks).",
                      u"Solo conflictos con un elemento del modelo abierto "
                      u"(según el nombre de archivo de Navisworks)."),
    "auswaehlen": (u"Elemente auswählen", u"Select elements",
                   u"Seleccionar elementos"),
    "aktive_ansicht": (u"Aktive 3D-Ansicht verwenden",
                       u"Use active 3D view", u"Usar la vista 3D activa"),
    "aktive_ansicht_tip": (u"Aus: eine eigene Ansicht 'pyMLG Clash - "
                           u"Benutzer' wird angelegt und benutzt.",
                           u"Off: a dedicated view 'pyMLG Clash - user' is "
                           u"created and used.",
                           u"Desactivado: se crea y usa una vista propia "
                           u"'pyMLG Clash - usuario'."),
    "weiter": (u"Nach Anwenden zur nächsten", u"Next clash after apply",
               u"Siguiente tras aplicar"),
    "koordinaten": (u"Koordinaten aus Navisworks:",
                    u"Coordinates from Navisworks:",
                    u"Coordenadas de Navisworks:"),
    "koordinaten_tip": (u"Nur für den Kollisionspunkt. Automatisch wählt je "
                        u"Kollision die Deutung, die zu den Elementen passt.",
                        u"Only for the clash point. Automatic picks the "
                        u"interpretation that fits the elements.",
                        u"Solo para el punto de conflicto. Automático elige "
                        u"la interpretación que encaja con los elementos."),
    "baumstruktur": (u"Baumstruktur", u"Tree structure",
                     u"Estructura del árbol"),
    "ebene1": (u"1. Ebene", u"1st level", u"1er nivel"),
    "ebene2": (u"2. Ebene", u"2nd level", u"2º nivel"),
    "protokoll": (u"Protokoll", u"Log", u"Registro"),
    "leeren": (u"Leeren", u"Clear", u"Vaciar"),
}

XAML = u"""<Grid %s Background="#F3F3F3" TextElement.FontFamily="Segoe UI"
      TextElement.FontSize="12">
  <Grid.Resources>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,3,0,3"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal" Background="Transparent">
              <Border x:Name="bahn" Width="28" Height="15" CornerRadius="7.5"
                      Background="#C8C8C8" VerticalAlignment="Center">
                <Ellipse x:Name="knopf" Width="11" Height="11" Fill="White"
                         HorizontalAlignment="Left" Margin="2,0,2,0"/>
              </Border>
              <ContentPresenter Margin="7,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter TargetName="bahn" Property="Background"
                        Value="#0696D7"/>
                <Setter TargetName="knopf" Property="HorizontalAlignment"
                        Value="Right"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.5"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="9,4,9,4"/>
      <Setter Property="Margin" Value="0,0,5,5"/>
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#C8C8C8"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="rahmen" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="3"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center"
                                VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="rahmen" Property="BorderBrush"
                        Value="#0696D7"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="primaer" TargetType="Button"
           BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="#0696D7"/>
      <Setter Property="BorderBrush" Value="#0696D7"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style x:Key="karte" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#DDDDDD"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="4"/>
      <Setter Property="Padding" Value="8,4,8,6"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>
    <Style TargetType="Expander">
      <Setter Property="Foreground" Value="#0077B6"/>
    </Style>
    <Style x:Key="klein" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#6E6E6E"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="FontSize" Value="11"/>
    </Style>
    <Style x:Key="beschriftung" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#333333"/>
      <Setter Property="Margin" Value="0,4,0,2"/>
    </Style>
  </Grid.Resources>
  <Grid.ColumnDefinitions>
    <ColumnDefinition Width="*" MinWidth="180"/>
    <ColumnDefinition Width="5"/>
    <ColumnDefinition Width="250" MinWidth="190"/>
  </Grid.ColumnDefinitions>

  <!-- links: Bericht -->
  <Border Grid.Column="0" Style="{StaticResource karte}" Margin="6,6,0,6"
          Padding="6">
    <DockPanel>
      <Border DockPanel.Dock="Top" Background="#E3F2FB" CornerRadius="3"
              Padding="6,4,6,4" Margin="0,0,0,6">
        <TextBlock x:Name="kopf" Text="{{kein_bericht}}" Foreground="#0077B6"
                   FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
      </Border>
      <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,4">
        <CheckBox x:Name="tabelle_an" Content="{{tabelle}}"
                  Margin="0,0,10,5" VerticalAlignment="Center"/>
        <TextBox x:Name="suche" Width="120" Margin="0,0,5,5" Padding="3,2,3,2"
                 ToolTip="{{suche_tip}}" VerticalContentAlignment="Center"/>
        <Button x:Name="zurueck" Content="&#x25C0;" ToolTip="{{zurueck_tip}}"
                Padding="7,3,7,3"/>
        <Button x:Name="vor" Content="&#x25B6;" ToolTip="{{vor_tip}}"
                Padding="7,3,7,3"/>
        <Button x:Name="aufklappen" Content="{{auf}}" Padding="7,3,7,3"/>
        <Button x:Name="zuklappen" Content="{{zu}}" Padding="7,3,7,3"/>
      </WrapPanel>
      <TextBlock DockPanel.Dock="Bottom" Text="{{hinweis}}"
                 Style="{StaticResource klein}" Margin="0,4,0,0"/>
      <TextBlock x:Name="zaehler" DockPanel.Dock="Bottom"
                 Style="{StaticResource klein}" Margin="0,4,0,0"/>
      <Grid>
        <TreeView x:Name="baum" BorderThickness="0"/>
        <ListView x:Name="liste" Visibility="Collapsed" BorderThickness="0"
                  SelectionMode="Extended"
                  VirtualizingStackPanel.IsVirtualizing="True">
          <ListView.View>
            <GridView>
              <GridViewColumn Width="110" DisplayMemberBinding="{Binding Name}">
                <GridViewColumnHeader Content="{{sp_name}}" Tag="Name"/>
              </GridViewColumn>
              <GridViewColumn Width="75" DisplayMemberBinding="{Binding Status}">
                <GridViewColumnHeader Content="{{sp_status}}" Tag="StatusNr"/>
              </GridViewColumn>
              <GridViewColumn Width="55" DisplayMemberBinding="{Binding Klasse}">
                <GridViewColumnHeader Content="{{sp_klasse}}" Tag="Klasse"/>
              </GridViewColumn>
              <GridViewColumn Width="65"
                              DisplayMemberBinding="{Binding AbstandText}">
                <GridViewColumnHeader Content="{{sp_abstand}}" Tag="Abstand"/>
              </GridViewColumn>
              <GridViewColumn Width="100" DisplayMemberBinding="{Binding Test}">
                <GridViewColumnHeader Content="{{sp_test}}" Tag="Test"/>
              </GridViewColumn>
              <GridViewColumn Width="80" DisplayMemberBinding="{Binding Ebene}">
                <GridViewColumnHeader Content="{{sp_ebene}}" Tag="Ebene"/>
              </GridViewColumn>
              <GridViewColumn Width="220"
                              DisplayMemberBinding="{Binding Elemente}">
                <GridViewColumnHeader Content="{{sp_elemente}}"
                                      Tag="Elemente"/>
              </GridViewColumn>
            </GridView>
          </ListView.View>
        </ListView>
      </Grid>
    </DockPanel>
  </Border>

  <GridSplitter Grid.Column="1" Width="5" HorizontalAlignment="Stretch"
                Background="Transparent"/>

  <!-- rechts: Bearbeiten und Einstellungen -->
  <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto"
                Margin="0,6,6,6">
    <StackPanel>
      <Border Style="{StaticResource karte}">
        <Expander Header="{{kollision}}" IsExpanded="True">
          <TextBox x:Name="details" IsReadOnly="True" BorderThickness="0"
                   TextWrapping="Wrap" Foreground="#333333" Margin="0,4,0,0"
                   Background="Transparent" MaxHeight="220"
                   VerticalScrollBarVisibility="Auto"/>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{bewerten}}" IsExpanded="True">
          <StackPanel Margin="0,2,0,0">
            <TextBlock Text="{{klasse}}" Style="{StaticResource beschriftung}"/>
            <ComboBox x:Name="klasse" IsEditable="True"/>
            <TextBlock Text="{{kommentar}}"
                       Style="{StaticResource beschriftung}"/>
            <TextBox x:Name="kommentar" Height="46" AcceptsReturn="True"
                     TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
            <DockPanel Margin="0,6,0,6">
              <CheckBox x:Name="status_aendern" Content="{{status_aendern}}"
                        IsChecked="True" DockPanel.Dock="Left"/>
              <ComboBox x:Name="status" Margin="8,0,0,0"/>
            </DockPanel>
            <WrapPanel>
              <Button x:Name="anwenden" Content="{{anwenden}}"
                      Style="{StaticResource primaer}"/>
              <Button x:Name="wie_vorher" Content="{{wie_vorher}}"
                      ToolTip="{{wie_vorher_tip}}"/>
            </WrapPanel>
            <TextBlock x:Name="bewerten_hinweis"
                       Style="{StaticResource klein}"/>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{loesen}}" IsExpanded="True">
          <StackPanel Margin="0,4,0,0">
            <WrapPanel>
              <CheckBox x:Name="winkel45" Content="45&#x00B0;"
                        Margin="0,3,12,3"/>
              <CheckBox x:Name="winkel90" Content="90&#x00B0;"
                        Margin="0,3,12,3"/>
            </WrapPanel>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,3">
              <TextBlock Text="{{abstand}}" VerticalAlignment="Center"
                         Foreground="#333333"/>
              <TextBox x:Name="abstand" Width="40" Margin="6,0,4,0"
                       HorizontalContentAlignment="Right"/>
              <TextBlock Text="cm" VerticalAlignment="Center"
                         Foreground="#333333"/>
            </StackPanel>
            <Button x:Name="berechnen" Content="{{berechnen}}"
                    ToolTip="{{berechnen_tip}}"
                    Style="{StaticResource primaer}" Margin="0,6,0,6"
                    HorizontalAlignment="Left"/>
            <TextBlock x:Name="loesen_info" Style="{StaticResource klein}"/>
            <StackPanel x:Name="karten" Margin="0,4,0,0"/>
            <WrapPanel Margin="0,4,0,0">
              <Button x:Name="uebernehmen" Content="{{uebernehmen}}"
                      ToolTip="{{uebernehmen_tip}}"
                      Style="{StaticResource primaer}" IsEnabled="False"/>
              <Button x:Name="vorschau_weg" Content="{{vorschau_weg}}"/>
            </WrapPanel>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{vorrang}}">
          <StackPanel Margin="0,4,0,0">
            <TextBlock Text="{{vorrang_hinweis}}"
                       Style="{StaticResource klein}" Margin="0,0,0,4"/>
            <ListBox x:Name="vorrang_liste" Height="170"/>
            <WrapPanel Margin="0,5,0,0">
              <Button x:Name="vorrang_hoch" Content="&#x25B2;"
                      Padding="8,3,8,3"/>
              <Button x:Name="vorrang_runter" Content="&#x25BC;"
                      Padding="8,3,8,3"/>
              <Button x:Name="vorrang_standard" Content="{{standard}}"/>
            </WrapPanel>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{rueckgaengig}}" IsExpanded="True">
          <StackPanel Margin="0,2,0,0">
            <TextBlock x:Name="undo_titel" FontWeight="SemiBold"
                       TextWrapping="Wrap" Foreground="#333333"/>
            <TextBlock x:Name="undo_text" Style="{StaticResource klein}"
                       Margin="0,2,0,4"/>
            <DockPanel>
              <Button x:Name="rueckgaengig" DockPanel.Dock="Left"
                      Content="&#x21B6; {{rueckgaengig}}"/>
              <TextBlock x:Name="undo_rest" Style="{StaticResource klein}"
                         VerticalAlignment="Center"/>
            </DockPanel>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{daten}}" IsExpanded="True">
          <StackPanel Margin="0,4,0,0">
            <WrapPanel>
              <Button x:Name="laden" Content="{{laden}}"
                      ToolTip="{{laden_tip}}"
                      Style="{StaticResource primaer}"/>
              <Button x:Name="neu_laden" Content="{{neu_laden}}"
                      ToolTip="{{neu_laden_tip}}"/>
            </WrapPanel>
            <TextBlock x:Name="datei_info" Style="{StaticResource klein}"/>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{optionen}}" IsExpanded="True">
          <StackPanel Margin="0,4,0,0">
            <CheckBox x:Name="schnitt" Content="{{schnitt}}"
                      ToolTip="{{schnitt_tip}}"/>
            <StackPanel Orientation="Horizontal" Margin="0,3,0,3">
              <TextBlock Text="{{rand}}" VerticalAlignment="Center"
                         Foreground="#333333"/>
              <TextBox x:Name="rand" Width="40" Margin="6,0,4,0"
                       HorizontalContentAlignment="Right"/>
              <TextBlock Text="cm" VerticalAlignment="Center"
                         Foreground="#333333"/>
            </StackPanel>
            <CheckBox x:Name="nur_offen" Content="{{nur_offen}}"/>
            <CheckBox x:Name="nur_aktiv" Content="{{nur_aktiv}}"
                      ToolTip="{{nur_aktiv_tip}}"/>
            <CheckBox x:Name="auswaehlen" Content="{{auswaehlen}}"/>
            <CheckBox x:Name="aktive_ansicht" Content="{{aktive_ansicht}}"
                      ToolTip="{{aktive_ansicht_tip}}"/>
            <CheckBox x:Name="weiter" Content="{{weiter}}"/>
            <TextBlock Text="{{koordinaten}}" ToolTip="{{koordinaten_tip}}"
                       Style="{StaticResource beschriftung}"/>
            <ComboBox x:Name="koordinaten" ToolTip="{{koordinaten_tip}}"/>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{baumstruktur}}">
          <StackPanel Margin="0,2,0,0">
            <TextBlock Text="{{ebene1}}" Style="{StaticResource beschriftung}"/>
            <ComboBox x:Name="ebene1"/>
            <TextBlock Text="{{ebene2}}" Style="{StaticResource beschriftung}"/>
            <ComboBox x:Name="ebene2"/>
          </StackPanel>
        </Expander>
      </Border>

      <Border Style="{StaticResource karte}">
        <Expander Header="{{protokoll}}">
          <StackPanel Margin="0,4,0,0">
            <TextBox x:Name="protokoll" IsReadOnly="True" Height="130"
                     TextWrapping="Wrap" FontSize="11"
                     VerticalScrollBarVisibility="Auto"/>
            <Button x:Name="leeren" Content="{{leeren}}" Margin="0,5,0,0"
                    HorizontalAlignment="Left"/>
          </StackPanel>
        </Expander>
      </Border>
    </StackPanel>
  </ScrollViewer>
</Grid>""" % XMLNS

FEHLERPROTOKOLL = os.path.join(sp.eigener_ordner(),
                               u"ClashNavigator_Fehler.log")


def schreibe_fehlerprotokoll(spur):
    try:
        with io.open(FEHLERPROTOKOLL, "a", encoding="utf-8") as datei:
            datei.write(sp.jetzt() + u"\n" + spur + u"\n" + u"-" * 70
                        + u"\n")
    except Exception:
        pass


ABLAUFPROTOKOLL = os.path.join(sp.eigener_ordner(),
                               u"ClashNavigator_Ablauf.log")


def schreibe_ablauf(text):
    """Eine Zeile ins Ablaufprotokoll; über 200 kB wird neu begonnen."""
    try:
        modus = "a"
        if os.path.isfile(ABLAUFPROTOKOLL) and \
                os.path.getsize(ABLAUFPROTOKOLL) > 200000:
            modus = "w"
        with io.open(ABLAUFPROTOKOLL, modus, encoding="utf-8") as datei:
            datei.write(u"%s  %s\n" % (time.strftime("%Y-%m-%d %H:%M:%S"),
                                       text))
    except Exception:
        pass


def meldung(text, warnung=False):
    MessageBox.Show(text, TITEL, MessageBoxButton.OK,
                    MessageBoxImage.Warning if warnung
                    else MessageBoxImage.Information)


def frage(text):
    return MessageBox.Show(text, TITEL, MessageBoxButton.YesNo,
                           MessageBoxImage.Question) == MessageBoxResult.Yes


# ---------------------------------------------------------------------------
# ExternalEvent: Aufträge an Revit
# ---------------------------------------------------------------------------

class Auftraege(IExternalEventHandler):
    """Sammelt Aufträge und führt sie im API-Kontext aus.

    Aufträge tragen einen Namen; ein neuer Auftrag gleichen Namens ersetzt
    den wartenden - so springt Revit bei schnellem Weiterklicken nur zur
    zuletzt gewählten Kollision.
    """

    def __init__(self):
        self.warteschlange = []
        self.bei_fehler = None

    def gib(self, name, funktion):
        self.warteschlange = [(n, f) for n, f in self.warteschlange
                              if n != name]
        self.warteschlange.append((name, funktion))

    def Execute(self, uiapp):
        auftraege, self.warteschlange = self.warteschlange, []
        for name, funktion in auftraege:
            # Beginn und Ende ins Ablaufprotokoll - hängt Revit, zeigt die
            # letzte Zeile ohne "Ende", welcher Auftrag es war
            beginn = time.time()
            schreibe_ablauf(u"Beginn %s" % name)
            try:
                funktion(uiapp)
                schreibe_ablauf(u"Ende   %s (%.1f s)"
                                % (name, time.time() - beginn))
            except Exception as fehler:
                schreibe_ablauf(u"Fehler %s (%.1f s): %s"
                                % (name, time.time() - beginn, fehler))
                schreibe_fehlerprotokoll(traceback.format_exc())
                if self.bei_fehler is not None:
                    try:
                        self.bei_fehler(fehler)
                    except Exception:
                        pass

    def GetName(self):
        return u"pyMLG Clash Navigator"


# ---------------------------------------------------------------------------
# Das Panel
# ---------------------------------------------------------------------------

class ClashPanel(forms.WPFPanel):
    panel_id = PANEL_ID
    panel_title = TITEL
    panel_source = u"(XAML im Code)"

    def __init__(self):
        # Bewusst ohne WPFPanel.__init__: der Inhalt kommt per XamlReader aus
        # dem Text oben (übersetzt), nicht aus einer Datei.
        self.inhalt = XamlReader.Parse(uebersetze_xaml(XAML, XAML_TEXTE))
        self.Content = self.inhalt

        self.auftraege = Auftraege()
        self.auftraege.bei_fehler = self._revit_fehler
        self.ereignis = None
        self.bericht = None
        self.speicher = None
        self.modellnamen = None
        self.aktuell = None             # Schlüssel der gewählten Kollision
        self.reihenfolge = []           # Schlüssel in Anzeigereihenfolge
        self.letzte_aenderung = None    # für "Wie vorher"
        self._still = False             # Auswahl ohne Sprung setzen
        self._zuletzt_geladen = False
        self._tabelle = None            # DataTable der Tabellenansicht
        self.analyse = None             # Umgehungsvarianten (umgehung_revit)
        self.analyse_schluessel = None  # ... für diese Kollision
        self.variante = None            # gewählte Variante
        self.einstellungen = sp.lies_einstellungen()

        for name in ("kopf", "tabelle_an", "suche", "zaehler", "baum",
                     "liste", "details", "klasse", "kommentar",
                     "status_aendern", "status", "bewerten_hinweis",
                     "undo_titel", "undo_text", "undo_rest", "rueckgaengig",
                     "datei_info", "neu_laden", "schnitt", "rand",
                     "nur_offen", "nur_aktiv", "auswaehlen",
                     "aktive_ansicht", "weiter", "koordinaten", "ebene1",
                     "ebene2", "protokoll", "anwenden", "wie_vorher",
                     "winkel45", "winkel90", "abstand",
                     "loesen_info", "karten", "uebernehmen", "vorschau_weg",
                     "berechnen", "vorrang_liste"):
            setattr(self, "_" + name, self.inhalt.FindName(name))

        self._fuelle_auswahllisten()
        self._uebernimm_einstellungen()
        self._verdrahte()
        self._zeige_rueckgaengig()
        self._zeige_details(None)
        self.Loaded += self._sicher(self._beim_anzeigen)

    # --- Grundlagen -------------------------------------------------------
    def _sicher(self, funktion):
        def handler(sender, args):
            try:
                funktion(sender, args)
            except Exception as fehler:
                schreibe_fehlerprotokoll(traceback.format_exc())
                self.log(u"%s: %s" % (t(u"Fehler", u"Error", u"Error"),
                                      fehler))
                try:
                    meldung(u"%s\n\n%s" % (fehler, t(
                        u"Details: %s", u"Details: %s", u"Detalles: %s")
                        % FEHLERPROTOKOLL), warnung=True)
                except Exception:
                    pass
        return handler

    def _revit_fehler(self, fehler):
        self.log(u"%s" % fehler)
        meldung(u"%s" % fehler, warnung=True)

    def log(self, text):
        zeile = u"%s  %s" % (sp.jetzt()[11:], text)
        self._protokoll.AppendText(zeile + u"\n")
        self._protokoll.ScrollToEnd()

    def an_revit(self, name, funktion):
        if self.ereignis is None:
            meldung(t(u"Der Clash Navigator ist nicht mit Revit verbunden. "
                      u"Bitte Revit neu starten.",
                      u"The Clash Navigator is not connected to Revit. "
                      u"Please restart Revit.",
                      u"Clash Navigator no está conectado a Revit. "
                      u"Reinicie Revit."), warnung=True)
            return
        self.auftraege.gib(name, funktion)
        self.ereignis.Raise()

    # --- Aufbau -----------------------------------------------------------
    def _fuelle_auswahllisten(self):
        for code in br.STATUS:
            self._status.Items.Add(self._eintrag(lg.status_text(code), code))
        for code, texte in KOORDINATEN:
            self._koordinaten.Items.Add(self._eintrag(tt(texte), code))
        for liste in (self._ebene1, self._ebene2):
            for code, texte in lg.GRUPPIERUNGEN:
                liste.Items.Add(self._eintrag(tt(texte), code))
        self._fuelle_klassen()
        self._fuelle_vorrang()

    def _fuelle_vorrang(self, gewaehlt=None):
        liste = vr.reihenfolge(self.einstellungen.get(u"vorrang"))
        self._vorrang_liste.Items.Clear()
        for nummer, code in enumerate(liste):
            eintrag = ListBoxItem()
            eintrag.Content = u"%d.  %s" % (nummer + 1, vr.text(code))
            eintrag.Tag = code
            self._vorrang_liste.Items.Add(eintrag)
            if code == gewaehlt:
                self._vorrang_liste.SelectedItem = eintrag

    def vorrang_schieben(self, richtung):
        eintrag = self._vorrang_liste.SelectedItem
        if eintrag is None:
            return
        liste = vr.reihenfolge(self.einstellungen.get(u"vorrang"))
        alt = liste.index(eintrag.Tag)
        neu = alt + richtung
        if not 0 <= neu < len(liste):
            return
        liste[alt], liste[neu] = liste[neu], liste[alt]
        self.einstellungen[u"vorrang"] = liste
        self._speichere_einstellungen()
        self._fuelle_vorrang(gewaehlt=eintrag.Tag)

    def vorrang_standard(self):
        self.einstellungen[u"vorrang"] = []
        self._speichere_einstellungen()
        self._fuelle_vorrang()

    def _fuelle_klassen(self):
        text = self._klasse.Text
        self._klasse.Items.Clear()
        for klasse in self._klassen():
            self._klasse.Items.Add(klasse)
        self._klasse.Text = text

    def _klassen(self):
        klassen = list(lg.KLASSEN)
        for klasse in self.einstellungen.get(u"klassen") or []:
            if klasse not in klassen:
                klassen.append(klasse)
        return klassen

    @staticmethod
    def _eintrag(text, code):
        eintrag = ComboBoxItem()
        eintrag.Content = text
        eintrag.Tag = code
        return eintrag

    @staticmethod
    def _waehle_code(liste, code):
        for eintrag in liste.Items:
            if eintrag.Tag == code:
                liste.SelectedItem = eintrag
                return
        if liste.Items.Count:
            liste.SelectedIndex = 0

    @staticmethod
    def _code(liste):
        eintrag = liste.SelectedItem
        return eintrag.Tag if eintrag is not None else None

    def _uebernimm_einstellungen(self):
        e = self.einstellungen
        self._rand.Text = (u"%g" % float(e[u"rand_cm"])).replace(u".", u",")
        for name in (u"schnitt", u"nur_offen", u"nur_aktiv", u"auswaehlen",
                     u"aktive_ansicht", u"weiter", u"winkel45",
                     u"winkel90"):
            getattr(self, "_" + name).IsChecked = bool(e[name])
        self._abstand.Text = (u"%g" % float(e[u"abstand_cm"])).replace(
            u".", u",")
        self._tabelle_an.IsChecked = bool(e[u"tabelle"])
        self._waehle_code(self._koordinaten, e[u"koordinaten"])
        self._waehle_code(self._ebene1, e[u"ebene1"])
        self._waehle_code(self._ebene2, e[u"ebene2"])
        self._waehle_code(self._status, br.GEPRUEFT)
        self._zeige_ansichtsart()

    def _speichere_einstellungen(self):
        e = self.einstellungen
        try:
            e[u"rand_cm"] = max(0.0, float(self._rand.Text.replace(u",",
                                                                   u".")))
        except ValueError:
            pass
        for name in (u"schnitt", u"nur_offen", u"nur_aktiv", u"auswaehlen",
                     u"aktive_ansicht", u"weiter", u"winkel45",
                     u"winkel90"):
            e[name] = bool(getattr(self, "_" + name).IsChecked)
        try:
            e[u"abstand_cm"] = max(0.0, float(self._abstand.Text.replace(
                u",", u".")))
        except ValueError:
            pass
        e[u"tabelle"] = bool(self._tabelle_an.IsChecked)
        e[u"koordinaten"] = self._code(self._koordinaten) or u"auto"
        e[u"ebene1"] = self._code(self._ebene1) or u""
        e[u"ebene2"] = self._code(self._ebene2) or u""
        sp.schreibe_einstellungen(e)

    def _verdrahte(self):
        s = self._sicher
        f = self.inhalt.FindName
        f("laden").Click += s(lambda _s, _a: self.laden_dialog())
        self._neu_laden.Click += s(lambda _s, _a: self.neu_laden())
        self._anwenden.Click += s(lambda _s, _a: self.anwenden())
        self._wie_vorher.Click += s(lambda _s, _a: self.wie_vorher())
        self._rueckgaengig.Click += s(lambda _s, _a: self.rueckgaengig())
        f("zurueck").Click += s(lambda _s, _a: self.schritt(-1))
        f("vor").Click += s(lambda _s, _a: self.schritt(1))
        f("aufklappen").Click += s(lambda _s, _a: self.klappe(True))
        f("zuklappen").Click += s(lambda _s, _a: self.klappe(False))
        f("leeren").Click += s(lambda _s, _a: self._protokoll.Clear())

        self._baum.SelectedItemChanged += s(self._baum_auswahl)
        self._liste.SelectionChanged += s(self._liste_auswahl)
        self._liste.MouseDoubleClick += s(self._liste_doppelklick)
        self._liste.AddHandler(GridViewColumnHeader.ClickEvent,
                               RoutedEventHandler(s(self._sortiere)))
        self._suche.TextChanged += s(lambda _s, _a: self.aktualisiere())

        def option_geaendert(_s, _a):
            self._speichere_einstellungen()
            self.aktualisiere()

        def nur_speichern(_s, _a):
            self._speichere_einstellungen()

        for name in ("nur_offen", "nur_aktiv"):
            box = getattr(self, "_" + name)
            box.Checked += s(option_geaendert)
            box.Unchecked += s(option_geaendert)
        for name in ("schnitt", "auswaehlen", "aktive_ansicht", "weiter"):
            box = getattr(self, "_" + name)
            box.Checked += s(nur_speichern)
            box.Unchecked += s(nur_speichern)
        self._rand.LostFocus += s(nur_speichern)
        self._abstand.LostFocus += s(nur_speichern)
        for name in ("winkel45", "winkel90"):
            box = getattr(self, "_" + name)
            box.Checked += s(nur_speichern)
            box.Unchecked += s(nur_speichern)
        self._berechnen.Click += s(lambda _s, _a: self.berechne())
        f("vorrang_hoch").Click += s(lambda _s, _a: self.vorrang_schieben(-1))
        f("vorrang_runter").Click += s(
            lambda _s, _a: self.vorrang_schieben(1))
        f("vorrang_standard").Click += s(
            lambda _s, _a: self.vorrang_standard())
        self._uebernehmen.Click += s(lambda _s, _a: self.uebernehme())
        self._vorschau_weg.Click += s(lambda _s, _a: self.entferne_vorschau())
        self._koordinaten.SelectionChanged += s(nur_speichern)
        self._ebene1.SelectionChanged += s(option_geaendert)
        self._ebene2.SelectionChanged += s(option_geaendert)

        def ansicht_gewechselt(_s, _a):
            self._speichere_einstellungen()
            self._zeige_ansichtsart()
            self.aktualisiere()

        self._tabelle_an.Checked += s(ansicht_gewechselt)
        self._tabelle_an.Unchecked += s(ansicht_gewechselt)

    def _zeige_ansichtsart(self):
        tabelle = bool(self._tabelle_an.IsChecked)
        self._liste.Visibility = (Visibility.Visible if tabelle
                                  else Visibility.Collapsed)
        self._baum.Visibility = (Visibility.Collapsed if tabelle
                                 else Visibility.Visible)

    def _beim_anzeigen(self, _sender, _args):
        """Beim ersten Anzeigen den zuletzt benutzten Bericht laden."""
        if self._zuletzt_geladen:
            return
        self._zuletzt_geladen = True
        pfad = self.einstellungen.get(u"letzte_datei")
        if pfad and os.path.isfile(pfad) and self.bericht is None:
            self.lade(pfad, still=True)

    # --- Daten ------------------------------------------------------------
    def laden_dialog(self):
        dialog = OpenFileDialog()
        dialog.Title = t(u"Kollisionsbericht aus Navisworks wählen",
                         u"Choose a clash report from Navisworks",
                         u"Elegir un informe de conflictos de Navisworks")
        dialog.Filter = u"|".join([
            t(u"Kollisionsberichte (*.xml, *.csv)",
              u"Clash reports (*.xml, *.csv)",
              u"Informes de conflictos (*.xml, *.csv)")
            + u"|*.xml;*.csv;*.txt",
            u"Navisworks (*.nwd, *.nwf, *.nwc)|*.nwd;*.nwf;*.nwc",
            t(u"Alle Dateien", u"All files", u"Todos los archivos")
            + u"|*.*"])
        letzte = self.einstellungen.get(u"letzte_datei")
        if letzte and os.path.isdir(os.path.dirname(letzte)):
            dialog.InitialDirectory = os.path.dirname(letzte)
        if dialog.ShowDialog():
            self.lade(dialog.FileName)

    def neu_laden(self):
        if self.bericht is None:
            self.laden_dialog()
            return
        self.lade(self.bericht.pfad)

    def lade(self, pfad, still=False):
        try:
            bericht = br.lies(pfad)
        except br.Fehler as fehler:
            self.log(u"%s" % fehler)
            if not still:
                meldung(u"%s" % fehler, warnung=True)
            return
        speicher = sp.Speicher(sp.speicherpfad(pfad))
        # Rückgängig bleibt beim Neuladen desselben Berichts erhalten
        if self.speicher is not None and self.speicher.pfad == speicher.pfad:
            speicher.stapel = self.speicher.stapel
        self.bericht = bericht
        self.speicher = speicher
        self.einstellungen[u"letzte_datei"] = pfad
        self._speichere_einstellungen()

        mit_id = sum(1 for clash in bericht.clashes
                     if any(o.element_id is not None for o in clash.objekte))
        self._kopf.Text = bericht.name
        self._kopf.ToolTip = pfad
        self._datei_info.Text = t(
            u"%s-Bericht, %d Kollisionen in %d Test(s), %d mit Element-ID.\n"
            u"Bewertungen: %s",
            u"%s report, %d clashes in %d test(s), %d with element id.\n"
            u"Reviews: %s",
            u"Informe %s, %d conflictos en %d prueba(s), %d con ID de "
            u"elemento.\nRevisiones: %s") % (
                bericht.format, len(bericht.clashes), len(bericht.tests),
                mit_id, speicher.pfad)
        self.log(t(u"Geladen: %s (%d Kollisionen)",
                   u"Loaded: %s (%d clashes)",
                   u"Cargado: %s (%d conflictos)")
                 % (bericht.name, len(bericht.clashes)))
        if bericht.clashes and not mit_id:
            self.log(t(u"Hinweis: Der Bericht enthält keine Element-IDs - "
                       u"die Schnittbox kommt nur vom Kollisionspunkt.",
                       u"Note: the report contains no element ids - the "
                       u"section box uses the clash point only.",
                       u"Aviso: el informe no contiene ID de elementos; la "
                       u"caja se basa solo en el punto de conflicto."))
        self._zeige_rueckgaengig()
        self.aktualisiere()
        self.an_revit(u"modelle", self._hole_modellnamen)

    def _hole_modellnamen(self, uiapp):
        from clash_navigator import revit as rv
        uidoc = uiapp.ActiveUIDocument
        namen = rv.dokumentnamen(uidoc.Document) if uidoc else []
        if namen != self.modellnamen:
            self.modellnamen = namen
            if self._nur_aktiv.IsChecked:
                self.aktualisiere()

    def _eintraege(self):
        return self.speicher.eintraege if self.speicher else {}

    # --- Anzeige ----------------------------------------------------------
    def _sichtbare(self):
        if self.bericht is None:
            return []
        namen = self.modellnamen if self._nur_aktiv.IsChecked else None
        return lg.filtere(self.bericht.clashes, self._eintraege(),
                          nur_offen=bool(self._nur_offen.IsChecked),
                          suche=self._suche.Text, modellnamen=namen)

    def aktualisiere(self):
        """Baum bzw. Tabelle neu aufbauen - Aufklappzustand und Auswahl
        bleiben erhalten."""
        if self.bericht is None:
            return
        sichtbar = self._sichtbare()
        eintraege = self._eintraege()
        offen = sum(1 for c in sichtbar if lg.ist_offen(
            c, eintraege.get(c.schluessel)))
        self._zaehler.Text = t(
            u"%d von %d Kollisionen · %d offen",
            u"%d of %d clashes · %d open",
            u"%d de %d conflictos · %d abiertos") % (
                len(sichtbar), len(self.bericht.clashes), offen)

        self._still = True
        try:
            if self._tabelle_an.IsChecked:
                self._baue_tabelle(sichtbar)
            else:
                self._baue_baum(sichtbar)
            if self.aktuell:
                self._markiere(self.aktuell)
        finally:
            self._still = False
        self._zeige_details(self._clash(self.aktuell))

    # Baum
    def _baue_baum(self, sichtbar):
        aufgeklappt = self._aufgeklappte_pfade()
        erstes_mal = not self._baum.Items.Count
        self._baum.Items.Clear()
        ebene1 = self._code(self._ebene1) or u""
        ebene2 = self._code(self._ebene2) or u""
        struktur = lg.baum(sichtbar, self._eintraege(), ebene1, ebene2)
        self.reihenfolge = self._reihenfolge(struktur)

        if not ebene1:
            for clash in sichtbar:
                self._baum.Items.Add(self._blatt(clash))
            return
        for wert, inhalt in struktur:
            zweig = self._zweig(ebene1, wert, inhalt, ebene2, (wert,))
            self._baum.Items.Add(zweig)
            pfad = (wert,)
            if pfad in aufgeklappt or (erstes_mal and len(struktur) == 1):
                zweig.IsExpanded = True
            for kind in zweig.Items:
                if isinstance(kind, TreeViewItem) and \
                        isinstance(kind.Tag, tuple) and \
                        kind.Tag[0] in aufgeklappt:
                    kind.IsExpanded = True

    @staticmethod
    def _flach(inhalt):
        """Kollisionen eines Zweigs, auch aus seinen Unterzweigen."""
        if inhalt and isinstance(inhalt[0], tuple):
            clashes = []
            for _wert, liste in inhalt:
                clashes.extend(ClashPanel._flach(liste))
            return clashes
        return list(inhalt)

    def _reihenfolge(self, struktur):
        """Schlüssel in Anzeigereihenfolge, jede Kollision einmal (bei 'je
        Element' steht sie unter mehreren Zweigen)."""
        gesehen = set()
        reihenfolge = []
        for clash in self._flach(struktur):
            if clash.schluessel not in gesehen:
                gesehen.add(clash.schluessel)
                reihenfolge.append(clash.schluessel)
        return reihenfolge

    def _zweig(self, art, wert, inhalt, art2, pfad):
        """Ein Zweig. inhalt: Liste von Kollisionen oder von Unterzweigen."""
        unterzweige = bool(inhalt) and isinstance(inhalt[0], tuple)
        clashes = self._flach(inhalt)
        schluessel = set(clash.schluessel for clash in clashes)

        zweig = TreeViewItem()
        farbe = STATUSFARBEN.get(wert) if art == u"status" else None
        zweig.Header = self._kopfzeile(lg.zweig_text(art, wert),
                                       len(schluessel), farbe, fett=True)
        zweig.Tag = (pfad, schluessel)
        if unterzweige:
            for wert2, liste in inhalt:
                zweig.Items.Add(self._zweig(art2, wert2, liste, u"",
                                            pfad + (wert2,)))
        else:
            zweig.Items.Add(u"...")         # Platzhalter bis zum Aufklappen
            zweig.Expanded += self._sicher(
                lambda sender, args, liste=clashes: self._fuelle_zweig(
                    sender, args, liste))
        return zweig

    def _fuelle_zweig(self, sender, args, clashes):
        if args.OriginalSource is not sender:
            return
        if sender.Items.Count == 1 and not isinstance(sender.Items[0],
                                                      TreeViewItem):
            sender.Items.Clear()
            for clash in clashes:
                sender.Items.Add(self._blatt(clash))

    def _blatt(self, clash):
        eintrag = self._eintraege().get(clash.schluessel)
        blatt = TreeViewItem()
        blatt.Header = self._kopfzeile(
            lg.blatt_text(clash, eintrag), None,
            STATUSFARBEN.get(lg.status(clash, eintrag)))
        blatt.Tag = clash.schluessel
        # Doppelklick springt erneut - auch zur schon gewählten Kollision,
        # etwa nachdem man in der Ansicht weggedreht hat
        blatt.MouseDoubleClick += self._sicher(
            lambda _s, args, c=clash: self._erneut_springen(args, c))
        blatt.ToolTip = u"%s\n%s" % (clash.test, u"\n".join(
            lg.objekt_text(o) for o in clash.objekte))
        return blatt

    @staticmethod
    def _kopfzeile(text, anzahl=None, farbe=None, fett=False):
        zeile = StackPanel()
        zeile.Orientation = Orientation.Horizontal
        if farbe is not None:
            punkt = Ellipse()
            punkt.Width = punkt.Height = 8.0
            punkt.Fill = farbe
            punkt.Margin = Thickness(0.0, 0.0, 6.0, 0.0)
            zeile.Children.Add(punkt)
        beschriftung = TextBlock()
        beschriftung.Inlines.Add(Run(text))
        if fett:
            beschriftung.FontWeight = FontWeights.SemiBold
        if anzahl is not None:
            zahl = Run(u"   %d" % anzahl)
            zahl.Foreground = GRAU
            beschriftung.Inlines.Add(zahl)
        zeile.Children.Add(beschriftung)
        return zeile

    def _alle_zweige(self, items=None):
        for item in (self._baum.Items if items is None else items):
            if isinstance(item, TreeViewItem) and isinstance(item.Tag,
                                                             tuple):
                yield item
                for kind in self._alle_zweige(item.Items):
                    yield kind

    def _aufgeklappte_pfade(self):
        return set(zweig.Tag[0] for zweig in self._alle_zweige()
                   if zweig.IsExpanded)

    def klappe(self, auf):
        for zweig in list(self._alle_zweige()):
            zweig.IsExpanded = auf

    def _markiere(self, schluessel):
        """Kollision im Baum bzw. in der Tabelle auswählen und zeigen."""
        if self._tabelle_an.IsChecked:
            if self._tabelle is None:
                return False
            for zeile in self._tabelle.DefaultView:
                if zeile[u"Schluessel"] == schluessel:
                    self._liste.SelectedItem = zeile
                    self._liste.ScrollIntoView(zeile)
                    return True
            return False
        return self._markiere_im_baum(self._baum.Items, schluessel)

    def _markiere_im_baum(self, items, schluessel):
        for item in items:
            if not isinstance(item, TreeViewItem):
                continue
            if item.Tag == schluessel:
                item.IsSelected = True
                item.BringIntoView()
                return True
            if isinstance(item.Tag, tuple) and schluessel in item.Tag[1]:
                item.IsExpanded = True
                if self._markiere_im_baum(item.Items, schluessel):
                    return True
        return False

    # Tabelle
    def _baue_tabelle(self, sichtbar):
        tabelle = DataTable()
        for name, typ in ((u"Schluessel", String), (u"Name", String),
                          (u"Status", String), (u"StatusNr", Double),
                          (u"Klasse", String), (u"Abstand", Double),
                          (u"AbstandText", String), (u"Test", String),
                          (u"Ebene", String), (u"Elemente", String)):
            tabelle.Columns.Add(name, clr.GetClrType(typ))
        eintraege = self._eintraege()
        for clash in sichtbar:
            eintrag = eintraege.get(clash.schluessel)
            status = lg.status(clash, eintrag)
            zeile = tabelle.NewRow()
            zeile[u"Schluessel"] = clash.schluessel
            zeile[u"Name"] = clash.name
            zeile[u"Status"] = lg.status_text(status)
            zeile[u"StatusNr"] = float(br.STATUS.index(status)
                                       if status in br.STATUS else 9)
            zeile[u"Klasse"] = lg.klasse(eintrag)
            zeile[u"Abstand"] = clash.abstand if clash.abstand is not None \
                else 0.0
            zeile[u"AbstandText"] = lg.abstand_text(clash.abstand)
            zeile[u"Test"] = clash.test
            zeile[u"Ebene"] = clash.ebene
            zeile[u"Elemente"] = u"  \u2194  ".join(
                lg.objekt_text(o) for o in clash.objekte)
            tabelle.Rows.Add(zeile)
        if self._tabelle is not None:
            tabelle.DefaultView.Sort = self._tabelle.DefaultView.Sort
        self._tabelle = tabelle
        self._liste.ItemsSource = tabelle.DefaultView
        self.reihenfolge = [zeile[u"Schluessel"]
                            for zeile in tabelle.DefaultView]

    def _sortiere(self, _sender, args):
        kopf = args.OriginalSource
        if not isinstance(kopf, GridViewColumnHeader) or kopf.Tag is None \
                or self._tabelle is None:
            return
        spalte = u"%s" % kopf.Tag
        ansicht = self._tabelle.DefaultView
        richtung = u"DESC" if ansicht.Sort == spalte + u" ASC" else u"ASC"
        ansicht.Sort = u"%s %s" % (spalte, richtung)
        self.reihenfolge = [zeile[u"Schluessel"] for zeile in ansicht]

    # --- Auswahl und Sprung -----------------------------------------------
    def _clash(self, schluessel):
        if self.bericht is None or not schluessel:
            return None
        for clash in self.bericht.clashes:
            if clash.schluessel == schluessel:
                return clash
        return None

    def _baum_auswahl(self, _sender, _args):
        item = self._baum.SelectedItem
        if not isinstance(item, TreeViewItem):
            return
        if isinstance(item.Tag, tuple):
            anzahl = len(item.Tag[1])
            self._bewerten_hinweis.Text = t(
                u"Zweig gewählt: Anwenden gilt für alle %d Kollisionen.",
                u"Branch selected: apply affects all %d clashes.",
                u"Rama seleccionada: aplicar afecta a los %d conflictos.") \
                % anzahl
            return
        self._gewaehlt(item.Tag)

    def _liste_auswahl(self, _sender, _args):
        gewaehlt = list(self._liste.SelectedItems)
        if len(gewaehlt) == 1:
            self._gewaehlt(gewaehlt[0][u"Schluessel"])
        elif len(gewaehlt) > 1:
            self._bewerten_hinweis.Text = t(
                u"%d Kollisionen gewählt - Anwenden gilt für alle.",
                u"%d clashes selected - apply affects all.",
                u"%d conflictos seleccionados: aplicar afecta a todos.") \
                % len(gewaehlt)

    def _liste_doppelklick(self, _sender, _args):
        clash = self._clash(self.aktuell)
        if clash is not None and self._liste.SelectedItems.Count == 1:
            self.springe(clash)

    def _gewaehlt(self, schluessel):
        clash = self._clash(schluessel)
        if clash is None:
            return
        if schluessel != self.analyse_schluessel and self.analyse is not None:
            self._loesung_zuruecksetzen(vorschau_loeschen=True)
        self.aktuell = schluessel
        eintrag = self._eintraege().get(schluessel) or {}
        self._klasse.Text = eintrag.get(u"klasse") or u""
        self._kommentar.Text = eintrag.get(u"kommentar") or u""
        self._bewerten_hinweis.Text = u""
        self._zeige_details(clash)
        if not self._still:
            self.springe(clash)

    def _erneut_springen(self, args, clash):
        args.Handled = True
        self.springe(clash)

    def springe(self, clash):
        from clash_navigator import revit as rv
        self._speichere_einstellungen()
        optionen = dict(self.einstellungen)

        def auftrag(uiapp):
            self._hole_modellnamen(uiapp)
            try:
                ergebnis = rv.gehe_zu(uiapp, clash, optionen)
            except rv.Fehler as fehler:
                self.log(u"%s: %s" % (clash.name, fehler))
                self._bewerten_hinweis.Text = u"%s" % fehler
                return
            text = t(u"%s: %d von %d Elementen gefunden",
                     u"%s: %d of %d elements found",
                     u"%s: %d de %d elementos encontrados") % (
                clash.name, ergebnis.gefunden, ergebnis.gesucht)
            if ergebnis.deutung and clash.punkt is not None:
                text += u" · " + dict(
                    (code, tt(texte)) for code, texte
                    in KOORDINATEN)[ergebnis.deutung]
            self.log(text)
            for zeile in ergebnis.meldungen:
                self.log(u"   " + zeile)

        self.an_revit(u"springen", auftrag)

    def schritt(self, richtung):
        if not self.reihenfolge:
            return
        if self.aktuell in self.reihenfolge:
            index = self.reihenfolge.index(self.aktuell) + richtung
        else:
            index = 0 if richtung > 0 else len(self.reihenfolge) - 1
        if 0 <= index < len(self.reihenfolge):
            self._markiere(self.reihenfolge[index])

    def _zeige_details(self, clash):
        if clash is None:
            self._details.Text = t(u"Keine Kollision gewählt.",
                                   u"No clash selected.",
                                   u"Ningún conflicto seleccionado.")
            return
        eintrag = self._eintraege().get(clash.schluessel) or {}
        zeilen = [clash.name]
        if clash.test:
            zeilen.append(t(u"Test: %s", u"Test: %s", u"Prueba: %s")
                          % clash.test)
        if clash.gruppe:
            zeilen.append(t(u"Gruppe: %s", u"Group: %s", u"Grupo: %s")
                          % clash.gruppe)
        status_navis = lg.status_text(clash.status_navis)
        status = lg.status_text(lg.status(clash, eintrag))
        zeilen.append(t(u"Status: %s (Navisworks: %s)",
                        u"Status: %s (Navisworks: %s)",
                        u"Estado: %s (Navisworks: %s)")
                      % (status, status_navis))
        if clash.abstand is not None:
            zeilen.append(t(u"Abstand: %s", u"Distance: %s",
                            u"Distancia: %s")
                          % lg.abstand_text(clash.abstand))
        if clash.raster:
            zeilen.append(t(u"Raster: %s", u"Grid: %s", u"Rejilla: %s")
                          % clash.raster)
        for nummer, objekt in enumerate(clash.objekte):
            zeilen.append(u"%s: %s%s" % (
                u"AB"[nummer] if nummer < 2 else nummer + 1,
                lg.objekt_text(objekt),
                u"  (%s)" % objekt.datei if objekt.datei else u""))
        for kommentar in clash.kommentare_navis:
            zeilen.append(u"Navisworks: %s" % kommentar)
        if eintrag.get(u"von"):
            zeilen.append(t(u"Bewertet von %s, %s",
                            u"Reviewed by %s, %s",
                            u"Revisado por %s, %s")
                          % (eintrag[u"von"], eintrag.get(u"zeit", u"")))
        self._details.Text = u"\n".join(zeilen)

    # --- Lösen (Umgehung) -------------------------------------------------
    def _loesen_optionen(self):
        self._speichere_einstellungen()
        optionen = dict(self.einstellungen)
        winkel = []
        if self._winkel45.IsChecked:
            winkel.append(45)
        if self._winkel90.IsChecked:
            winkel.append(90)
        optionen[u"winkel"] = winkel or [45]
        return optionen

    def berechne(self):
        from clash_navigator import revit as rv
        from clash_navigator import umgehung as ug
        from clash_navigator import umgehung_revit as uv
        clash = self._clash(self.aktuell)
        if clash is None:
            self._loesen_info.Text = t(u"Zuerst eine Kollision wählen.",
                                       u"Select a clash first.",
                                       u"Primero seleccione un conflicto.")
            return
        optionen = self._loesen_optionen()
        self._loesung_zuruecksetzen()
        self._loesen_info.Text = t(u"Berechne...", u"Calculating...",
                                   u"Calculando...")
        # Gesperrt bis zum Ende - sonst stapeln sich Berechnungen
        self._berechnen.IsEnabled = False

        def auftrag(uiapp):
            beginn = time.time()
            try:
                analyse = uv.analysiere(uiapp, clash, optionen)
            except (rv.Fehler, ug.Fehler) as fehler:
                self._loesen_info.Text = u"%s" % fehler
                self.log(u"%s: %s" % (clash.name, fehler))
                return
            finally:
                self._berechnen.IsEnabled = True
            self.log(t(u"%s: %d Varianten in %.1f s geprüft",
                       u"%s: %d options checked in %.1f s",
                       u"%s: %d variantes comprobadas en %.1f s") % (
                clash.name, len(analyse.varianten), time.time() - beginn))
            self.analyse = analyse
            self.analyse_schluessel = clash.schluessel
            if len(analyse.leitungen) > 1:
                zeilen = [u"%s: %s" % (l.kennung, l.text)
                          for l in analyse.leitungen]
            else:
                leitung = analyse.leitungen[0]
                zeilen = [t(u"Verlegt wird: %s", u"Rerouted: %s",
                            u"Se desvía: %s") % leitung.text,
                          t(u"Hindernis: %s", u"Obstacle: %s",
                            u"Obstáculo: %s") % leitung.hindernis_text]
            zeilen.extend(analyse.hinweise)
            self._loesen_info.Text = u"\n".join(zeilen)
            self._zeige_karten()
            for nummer, variante in enumerate(analyse.varianten):
                if variante.gueltig:
                    self._markiere_karte(nummer)
                    uv.zeige_vorschau(uiapp, analyse, variante)
                    break
            else:
                uv.entferne_vorschau(uiapp)
            self._markiere_unloesbar(clash, analyse)

        self.an_revit(u"loesen", auftrag)
        if self.ereignis is None:
            self._berechnen.IsEnabled = True

    def _markiere_unloesbar(self, clash, analyse):
        """Passt keine Variante kollisionsfrei, bekommt die Kollision die
        Klasse MANUELL und den Grund als Kommentar - für die Handarbeit bzw.
        die Koordination. Über "Rückgängig" zurückzunehmen."""
        if self.speicher is None:
            return
        if any(v.gueltig and not v.kollisionen for v in analyse.varianten):
            return
        gruende = []
        for variante in analyse.varianten:
            grund = variante.grund if not variante.gueltig else (
                t(u"Kollision mit %s", u"Clash with %s",
                  u"Conflicto con %s") % variante.kollisionen[0])
            if grund and grund not in gruende:
                gruende.append(grund)
        text = t(u"Automatisch nicht lösbar: %s",
                 u"Not resolvable automatically: %s",
                 u"No se puede resolver automáticamente: %s") % (
            u"; ".join(gruende[:3]) or u"-")
        alt = (self._eintraege().get(clash.schluessel) or {}).get(
            u"kommentar") or u""
        if text in alt:
            return                  # schon beim letzten Berechnen vermerkt
        self.speicher.anwenden(
            [clash.schluessel],
            {u"klasse": lg.KLASSE_MANUELL,
             u"kommentar": (alt + u"\n" + text).strip()},
            t(u"%s: als nicht lösbar markiert",
              u"%s: marked as not resolvable",
              u"%s: marcado como no resoluble") % clash.name)
        self.log(u"%s: %s" % (clash.name, text))
        self._loesen_info.Text += u"\n" + t(
            u"Als MANUELL markiert (Rückgängig möglich).",
            u"Marked as MANUELL (can be undone).",
            u"Marcado como MANUELL (se puede deshacer).")
        self._zeige_rueckgaengig()
        self.aktualisiere()

    def _loesung_zuruecksetzen(self, vorschau_loeschen=False):
        hatte_vorschau = self.analyse is not None
        self.analyse = None
        self.analyse_schluessel = None
        self.variante = None
        self._karten.Children.Clear()
        self._loesen_info.Text = u""
        self._uebernehmen.IsEnabled = False
        if vorschau_loeschen and hatte_vorschau:
            from clash_navigator import umgehung_revit as uv
            self.an_revit(u"vorschau", uv.entferne_vorschau)

    def _zeige_karten(self):
        self._karten.Children.Clear()
        for nummer, variante in enumerate(self.analyse.varianten):
            self._karten.Children.Add(self._karte(nummer, variante))

    def _karte(self, nummer, variante):
        if not variante.gueltig:
            farbe, zustand = GRAU, variante.grund
        elif variante.kollisionen:
            farbe = STATUSFARBEN[br.AKTIV]
            zustand = t(u"✗ %d Kollision(en): %s",
                        u"✗ %d clash(es): %s",
                        u"✗ %d conflicto(s): %s") % (
                len(variante.kollisionen),
                u", ".join(variante.kollisionen[:2]))
        else:
            farbe = STATUSFARBEN[br.GENEHMIGT]
            zustand = t(u"✓ kollisionsfrei", u"✓ clash-free",
                        u"✓ sin conflictos")

        inhalt = StackPanel()
        kopf = TextBlock()
        if len(self.analyse.leitungen) > 1:
            # Welches Element bewegt sich? A = kleinere, B = grössere Leitung
            kopf.Text = u"%d  %s · %s" % (nummer + 1,
                                          variante.leitung.kennung,
                                          variante.name)
            kopf.ToolTip = variante.leitung.text
        else:
            kopf.Text = u"%d  %s" % (nummer + 1, variante.name)
        kopf.FontWeight = FontWeights.SemiBold
        inhalt.Children.Add(kopf)
        inhalt.Children.Add(self._skizze(variante, farbe))
        warnfarbe = STATUSFARBEN[br.AKTIV]
        zeilen = [(variante.zusammenfassung() if variante.gueltig else u"",
                   GRAU), (zustand, farbe)]
        if variante.gueltig:
            for text, kritisch in variante.warnungen:
                zeilen.append((u"⚠ " + text, warnfarbe) if kritisch
                              else (text, GRAU))
        for text, textfarbe in zeilen:
            if not text:
                continue
            zeile = TextBlock()
            zeile.Text = text
            zeile.FontSize = 11.0
            zeile.TextWrapping = TextWrapping.Wrap
            zeile.Foreground = textfarbe
            inhalt.Children.Add(zeile)

        karte = Border()
        karte.Child = inhalt
        karte.BorderThickness = Thickness(1.0)
        karte.BorderBrush = RAHMEN
        karte.CornerRadius = CornerRadius(3.0)
        karte.Padding = Thickness(6.0, 4.0, 6.0, 4.0)
        karte.Margin = Thickness(0.0, 0.0, 0.0, 5.0)
        karte.Background = WEISS
        karte.Tag = nummer
        if variante.gueltig:
            karte.Cursor = Cursors.Hand
            karte.MouseLeftButtonUp += self._sicher(
                lambda sender, _a: self.waehle_variante(sender.Tag))
        else:
            karte.Opacity = 0.6
        return karte

    @staticmethod
    def _skizze(variante, farbe, breite=200.0, hoehe=54.0):
        """Seitenansicht der Umgehung: Hindernis, alte Lage, neuer Weg."""
        daten = variante.skizze
        pfad = daten[u"pfad"]
        halb = daten[u"halb"]
        t0, t1 = daten[u"hindernis_t"]
        d0, d1 = daten[u"hindernis_d"]
        links = min(min(p[0] for p in pfad[1:-1]), t0)
        rechts = max(max(p[0] for p in pfad[1:-1]), t1)
        # Liegt ein Ende des Stücks in der Nähe, gehört es ins Bild - so sieht
        # man, wenn die Umgehung nicht mehr auf das Stück passt
        spanne = rechts - links
        if links - spanne <= 0.0:
            links = min(links, 0.0)
        if rechts + spanne >= daten[u"laenge"]:
            rechts = max(rechts, daten[u"laenge"])
        rand = (rechts - links) * 0.15 + halb
        links, rechts = links - rand, rechts + rand
        unten = min(0.0, d0) - halb * 2.0
        oben = max(variante.versatz, d1) + halb * 2.0
        massstab = min(breite / max(rechts - links, 1e-6),
                       hoehe / max(oben - unten, 1e-6))

        def punkt(t_wert, d_wert):
            return Point((t_wert - links) * massstab,
                         hoehe - (d_wert - unten) * massstab)

        leinwand = Canvas()
        leinwand.Width = breite
        leinwand.Height = hoehe
        leinwand.ClipToBounds = True
        leinwand.Margin = Thickness(0.0, 3.0, 0.0, 3.0)

        klotz = Rectangle()
        ecke = punkt(t0, d1)
        klotz.Width = max((t1 - t0) * massstab, 2.0)
        klotz.Height = max((d1 - d0) * massstab, 2.0)
        klotz.Fill = HINDERNIS
        Canvas.SetLeft(klotz, ecke.X)
        Canvas.SetTop(klotz, ecke.Y)
        leinwand.Children.Add(klotz)

        dicke = min(max(halb * 2.0 * massstab, 2.0), 10.0)
        # Heutige Lage: gestrichelt, und nur so weit, wie das Stück wirklich
        # reicht - mit Strich an den Enden (dort sitzt ein Bogen o.ä.)
        laenge = daten[u"laenge"]
        alt = Polyline()
        alt.Points = PointCollection()
        alt.Points.Add(punkt(max(links, 0.0), 0.0))
        alt.Points.Add(punkt(min(rechts, laenge), 0.0))
        alt.Stroke = GRAU
        alt.StrokeThickness = 1.5
        striche = DoubleCollection()
        striche.Add(4.0)
        striche.Add(3.0)
        alt.StrokeDashArray = striche
        leinwand.Children.Add(alt)
        for ende in (0.0, laenge):
            if links <= ende <= rechts:
                strich = Polyline()
                strich.Points = PointCollection()
                strich.Points.Add(punkt(ende, -halb * 2.0))
                strich.Points.Add(punkt(ende, halb * 2.0))
                strich.Stroke = DUNKEL
                strich.StrokeThickness = 2.0
                leinwand.Children.Add(strich)

        neu = Polyline()
        neu.Points = PointCollection()
        for t_wert, d_wert in pfad:
            neu.Points.Add(punkt(min(max(t_wert, links), rechts), d_wert))
        neu.Stroke = farbe
        neu.StrokeThickness = dicke
        leinwand.Children.Add(neu)
        return leinwand

    def _markiere_karte(self, nummer):
        self.variante = self.analyse.varianten[nummer]
        for karte in self._karten.Children:
            gewaehlt = karte.Tag == nummer
            karte.BorderBrush = AKZENT if gewaehlt else RAHMEN
            karte.BorderThickness = Thickness(2.0 if gewaehlt else 1.0)
        self._uebernehmen.IsEnabled = self.variante.gueltig

    def waehle_variante(self, nummer):
        from clash_navigator import umgehung_revit as uv
        if self.analyse is None:
            return
        self._markiere_karte(nummer)
        analyse, variante = self.analyse, self.variante
        self.an_revit(u"vorschau", lambda uiapp: uv.zeige_vorschau(
            uiapp, analyse, variante))

    def entferne_vorschau(self):
        from clash_navigator import umgehung_revit as uv
        self._loesung_zuruecksetzen()
        self.an_revit(u"vorschau", uv.entferne_vorschau)

    def uebernehme(self):
        from clash_navigator import revit as rv
        from clash_navigator import umgehung_revit as uv
        if self.analyse is None or self.variante is None:
            return
        analyse, variante = self.analyse, self.variante
        schluessel = self.analyse_schluessel
        if variante.kollisionen and not frage(t(
                u"Die Variante hat noch %d Kollision(en):\n%s\n\n"
                u"Trotzdem übernehmen?",
                u"The option still has %d clash(es):\n%s\n\nApply anyway?",
                u"La variante aún tiene %d conflicto(s):\n%s\n\n"
                u"¿Aplicar de todos modos?") % (
                    len(variante.kollisionen),
                    u"\n".join(variante.kollisionen[:5]))):
            return

        def auftrag(uiapp):
            try:
                anzahl = uv.uebernehme(uiapp, analyse, variante)
            except rv.Fehler as fehler:
                self.log(u"%s" % fehler)
                meldung(u"%s" % fehler, warnung=True)
                return
            text = t(u"Automatisch gelöst: %s, Versatz %d cm",
                     u"Resolved automatically: %s, offset %d cm",
                     u"Resuelto automáticamente: %s, desvío %d cm") % (
                variante.name, round(variante.versatz * 100.0))
            self.log(u"%s (%d %s)" % (text, anzahl, t(
                u"neue Elemente", u"new elements", u"elementos nuevos")))
            self._loesung_zuruecksetzen()
            if self.speicher is not None:
                alt = (self._eintraege().get(schluessel) or {}).get(
                    u"kommentar") or u""
                self.speicher.anwenden(
                    [schluessel],
                    {u"status": br.BEHOBEN,
                     u"kommentar": (alt + u"\n" + text).strip()},
                    text)
                self._zeige_rueckgaengig()
                self.aktualisiere()
            self._loesen_info.Text = text + u"\n" + t(
                u"Strg+Z in Revit nimmt die Änderung zurück.",
                u"Ctrl+Z in Revit undoes the change.",
                u"Ctrl+Z en Revit deshace el cambio.")

        self.an_revit(u"uebernehmen", auftrag)

    # --- Bewerten ---------------------------------------------------------
    def _ziel_schluessel(self):
        """Die Kollisionen, für die Anwenden gilt."""
        if self._tabelle_an.IsChecked:
            return [zeile[u"Schluessel"] for zeile
                    in list(self._liste.SelectedItems)]
        item = self._baum.SelectedItem
        if isinstance(item, TreeViewItem):
            if isinstance(item.Tag, tuple):
                return [k for k in self.reihenfolge if k in item.Tag[1]]
            return [item.Tag]
        return []

    def _aenderung_aus_formular(self):
        aenderung = {u"klasse": (self._klasse.Text or u"").strip(),
                     u"kommentar": (self._kommentar.Text or u"").strip()}
        if self._status_aendern.IsChecked:
            aenderung[u"status"] = self._code(self._status)
        return aenderung

    def anwenden(self):
        self._wende_an(self._aenderung_aus_formular())

    def wie_vorher(self):
        if not self.letzte_aenderung:
            self._bewerten_hinweis.Text = t(
                u"Noch nichts angewendet.", u"Nothing applied yet.",
                u"Todavía no se ha aplicado nada.")
            return
        self._wende_an(dict(self.letzte_aenderung))

    def _beschreibung(self, aenderung, anzahl):
        teile = []
        if aenderung.get(u"klasse"):
            teile.append(t(u"Klasse \u201e%s\u201c", u"Class \"%s\"",
                           u"Clase \"%s\"") % aenderung[u"klasse"])
        if aenderung.get(u"status"):
            teile.append(u"Status \u2192 %s"
                         % lg.status_text(aenderung[u"status"]))
        if aenderung.get(u"kommentar"):
            teile.append(t(u"Kommentar", u"Comment", u"Comentario"))
        if not teile:
            teile.append(t(u"Bewertung geleert", u"Review cleared",
                           u"Revisión borrada"))
        return t(u"%s · %d Kollision(en)", u"%s · %d clash(es)",
                 u"%s · %d conflicto(s)") % (u" + ".join(teile), anzahl)

    def _wende_an(self, aenderung):
        if self.speicher is None:
            self._bewerten_hinweis.Text = t(
                u"Zuerst einen Bericht laden.", u"Load a report first.",
                u"Primero cargue un informe.")
            return
        ziele = self._ziel_schluessel()
        if not ziele:
            self._bewerten_hinweis.Text = t(
                u"Zuerst eine Kollision wählen.", u"Select a clash first.",
                u"Primero seleccione un conflicto.")
            return
        if len(ziele) > 1 and not frage(t(
                u"Bewertung auf %d Kollisionen anwenden?",
                u"Apply the review to %d clashes?",
                u"¿Aplicar la revisión a %d conflictos?") % len(ziele)):
            return

        # Nächste Kollision vor dem Neuaufbau merken - danach kann die
        # aktuelle in einen anderen Zweig gewandert oder ausgeblendet sein
        naechste = None
        if len(ziele) == 1 and self._weiter.IsChecked and \
                ziele[0] in self.reihenfolge:
            index = self.reihenfolge.index(ziele[0])
            if index + 1 < len(self.reihenfolge):
                naechste = self.reihenfolge[index + 1]

        beschreibung = self._beschreibung(aenderung, len(ziele))
        try:
            anzahl = self.speicher.anwenden(ziele, aenderung, beschreibung)
        except (IOError, OSError) as fehler:
            meldung(t(u"Die Bewertung ließ sich nicht speichern:\n%s",
                      u"The review could not be saved:\n%s",
                      u"No se pudo guardar la revisión:\n%s") % fehler,
                    warnung=True)
            return
        self.letzte_aenderung = dict(aenderung)
        self._merke_klasse(aenderung.get(u"klasse"))
        if anzahl:
            self.log(beschreibung)
        self._bewerten_hinweis.Text = (
            beschreibung if anzahl else t(u"Keine Änderung.", u"No change.",
                                          u"Sin cambios."))
        self._zeige_rueckgaengig()
        if naechste:
            # Nicht erst die bewertete Kollision wieder markieren - sonst
            # klappt ihr neuer Zweig (z.B. "Geprüft") unnötig auf
            self.aktuell = None
        self.aktualisiere()
        if naechste:
            self._markiere(naechste)

    def _merke_klasse(self, klasse):
        if not klasse or klasse in self._klassen():
            return
        self.einstellungen[u"klassen"] = list(
            self.einstellungen.get(u"klassen") or []) + [klasse]
        self._speichere_einstellungen()
        self._fuelle_klassen()

    def rueckgaengig(self):
        if self.speicher is None:
            return
        beschreibung = self.speicher.rueckgaengig()
        if beschreibung is None:
            return
        self.log(t(u"Rückgängig: %s", u"Undone: %s", u"Deshecho: %s")
                 % beschreibung)
        self._zeige_rueckgaengig()
        self.aktualisiere()

    def _zeige_rueckgaengig(self):
        schritt = self.speicher.letzter_schritt if self.speicher else None
        self._rueckgaengig.IsEnabled = schritt is not None
        if schritt is None:
            self._undo_titel.Text = t(u"Nichts rückgängig zu machen.",
                                      u"Nothing to undo.",
                                      u"Nada que deshacer.")
            self._undo_text.Text = u""
            self._undo_rest.Text = u""
            return
        self._undo_titel.Text = u"%s · %s" % (schritt[u"beschreibung"],
                                              schritt[u"zeit"])
        vorher = [eintrag for eintrag in schritt[u"vorher"].values()]
        leer = sum(1 for eintrag in vorher if not eintrag)
        self._undo_text.Text = t(
            u"Vorher: %d ohne Bewertung, %d bewertet.",
            u"Before: %d without review, %d reviewed.",
            u"Antes: %d sin revisión, %d revisados.") % (
                leer, len(vorher) - leer)
        weitere = len(self.speicher.stapel) - 1
        self._undo_rest.Text = (t(u"  %d weitere Schritte danach",
                                  u"  %d more step(s) after this one",
                                  u"  %d paso(s) más después")
                                % weitere) if weitere else u""


# ---------------------------------------------------------------------------
# Registrierung (aus startup.py)
# ---------------------------------------------------------------------------

def registriere():
    """Panel bei Revit anmelden und das ExternalEvent anlegen.

    Muss beim Start laufen: Revit nimmt andockbare Panels nur während des
    Hochfahrens an, und ExternalEvent.Create braucht den API-Kontext.
    """
    if forms.is_registered_dockable_panel(ClashPanel):
        return None
    panel = forms.register_dockable_panel(ClashPanel, default_visible=False)
    panel.ereignis = ExternalEvent.Create(panel.auftraege)
    return panel
