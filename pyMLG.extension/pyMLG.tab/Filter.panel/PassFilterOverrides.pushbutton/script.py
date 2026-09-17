# -*- coding: utf-8 -*-
__title__ = "PassFilter\nOverrides"
__doc__ = "Überträgt grafische Filterüberschreibungen von einer Ansichtsvorlage zu einer anderen"
__author__ = "Manuel"

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import Transaction, FilteredElementCollector
from collections import defaultdict
from mlg_sprache import t

doc = revit.doc


def get_all_view_templates():
    """Sammelt alle Ansichtsvorlagen im Projekt"""
    collector = FilteredElementCollector(doc).OfClass(DB.View)
    return [v for v in collector if v.IsTemplate]


def get_filter_data(template):
    """
    Sammelt alle Filter-Daten einer Vorlage
    Returns: dict {filter_id: (filter_name, override_settings)}
    """
    filter_data = {}

    try:
        filter_ids = template.GetFilters()

        for filter_id in filter_ids:
            filter_elem = doc.GetElement(filter_id)
            if filter_elem:
                overrides = template.GetFilterOverrides(filter_id)
                filter_data[filter_id] = (filter_elem.Name, overrides)

    except Exception as e:
        print(t("Fehler beim Auslesen der Filter: {}", u"Error reading the filters: {}", u"Error al leer los filtros: {}").format(e))

    return filter_data


def copy_filter_overrides_optimized(source_template, target_templates):
    """Optimierte Version - sammelt erst alle Daten, dann eine Transaction"""

    # PHASE 1: Daten sammeln (außerhalb Transaction)
    source_filters = get_filter_data(source_template)

    if not source_filters:
        forms.alert(t(u"Keine Filter in dieser Vorlage", u"No filters in this template", u"No hay filtros en esta plantilla"), exitscript=True)

    # Ziel-Daten sammeln (welche Filter existieren bereits?)
    target_data = {}
    for target in target_templates:
        existing_filters = set(target.GetFilters())
        target_data[target.Id] = existing_filters

    # Statistik
    stats = defaultdict(int)
    errors = []

    # PHASE 2: Eine Transaction für alles
    with Transaction(doc, t("Filter Overrides übertragen", u"Transfer filter overrides", u"Transferir modificaciones de filtros")) as transaktion:
        transaktion.Start()

        try:
            for target_template in target_templates:
                existing_filters = target_data[target_template.Id]

                for filter_id, (filter_name, overrides) in source_filters.items():
                    try:
                        if filter_id in existing_filters:
                            # Filter existiert - nur Overrides updaten
                            target_template.SetFilterOverrides(filter_id, overrides)
                            stats['updated'] += 1
                        else:
                            # Filter hinzufügen + Overrides setzen
                            target_template.AddFilter(filter_id)
                            target_template.SetFilterOverrides(filter_id, overrides)
                            stats['added'] += 1

                        stats['total'] += 1

                    except Exception as e:
                        error_msg = "Filter '{}' → {}: {}".format(
                            filter_name,
                            target_template.Name,
                            str(e)
                        )
                        errors.append(error_msg)

            transaktion.Commit()

        except Exception as e:
            transaktion.RollBack()
            forms.alert(t(u"Kritischer Fehler: {}", u"Critical error: {}", u"Error crítico: {}").format(str(e)), exitscript=True)

    return stats, errors


def show_results_compact(stats, errors, source_name, target_count):
    """Kompakte Ergebnis-Anzeige"""
    output = script.get_output()

    output.print_md(t(u"# ✓ Filterüberschreibungen übertragen", u"# ✓ Filter overrides transferred", u"# ✓ Modificaciones de filtros transferidas"))
    output.print_md(t(u"**Quelle:** {}", u"**Source:** {}", u"**Origen:** {}").format(source_name))
    output.print_md(t(u"**Ziel:** {} Vorlagen", u"**Target:** {} templates", u"**Destino:** {} plantillas").format(target_count))
    output.print_md("")
    output.print_md(t(u"**Gesamt:** {} Filter-Operationen", u"**Total:** {} filter operations", u"**Total:** {} operaciones de filtro").format(stats['total']))
    output.print_md(t(u"- Hinzugefügt: {}", u"- Added: {}", u"- Añadidos: {}").format(stats['added']))
    output.print_md(t(u"- Aktualisiert: {}", u"- Updated: {}", u"- Actualizados: {}").format(stats['updated']))

    if errors:
        output.print_md("")
        output.print_md(t(u"## ⚠️ Fehler ({}):", u"## ⚠️ Errors ({}):", u"## ⚠️ Errores ({}):").format(len(errors)))
        for error in errors:
            output.print_md("- {}".format(error))


# ======================== HAUPTPROGRAMM ========================

all_templates = get_all_view_templates()

if not all_templates:
    forms.alert(t(u"Keine Ansichtsvorlagen.", u"No view templates.", u"No hay plantillas de vista."), exitscript=True)

# Dictionary mit verbesserter Anzeige
template_dict = {
    vorlage.Name: vorlage for vorlage in all_templates
}

# Quell-Vorlage
source_name = forms.SelectFromList.show(
    sorted(template_dict.keys()),
    title=t(u"Quellvorlage", u"Source template", u"Plantilla de origen"),
    button_name=t(u"Weiter", u"Continue", u"Continuar"),
    multiselect=False
)

if not source_name:
    script.exit()

source_template = template_dict[source_name]

# Ziel-Vorlagen (ohne Quelle)
target_options = [name for name in template_dict.keys() if name != source_name]

target_names = forms.SelectFromList.show(
    sorted(target_options),
    title=t(u"Zielvorlagen", u"Target templates", u"Plantillas de destino"),
    button_name=t(u"Übertragen", u"Transfer", u"Transferir"),
    multiselect=True
)

if not target_names:
    script.exit()

target_templates = [template_dict[name] for name in target_names]

# Bestätigung
if not forms.alert(
        t(u"Filterüberschreibungen von '{}' auf {} Vorlagen übertragen?", u"Transfer filter overrides from '{}' to {} templates?", u"¿Transferir las modificaciones de filtros de '{}' a {} plantillas?").format(
            source_name,
            len(target_templates)
        ),
        yes=True,
        no=True
):
    script.exit()

# Übertragung
stats, errors = copy_filter_overrides_optimized(source_template, target_templates)

# Ergebnis
show_results_compact(stats, errors, source_name, len(target_templates))