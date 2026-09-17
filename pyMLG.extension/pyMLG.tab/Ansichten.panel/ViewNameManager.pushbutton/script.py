# -*- coding: utf-8 -*-
"""
View Name Manager with Preview
"""

__title__ = "View Name\nManager"
__author__ = "Manuel"

from pyrevit import revit, DB, forms
from mlg_sprache import t

doc = revit.doc
uidoc = revit.uidoc


def get_selected_views():
    """Get views selected in Project Browser"""
    selection = uidoc.Selection.GetElementIds()
    views = []

    for elem_id in selection:
        elem = doc.GetElement(elem_id)
        if isinstance(elem, DB.View) and not elem.IsTemplate:
            views.append(elem)

    return views


def preview_changes(views, operation, text):
    """Show preview of name changes"""
    preview_lines = []

    for view in views[:20]:  # Max 20 in preview
        old_name = view.Name

        if operation == "prefix":
            new_name = text + old_name
        elif operation == "suffix":
            new_name = old_name + text
        else:  # replace
            new_name = text.replace("{name}", old_name)

        preview_lines.append("{}\n  -> {}".format(old_name, new_name))

    if len(views) > 20:
        preview_lines.append(t("\n... und {} weitere Ansichten", u"\n... and {} more views", u"\n... y {} vistas más").format(len(views) - 20))

    preview_text = "\n\n".join(preview_lines)

    return forms.alert(
        t("Vorschau der Aenderungen:\n\n{}\n\nMoechtest du fortfahren?", u"Preview of the changes:\n\n{}\n\nDo you want to continue?", u"Vista previa de los cambios:\n\n{}\n\n¿Desea continuar?").format(preview_text),
        title=t("Vorschau ({} Ansichten)", u"Preview ({} views)", u"Vista previa ({} vistas)").format(len(views)),
        yes=True,
        no=True
    )


def main():
    try:
        # Get selected views from Project Browser
        selected_views = get_selected_views()

        if not selected_views:
            forms.alert(
                t("Keine Ansichten ausgewaehlt!\n\n"
                "Bitte waehle Ansichten im Projektbrowser aus und fuehre das Tool erneut aus.", u"No views selected!\n\nPlease select views in the Project Browser and run the tool again.", u"¡No hay vistas seleccionadas!\n\nSeleccione vistas en el navegador de proyectos y ejecute de nuevo la herramienta.")
            )
            return

        # Ask for operation type
        ops = [
            t("1. Praefix hinzufuegen", u"1. Add prefix", u"1. Añadir prefijo"),
            t("2. Suffix hinzufuegen", u"2. Add suffix", u"2. Añadir sufijo"),
            t("3. Namen ersetzen (nutze {name} als Platzhalter)", u"3. Replace names (use {name} as placeholder)", u"3. Reemplazar nombres (use {name} como marcador)")
        ]

        op_choice = forms.CommandSwitchWindow.show(
            ops,
            message=t("{} Ansichten ausgewaehlt\n\nWas moechtest du tun?", u"{} views selected\n\nWhat do you want to do?", u"{} vistas seleccionadas\n\n¿Qué desea hacer?").format(
                len(selected_views)
            )
        )

        if not op_choice:
            return

        # Determine operation
        if op_choice == ops[0]:
            operation = "prefix"
            prompt = t("Praefix eingeben:", u"Enter prefix:", u"Introduzca el prefijo:")
            default = "NEW_"
        elif op_choice == ops[1]:
            operation = "suffix"
            prompt = t("Suffix eingeben:", u"Enter suffix:", u"Introduzca el sufijo:")
            default = "_NEW"
        else:
            operation = "replace"
            prompt = t("Neuen Namen eingeben (nutze {name} fuer aktuellen Namen):", u"Enter new name (use {name} for the current name):", u"Introduzca el nombre nuevo (use {name} para el nombre actual):")
            default = t("{name}_Kopie", u"{name}_Copy", u"{name}_Copia")

        # Get the text input
        text = forms.ask_for_string(
            prompt=prompt,
            default=default,
            title="View Name Manager"
        )

        if not text:
            return

        # Show preview and ask for confirmation
        if not preview_changes(selected_views, operation, text):
            return

        # Rename views
        renamed = 0
        errors = []

        with revit.Transaction(t("Ansichten umbenennen", u"Rename views", u"Renombrar vistas")):
            for view in selected_views:
                try:
                    old_name = view.Name

                    if operation == "prefix":
                        new_name = text + old_name
                    elif operation == "suffix":
                        new_name = old_name + text
                    else:  # replace
                        new_name = text.replace("{name}", old_name)

                    if new_name != old_name:
                        view.Name = new_name
                        renamed += 1

                except Exception as e:
                    errors.append("{}: {}".format(old_name, str(e)))

        # Show results
        msg = t("Erfolgreich {} Ansichten umbenannt!\n\nRueckgaengig mit Strg+Z", u"{} views renamed successfully!\n\nUndo with Ctrl+Z", u"¡{} vistas renombradas correctamente!\n\nDeshacer con Ctrl+Z").format(renamed)
        if errors:
            msg += t("\n\nFehler ({}):\n{}", u"\n\nErrors ({}):\n{}", u"\n\nErrores ({}):\n{}").format(
                len(errors),
                "\n".join(errors[:5])
            )
            if len(errors) > 5:
                msg += t("\n... und {} weitere Fehler", u"\n... and {} more errors", u"\n... y {} errores más").format(len(errors) - 5)

        forms.alert(msg, title=t("Fertig", u"Done", u"Listo"))

    except Exception as e:
        import traceback
        forms.alert(t("Fehler:\n{}\n\n{}", u"Error:\n{}\n\n{}", u"Error:\n{}\n\n{}").format(str(e), traceback.format_exc()))


if __name__ == '__main__':
    main()