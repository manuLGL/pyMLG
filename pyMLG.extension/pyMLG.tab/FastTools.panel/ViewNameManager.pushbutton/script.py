# -*- coding: utf-8 -*-
"""
View Name Manager with Preview
"""

__title__ = "View Name\nManager"
__author__ = "Manuel"

from pyrevit import revit, DB, forms

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
        preview_lines.append("\n... und {} weitere Ansichten".format(len(views) - 20))

    preview_text = "\n\n".join(preview_lines)

    return forms.alert(
        "Vorschau der Aenderungen:\n\n{}\n\nMoechtest du fortfahren?".format(preview_text),
        title="Vorschau ({} Ansichten)".format(len(views)),
        yes=True,
        no=True
    )


def main():
    try:
        # Get selected views from Project Browser
        selected_views = get_selected_views()

        if not selected_views:
            forms.alert(
                "Keine Ansichten ausgewaehlt!\n\n"
                "Bitte waehle Ansichten im Projektbrowser aus und fuehre das Tool erneut aus."
            )
            return

        # Ask for operation type
        ops = [
            "1. Praefix hinzufuegen",
            "2. Suffix hinzufuegen",
            "3. Namen ersetzen (nutze {name} als Platzhalter)"
        ]

        op_choice = forms.CommandSwitchWindow.show(
            ops,
            message="{} Ansichten ausgewaehlt\n\nWas moechtest du tun?".format(
                len(selected_views)
            )
        )

        if not op_choice:
            return

        # Determine operation
        if "Praefix" in op_choice:
            operation = "prefix"
            prompt = "Praefix eingeben:"
            default = "NEW_"
        elif "Suffix" in op_choice:
            operation = "suffix"
            prompt = "Suffix eingeben:"
            default = "_NEW"
        else:
            operation = "replace"
            prompt = "Neuen Namen eingeben (nutze {name} fuer aktuellen Namen):"
            default = "{name}_Kopie"

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

        with revit.Transaction("Ansichten umbenennen"):
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
        msg = "Erfolgreich {} Ansichten umbenannt!\n\nRueckgaengig mit Strg+Z".format(renamed)
        if errors:
            msg += "\n\nFehler ({}):\n{}".format(
                len(errors),
                "\n".join(errors[:5])
            )
            if len(errors) > 5:
                msg += "\n... und {} weitere Fehler".format(len(errors) - 5)

        forms.alert(msg, title="Fertig")

    except Exception as e:
        import traceback
        forms.alert("Fehler:\n{}\n\n{}".format(str(e), traceback.format_exc()))


if __name__ == '__main__':
    main()