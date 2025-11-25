from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document

cad_name = "A$Cc73e60c4"

print("Suche nach CAD-Links mit '{}' im Namen".format(cad_name))
print("=" * 60)

# CAD-Link-TYPES suchen
link_types = FilteredElementCollector(doc).OfClass(CADLinkType)

for link_type in link_types:
    print("CAD-Link Type: '{}'".format(link_type.Name))

    if cad_name in link_type.Name:
        print("  -> TREFFER!")

        # Jetzt die Instanzen dieses Links finden
        instances = FilteredElementCollector(doc) \
            .OfClass(ImportInstance) \
            .WhereElementIsNotElementType()

        for imp in instances:
            if imp.GetTypeId() == link_type.Id:
                view_id = imp.OwnerViewId
                print("  View ID: {}".format(view_id.IntegerValue))

                if view_id.IntegerValue > 0:
                    view = doc.GetElement(view_id)
                    print("  -> View: {}".format(view.Name))
                else:
                    print("  -> Projekt-weit")

print("=" * 60)