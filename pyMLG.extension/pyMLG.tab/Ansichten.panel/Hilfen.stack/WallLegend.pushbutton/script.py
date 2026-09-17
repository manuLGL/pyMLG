# -*- coding: utf-8 -*-
"""Crear Secciones por Tipo de Muro
Crea automaticamente una seccion por cada tipo de muro usado en el proyecto
"""
__title__ = 'WallLegend'
__author__ = 'Manuel'

from pyrevit import revit, DB, forms, script

doc = revit.doc
uidoc = revit.uidoc

# Recopilar todos los muros del proyecto
all_walls_collector = DB.FilteredElementCollector(doc) \
    .OfClass(DB.Wall) \
    .WhereElementIsNotElementType()

all_walls = list(all_walls_collector)

if not all_walls:
    forms.alert('No hay muros en el proyecto', exitscript=True)

print('\nTotal de muros en proyecto: {}'.format(len(all_walls)))

# Agrupar muros por tipo
wall_types_dict = {}

for wall in all_walls:
    try:
        wall_type_id = wall.GetTypeId()
        type_id_int = wall_type_id.IntegerValue

        if type_id_int not in wall_types_dict:
            wall_type = doc.GetElement(wall_type_id)

            if wall_type:
                # Obtener nombre del tipo de forma segura
                try:
                    type_name = wall_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                except:
                    type_name = "Tipo_{}".format(type_id_int)

                wall_types_dict[type_id_int] = {
                    'type': wall_type,
                    'type_name': type_name,
                    'walls': []
                }

        wall_types_dict[type_id_int]['walls'].append(wall)

    except:
        continue

if not wall_types_dict:
    forms.alert('No se pudieron procesar los tipos de muro', exitscript=True)

# Mostrar tipos encontrados
print('\n' + '=' * 70)
print('TIPOS DE MURO ENCONTRADOS:')
print('=' * 70)

for type_id, data in wall_types_dict.items():
    print('\n- {}'.format(data['type_name']))
    print('  Cantidad: {} muros'.format(len(data['walls'])))

# Confirmar
msg = 'Se encontraron {} tipos de muro diferentes.\n\n'.format(len(wall_types_dict))
msg += 'Se creara UNA seccion por cada tipo.\n\nContinuar?'

if not forms.alert(msg, yes=True, no=True):
    script.exit()

# Buscar tipo de seccion
section_type_id = None
collector = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)

for vft in collector:
    if vft.ViewFamily == DB.ViewFamily.Section:
        section_type_id = vft.Id
        break

if not section_type_id:
    forms.alert('No se encontro tipo de vista de seccion', exitscript=True)

# Procesar cada tipo de muro
output = script.get_output()
created_sections = []
errors = []

with revit.Transaction('Crear Secciones por Tipo de Muro'):
    for type_id, data in wall_types_dict.items():
        try:
            wall_type_name = data['type_name']
            representative_wall = data['walls'][0]
            wall_id = representative_wall.Id.IntegerValue

            print('\n' + '-' * 70)
            print('Procesando: {}'.format(wall_type_name))
            print('Muro ID: {}'.format(wall_id))

            # Datos del muro
            loc = representative_wall.Location
            curve = loc.Curve
            p1 = curve.GetEndPoint(0)
            p2 = curve.GetEndPoint(1)

            # Punto medio
            mid_x = (p1.X + p2.X) / 2.0
            mid_y = (p1.Y + p2.Y) / 2.0
            mid_z = (p1.Z + p2.Z) / 2.0
            midpoint = DB.XYZ(mid_x, mid_y, mid_z)

            # Direccion del muro
            dx = p2.X - p1.X
            dy = p2.Y - p1.Y
            wall_length = (dx * dx + dy * dy) ** 0.5

            if wall_length > 0:
                wall_dir_x = dx / wall_length
                wall_dir_y = dy / wall_length
            else:
                wall_dir_x = 1.0
                wall_dir_y = 0.0

            wall_direction = DB.XYZ(wall_dir_x, wall_dir_y, 0)
            view_direction = DB.XYZ(-wall_dir_y, wall_dir_x, 0)

            # Altura
            height = 10.0
            try:
                h_param = representative_wall.get_Parameter(DB.BuiltInParameter.WALL_USER_HEIGHT_PARAM)
                if h_param:
                    h = h_param.AsDouble()
                    if h > 0:
                        height = h
            except:
                pass

            # Ancho
            width = representative_wall.WallType.Width

            # BoundingBox
            bbox = DB.BoundingBoxXYZ()

            transform = DB.Transform.Identity
            transform.Origin = midpoint
            transform.BasisX = view_direction
            transform.BasisY = DB.XYZ.BasisZ
            transform.BasisZ = wall_direction

            bbox.Transform = transform

            # Dimensiones
            margin = 1.0
            cut_depth = width / 2.0 + 0.2
            view_width = height + margin * 2

            bbox.Min = DB.XYZ(-cut_depth, -margin, -view_width / 2)
            bbox.Max = DB.XYZ(cut_depth, height + margin, view_width / 2)

            # Crear seccion
            section = DB.ViewSection.CreateSection(doc, section_type_id, bbox)
            section.Scale = 10
            section.DetailLevel = DB.ViewDetailLevel.Fine

            # Nombre limpio
            clean_name = wall_type_name.replace('/', '-').replace('\\', '-').replace(':', '-')
            clean_name = clean_name.replace(' ', '_')

            if len(clean_name) > 40:
                clean_name = clean_name[:40]

            # Asignar nombre
            counter = 1
            base_name = "Seccion_{}".format(clean_name)
            section_name = base_name

            # Verificar nombres duplicados
            all_view_names = [v.Name for v in DB.FilteredElementCollector(doc).OfClass(DB.View)]

            while section_name in all_view_names:
                section_name = "{}_{}".format(base_name, counter)
                counter += 1

            section.Name = section_name

            print('Seccion creada: {}'.format(section.Name))
            print('Ancho: {:.0f} mm'.format(width * 304.8))

            created_sections.append({
                'section': section,
                'wall': representative_wall,
                'type_name': wall_type_name,
                'wall_id': wall_id,
                'count': len(data['walls'])
            })

        except Exception as e:
            error_msg = 'Error con tipo "{}": {}'.format(data.get('type_name', 'Desconocido'), str(e))
            errors.append(error_msg)
            print('ERROR: {}'.format(error_msg))

# Aislar muros
with revit.Transaction('Aislar Muros'):
    for section_data in created_sections:
        try:
            section = section_data['section']
            wall = section_data['wall']
            wall_id = wall.Id.IntegerValue

            # Obtener elementos en vista
            collector = DB.FilteredElementCollector(doc, section.Id)
            all_ids = collector.WhereElementIsNotElementType().ToElementIds()

            # Ocultar todo excepto el muro
            to_hide = []
            for elem_id in all_ids:
                if elem_id.IntegerValue != wall_id:
                    to_hide.append(elem_id)

            if to_hide:
                section.HideElements(list(to_hide))

            # Info de capas
            try:
                wall_type = doc.GetElement(wall.GetTypeId())
                structure = wall_type.GetCompoundStructure()

                if structure:
                    layers = structure.GetLayers()

                    for i in range(layers.Count):
                        layer = layers[i]
                        width_mm = layer.Width * 304.8

                        mat_id = layer.MaterialId
                        mat_name = "Sin material"

                        if mat_id != DB.ElementId.InvalidElementId:
                            mat = doc.GetElement(mat_id)
                            if mat:
                                mat_name = mat.Name
            except:
                pass

        except:
            pass

# Resultado
output.print_md('# Resultado')
output.print_md('---')
output.print_md('**Tipos procesados:** {}'.format(len(wall_types_dict)))
output.print_md('**Secciones creadas:** {}'.format(len(created_sections)))

if errors:
    output.print_md('\n## Errores: {}'.format(len(errors)))

if created_sections:
    output.print_md('\n## Secciones')

    for idx, sd in enumerate(created_sections, 1):
        output.print_md('\n**{}. {}**'.format(idx, sd['section'].Name))
        output.print_md('- Tipo: {}'.format(sd['type_name']))
        output.print_md('- Muros de este tipo: {}'.format(sd['count']))

print('\n' + '=' * 70)
print('COMPLETADO')
print('=' * 70)