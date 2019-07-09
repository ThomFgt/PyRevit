
# """Import a list of points in order to edit the shape of a slab"""

# __title__ = 'TestAlex1'

# __doc__ = 'Needed : an Excel list of points\' coordinates (in centimeters), with for first colomn x, second y and third z'

# # -*- coding: utf-8 -*-
# e_a = str("\xe9")
# a_a = str("\xe0")

# import clr
# import math
# clr.AddReference('RevitAPI') 
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import * 
# from pyrevit import forms
# from rpw.ui.forms import (select_file, Alert, TextInput, FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CommandLink, TaskDialog)
# import rpw
# from rpw import revit, db, ui, DB, UI

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument

# options = __revit__.Application.Create.NewGeometryOptions()

# beam_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
#                 .OfCategory(BuiltInCategory.OST_StructuralFraming)\
#                 .WhereElementIsNotElementType()\
#                 .ToElements()

# level_collector = FilteredElementCollector(doc)\
#             .OfCategory(BuiltInCategory.OST_Levels)\
#             .WhereElementIsNotElementType()\
#             .ToElements()
            
# levels = {}
# for level in level_collector:
#   levels[level.Name] = level

# basepoint_category = db.Collector(of_category='OST_ProjectBasePoint', is_not_type = True)
# basepoint_elements = basepoint_category.get_elements()
# basepoint_elevation = [bp.LookupParameter("El"+e_a+"v.").AsDouble() for bp in basepoint_elements]
# basepoint = basepoint_elevation[0]/3.2808399

# components = [Label('Select the floor\'s level :'),
#           ComboBox('floor_level', levels),
#           Label('Enter the height required under the beams:'),
#           TextBox('height_required', Text="2.80"),
#           Separator(),
#           Button('OK')]
# form = FlexForm('Title', components)
# form.show()

# floor_level = round(form.values["floor_level"].Elevation/3.2808399,3)
# height_required = float(form.values["height_required"])
# # Alert('Pick a slab please (clicking on it)', title = "Select a slab", exit = False)
# # # Pick an element
# # sel = uidoc.Selection
# # obType = Selection.ObjectType.Element
# # ref = sel.PickObject(obType, "Select floor.")
# # floor = doc.GetElement(ref.ElementId)

# t = Transaction(doc, 'Tag beam')
# t.Start()

# for beam in beam_collector:
#   z_max = (beam.get_Geometry(options).GetBoundingBox().Max.Z + beam.LookupParameter("Valeur de d"+e_a+"calage"+" Z").AsDouble())/3.2808399
#   z_min = ((beam.get_Geometry(options).GetBoundingBox().Min).Z/3.2808399)
#   heightBeam = round(z_max - z_min, 2)
#   delta = z_min - floor_level - height_required
#   beam.LookupParameter("HSP").Set(z_min - floor_level + basepoint)
#     # beam.LookupParameter("HSP").Set(height_required + 0.01)



# t.Commit()






# Version 2


__title__ = 'HSP Poutres/Dalles'

__doc__ = 'Les poutres et les dalles en question doivent etre visibles dans la vue active. Attention au placement des niveaux dans le projet. Le parametre doit etre de type nombre.'




# IMPORT
import clr
import math
import System 
import rpw
from rpw import revit, db, ui, DB, UI
from rpw.ui.forms import *
from System import Array
from System.Runtime.InteropServices import Marshal
import sys
import Autodesk


e_a = str("\xe9")
a_a = str("\xe0")

options = __revit__.Application.Create.NewGeometryOptions()


floor_collector = DB.FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id)\
                .OfCategory(DB.BuiltInCategory.OST_Floors)\
                .WhereElementIsNotElementType()\
                .ToElements()


beam_collector = DB.FilteredElementCollector(revit.doc, revit.doc.ActiveView.Id)\
                .OfCategory(DB.BuiltInCategory.OST_StructuralFraming)\
                .WhereElementIsNotElementType()\
                .ToElements()

level_collector = DB.FilteredElementCollector(revit.doc)\
            .OfCategory(DB.BuiltInCategory.OST_Levels)\
            .WhereElementIsNotElementType()\
            .ToElements()
            


levels = {}
for level in level_collector:
  levels[level.Name] = level


components = [Label('Select the floor\'s level :'),
          ComboBox('floor_level', levels),
          Label('Select the element'),
          ComboBox('pick_elem', {'Poutres': 0, 'Dalles': 1}),
          Label('write the parameter'),
          TextBox('textbox'),
          Separator(),
          Button('OK')]
form = FlexForm('Your moment', components)
form.show()

floor_level = round(form.values["floor_level"].Elevation/3.2808399,3)
elem_number = form.values["pick_elem"]
parameter_name = form.values["textbox"]

if elem_number == 1:
  with db.Transaction('Tag floor'):
    for floor in floor_collector:
      z_min = ((floor.get_Geometry(options).GetBoundingBox().Min).Z/3.2808399) # + valeur point de base possiblement
      floor.LookupParameter(parameter_name).Set(z_min - floor_level)



elif elem_number == 0:
  with db.Transaction('Tag beam'):
    for beam in beam_collector:
      z_min = ((beam.get_Geometry(options).GetBoundingBox().Min).Z/3.2808399) # + valeur point de base possiblement
      beam.LookupParameter(parameter_name).Set(z_min - floor_level)