#crash test2
# IMPORT
import clr
import System 
import rpw
from rpw import revit, db, ui, DB, UI
from rpw.ui.forms import *
from System import Array
from System.Runtime.InteropServices import Marshal
import sys
import Autodesk

# # FONCTION DUPLICATE VIEW
# def DuplicateView(view, name, doc):
#   try:
#     newViewId = view.Duplicate(Autodesk.Revit.DB.ViewDuplicateOption.Duplicate)
#     newView = doc.GetElement(newViewId)
#     try: newView.Name = name
#     except: pass
#     return newView
#   except: return None

# COLLECT VIEW MAQUETTE
view_category = db.Collector(of_category='Views')
view_elements = view_category.get_elements()
view_name = [views.Name for views in view_elements]
view_id = [views.Id for views in view_elements]
view_get_type_id = [views.LookupParameter("Type").AsElementId() for views in view_elements]



# # # DUPLICATE VIEW
# # with db.Transaction('duplicate views'):
# #   for view in view_elements:
# #     [DuplicateView(view, x, revit.doc) for x in name]

# COLLECT LEVELS MAQUETTE
level_category = db.Collector(of_category='Levels', is_not_type=True)
level_elements = level_category.get_elements()
level_name = [levels.LookupParameter("Nom").AsString() for levels in level_elements]
level_elevation = [levels.Elevation for levels in level_elements]
level_id = [levels.Id for levels in level_elements]

# FONCTION CREATE VIEW
def CreateView(doc, viewFamilyTypeId, levelId):
  newView = DB.ViewPlan.Create(doc, viewFamilyTypeId, levelId)


# CREATE VIEW
with db.Transaction('create views'):
  new_view = CreateView(revit.doc, view_get_type_id[0], level_id[0])



print view_get_type_id