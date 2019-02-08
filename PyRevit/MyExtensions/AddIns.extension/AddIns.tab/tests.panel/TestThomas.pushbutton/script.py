# -*- coding: utf-8 -*-
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

e_a = str("\xe9")
a_a = str("\xe0")



# INSERT LINK
# User select a revit file in folders
filepath = select_file('Revit Model (*.rvt)|*.rvt')
#start transaction
with db.Transaction('insert link'):
  for path in filepath:
    linkpath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath)   
  linkoptions = DB.RevitLinkOptions(relative=True)
  linkloadresult = DB.RevitLinkType.Create(revit.doc, linkpath, linkoptions)
  linkinstance = DB.RevitLinkInstance.Create(revit.doc, linkloadresult.ElementId)
  # # Insert from Base Point to Base Point
  # DB.RevitLinkInstance.MoveBasePointToHostBasePoint(linkinstance, resetToOriginalRotation = True)
  # # Insert from Origin To Origin
  # DB.RevitLinkInstance.MoveOriginToHostOrigin(linkinstance, resetToOriginalRotation = True)


# ACQUIRE COORDINATES FROM REVIT LINK
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
collector_link = DB.FilteredElementCollector(linkDoc[0])
linkInstancesId = linkInstances.ToElementIds()
# SET COORDINATE TO MAQUETTE
with db.Transaction('set acquire coordinates'):
  DB.Document.AcquireCoordinates(revit.doc, linkInstancesId[0])
 


# GET SITE LOCATION LINK
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
collector_link = DB.FilteredElementCollector(linkDoc[0])

projectloc_elements_link = collector_link.OfClass(DB.ProjectLocation).ToElements()
for proj in projectloc_elements_link:
  if proj.Name != "Projet" and proj.Name != "Interne":
    projectloc_name_link = proj.Name
  else :
    projectloc_name_link = "Georef"


# RENAME PROJECT LOCATION MAQUETTE
projectloc_category_maq = db.Collector(of_class='ProjectLocation')
projectloc_elements_maq = projectloc_category_maq.get_elements()
for proj in projectloc_elements_maq:
  if proj.Name != "Projet":
    with db.Transaction('rename project location'):
	  proj.Name = projectloc_name_link



# GET BASE POINT LINK
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
collector_link = DB.FilteredElementCollector(linkDoc[0])

basepoint_elements_link = collector_link.OfCategory(DB.BuiltInCategory.OST_ProjectBasePoint).ToElements()
basepoint_nordsud_link = [bp.LookupParameter("N/S").AsDouble() for bp in basepoint_elements_link]
basepoint_estouest_link = [bp.LookupParameter("E/O").AsDouble() for bp in basepoint_elements_link]
basepoint_elevation_link = [bp.LookupParameter("El"+e_a+"v.").AsDouble() for bp in basepoint_elements_link]
basepoint_angle_link = [bp.LookupParameter("Angle par rapport au nord g"+e_a+"ographique").AsDouble() for bp in basepoint_elements_link]


# SET BASE POINT MAQUETTE
basepoint_category = db.Collector(of_category='OST_ProjectBasePoint', is_not_type = True)
basepoint_elements = basepoint_category.get_elements()
with db.Transaction('set base point'):  
  for bp in basepoint_elements:
    bp.Pinned = False
    bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).Set(basepoint_nordsud_link[0])
    bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_EASTWEST_PARAM).Set(basepoint_estouest_link[0])
    bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_ELEVATION_PARAM).Set(basepoint_elevation_link[0])
    bp.get_Parameter(DB.BuiltInParameter.BASEPOINT_ANGLETON_PARAM).Set(basepoint_angle_link[0])



# COLLECT LEVELS MAQUETTE
level_category = db.Collector(of_category='Levels', is_not_type=True)
level_elements = level_category.get_elements()
level_name = [levels.LookupParameter("Nom").AsString() for levels in level_elements]
level_elevation = [levels.Elevation for levels in level_elements]
level_id = [levels.Id for levels in level_elements]


# COLLECT GRIDS MAQUETTE
grid_category = db.Collector(of_category='OST_Grids', is_not_type=True)
grid_elements = grid_category.get_elements()


# COLLECT LEVELS LINK
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
collector_link = DB.FilteredElementCollector(linkDoc[0])

level_elements_link = collector_link.OfClass(DB.Level).ToElements()
level_name_link = [levels.LookupParameter("Nom").AsString() for levels in level_elements_link]
level_elevation_link = [levels.Elevation for levels in level_elements_link]
level_id_link = [levels.Id for levels in level_elements_link]


# COLLECT GRIDS LINK
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]   
collector_link = DB.FilteredElementCollector(linkDoc[0])

grid_elements_link = collector_link.OfClass(DB.Grid).ToElements()
grid_name_link = [grids.Name for grids in grid_elements_link]
grid_length_link = [grids.Curve.Length for grids in grid_elements_link]
grid_origin_link = [grids.Curve.Origin for grids in grid_elements_link]
grid_direction_link = [grids.Curve.Direction for grids in grid_elements_link]


# COLLECT VIEW FAMILY TYPE
family_type_category = db.Collector(of_class= 'ViewFamilyType', is_type = True)
family_type_elements = family_type_category.get_elements()
family_type_id = [views.Id for views in family_type_elements]



# CREATE LEVELS FROM THE LINK
new_level_id = []
with db.Transaction('create levels'):
  for k in range(len(level_elements_link)):
    NewLevel = DB.Level.Create(document = revit.doc, elevation = 0)
    NewLevel.Name = level_name_link[k]
    NewLevel.Elevation = level_elevation_link[k]-basepoint_elevation_link[0]
    new_level_id.append(NewLevel.Id)


# CREATE VIEW
with db.Transaction('create plan views'):
  newplanview = DB.ViewPlan.Create(revit.doc, family_type_id[0], new_level_id[0])
# GO TO THE VIEW
revit.uidoc.ActiveView = newplanview
UI.UIDocument.RefreshActiveView(revit.uidoc)


# CREATE GRIDS FROM THE LINK
with db.Transaction('create grids'):
  for k in range(len(grid_elements_link)):
    start = DB.XYZ(grid_origin_link[k][0],grid_origin_link[k][1],grid_origin_link[k][2])
    end = DB.XYZ(grid_direction_link[k][0]*grid_length_link[k]+grid_origin_link[k][0], 
    	         grid_direction_link[k][1]*grid_length_link[k]+grid_origin_link[k][1],
    	         grid_direction_link[k][2]*grid_length_link[k]+grid_origin_link[k][2])
    BaseLine = DB.Line.CreateBound(start,end)
    NewGrid = DB.Grid.Create(document = revit.doc, line = BaseLine)
    NewGrid.Name = grid_name_link[k]



# DELETE FIRST LEVELS FROM MAQUETTE
with db.Transaction('delete levels'):
  for k in range(len(level_id)):
    DB.Document.Delete(revit.doc, level_id[k])



# # OPTION
# # DELETE LINK
# with db.Transaction('delete link'):
#   collector_maq = DB.FilteredElementCollector(revit.doc)
#   linkedFile = collector_maq.OfCategory(DB.BuiltInCategory.OST_RvtLinks)
#   linkedFileId = linkedFile.ToElementIds()
#   DB.Document.Delete(revit.doc, linkedFileId)