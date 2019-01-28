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




# # USERFORM BIM CHECKER

# components = [Label('Check a choisir'),
#               CheckBox('checkbox1', 'Verifier le nom de la maquette'),
#               CheckBox('checkbox2', 'Verifier si la maquette est centrale ou local'),
#               CheckBox('checkbox3', 'Verifier le nombre de groupe dans la maquette' ),
#               CheckBox('checkbox4', 'Verifier emplacement partage'),
#               CheckBox('checkbox5', 'Verifier le point de base'),
#               CheckBox('checkbox6', 'Verifier les niveaux'),
#               CheckBox('checkbox7', 'Verifier les quadrillages'),
#               CheckBox('checkbox8', 'Verifier les avertissements'),
#               CheckBox('checkbox9', 'Verifier si chaque vue est utilisee pour une vue en plan'),
#               CheckBox('checkbox10', 'Verifier le poids de la maquette'),
#               CheckBox('checkbox11', 'Verifier la version Revit'),
#               Label('Version Revit'),
#               ComboBox('combobox1', {'2015':2015, '2016':2016, '2017':2017, '2018':2018, '2019':2019}),
#               Label('Nom maquette exacte'),
#               TextBox('textbox1', Text="EXEMPLE_MAQUETTE.rvt"),
#               Label('Poids maximum maquette'),
#               ComboBox('combobox2', {'150':150, '200':200, '250':250, '300':300, '350':350, '400':400, '450':450}),
#               Label('choisir maquette QNP'),
#               Button(button_text='QNP'),             
#               Label('choisir visa Excel'),
#               Button(button_text='VISA BIM'),
#               Separator(), 
#               Button('OK GO'),]
# form = FlexForm('BIM CHECKER DU FUTUR', components)
# form.show()


# # /////////////////////////////////////////////////////////////////////////////////////////////////////

# # INSERT LINK
# # User select a revit file in folders
# filepath = select_file('Revit Model (*.rvt)|*.rvt')
# #start transaction
# with db.Transaction('insert link'):
#   for path in filepath:
#     linkpath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath)   
#   linkoptions = DB.RevitLinkOptions(relative=True)
#   linkloadresult = DB.RevitLinkType.Create(revit.doc, linkpath, linkoptions)
#   linkinstance = DB.RevitLinkInstance.Create(revit.doc, linkloadresult.ElementId)
#   # # Insert from Base Point to Base Point
#   # DB.RevitLinkInstance.MoveBasePointToHostBasePoint(linkinstance, resetToOriginalRotation = True)
#   # # Insert from Origin To Origin
#   # DB.RevitLinkInstance.MoveOriginToHostOrigin(linkinstance, resetToOriginalRotation = True)

# # ACQUIRE COORDINATES FROM REVIT LINK
# collector_maq = DB.FilteredElementCollector(revit.doc)
# linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
# linkDoc = [links.GetLinkDocument() for links in linkInstances]
# collector_link = DB.FilteredElementCollector(linkDoc[0])
# linkInstancesId = linkInstances.ToElementIds()
# with db.Transaction('acquire coordinates'):
#   DB.Document.AcquireCoordinates(revit.doc, linkInstancesId[0])
 

# # DELETE LINK
# with db.Transaction('delete link'):
#   collector_maq = DB.FilteredElementCollector(revit.doc)
#   linkedFile = collector_maq.OfCategory(DB.BuiltInCategory.OST_RvtLinks)
#   linkedFileId = linkedFile.ToElementIds()
#   DB.Document.Delete(revit.doc, linkedFileId)


# # INSERT LINK
# with db.Transaction('insert link'):
#   for path in filepath:
#     linkpath = DB.ModelPathUtils.ConvertUserVisiblePathToModelPath(filepath)   
#   linkoptions = DB.RevitLinkOptions(relative=True)
#   linkloadresult = DB.RevitLinkType.Create(revit.doc, linkpath, linkoptions)
#   linkinstance = DB.RevitLinkInstance.Create(revit.doc, linkloadresult.ElementId)
#   # # Insert from Base Point to Base Point
#   # DB.RevitLinkInstance.MoveBasePointToHostBasePoint(linkinstance, resetToOriginalRotation = True)
#   # # Insert from Origin To Origin
#   # DB.RevitLinkInstance.MoveOriginToHostOrigin(linkinstance, resetToOriginalRotation = True)
# # /////////////////////////////////////////////////////////////////////////////////////////////////////




# # GET SITE LOCATION LINK
# collector_maq = DB.FilteredElementCollector(revit.doc)
# linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
# linkDoc = [links.GetLinkDocument() for links in linkInstances]
# collector_link = DB.FilteredElementCollector(linkDoc[0])

# siteloc_elements_link = collector_link.OfClass(DB.ProjectLocation).ToElements()
# projectloc_name_link = [proj.Name for proj in siteloc_elements_link]
# for proj in siteloc_elements_link:
#   if proj.Name != "Projet" and proj.Name != "Interne":
#     siteloc_longitude_link = proj.SiteLocation.Longitude
#     siteloc_latitude_link = proj.SiteLocation.Latitude
#     siteloc_elevation_link = proj.SiteLocation.Elevation
#     siteloc_placename_link = proj.SiteLocation.PlaceName

# print siteloc_elements_link, projectloc_name_link, siteloc_longitude_link, siteloc_latitude_link, siteloc_elevation_link, siteloc_placename_link

projectloc_category_maq = db.Collector(of_class='ProjectLocation')
projectloc_elements_maq = projectloc_category_maq.get_elements()
for proj in projectloc_elements_maq:
  if proj.Name != "Projet" : 
    siteloc_longitude_maq = [proj.SiteLocation.Longitude for proj in projectloc_elements_maq]
    siteloc_latitude_maq = [proj.SiteLocation.Latitude for proj in projectloc_elements_maq]
    siteloc_elevation_maq = [proj.SiteLocation.Elevation for proj in projectloc_elements_maq]
    siteloc_placename_maq = [proj.SiteLocation.PlaceName for proj in projectloc_elements_maq] 

print(projectloc_category_maq, projectloc_elements_maq, siteloc_longitude_maq, siteloc_latitude_maq, siteloc_elevation_maq, siteloc_placename_maq)

# # COLLECT LEVELS MAQUETTE
# level_category = db.Collector(of_category='Levels', is_not_type=True)
# level_elements = level_category.get_elements()
# level_name = [levels.LookupParameter("Nom").AsString() for levels in level_elements]
# level_elevation = [levels.Elevation for levels in level_elements]
# level_id = [levels.Id for levels in level_elements]

# # COLLECT GRIDS MAQUETTE
# grid_category = db.Collector(of_category='OST_Grids', is_not_type=True)
# grid_elements = grid_category.get_elements()

# # COLLECT LEVELS LINK
# collector_maq = DB.FilteredElementCollector(revit.doc)
# linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
# linkDoc = [links.GetLinkDocument() for links in linkInstances]
# collector_link = DB.FilteredElementCollector(linkDoc[0])

# level_elements_link = collector_link.OfClass(DB.Level).ToElements()
# level_name_link = [levels.LookupParameter("Nom").AsString() for levels in level_elements_link]
# level_elevation_link = [levels.Elevation for levels in level_elements_link]

# # COLLECT GRIDS LINK
# collector_maq = DB.FilteredElementCollector(revit.doc)
# linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
# linkDoc = [links.GetLinkDocument() for links in linkInstances]   
# collector_link = DB.FilteredElementCollector(linkDoc[0])

# grid_elements_link = collector_link.OfClass(DB.Grid).ToElements()
# grid_name_link = [grids.Name for grids in grid_elements_link]
# grid_length_link = [grids.Curve.Length for grids in grid_elements_link]
# grid_origin_link = [grids.Curve.Origin for grids in grid_elements_link]
# grid_direction_link = [grids.Curve.Direction for grids in grid_elements_link]


# # CREATE LEVELS
# with db.Transaction('create levels'):
#   for k in range(len(level_elements_link)):
#     NewLevel = DB.Level.Create(document = revit.doc, elevation = 0)
#     NewLevel.Name = level_name_link[k]
#     NewLevel.Elevation = level_elevation_link[k]

# # DELETE LEVELS 
# with db.Transaction('delete levels'):
#   for levels in level_id:
#     DB.Document.Delete(revit.doc, levels)

# # CREATE GRIDS
# with db.Transaction('create grids'):
#   for k in range(len(grid_elements_link)):
#     start = db.XYZ(0,0,0)
#     end = db.XYZ(0,0,0)
#     BaseLine = db.Line.new([0,0],[1,1])
#     NewGrid = DB.Grid.Create(document = revit.doc, line = BaseLine)
#     NewGrid.Name = grid_name_link[k]
#     NewGrid.Curve.Origin = grid_origin_link[k]
#     NewGrid.Curve.Length = grid_length_link[k]
#     NewGrid.Curve.Direction = grid_direction_link[k]




# # OUT

# print siteloc_elements_link, siteloc_elevation_link, siteloc_latitude_link, siteloc_longitude_link, siteloc_placename_link