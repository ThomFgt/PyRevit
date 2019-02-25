# -*- coding: utf-8 -*-

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



# COLLECT VIEWS
view_category = db.Collector(of_category = 'OST_Views', is_not_type = True)
view_element = view_category.get_elements()
view_id = [views.Id for views in view_element]
# COLLECT SHEETS
sheet_category = db.Collector(of_category = 'OST_Sheets', is_not_type = True)
sheet_element = sheet_category.get_elements()
sheet_id = [sheets.Id for sheets in sheet_element]
# SET POINT 
point = DB.XYZ(1,1,0)
# PLACE VIEW IN SHEET 
with db.Transaction('place view in sheet'):
	DB.Viewport.Create(revit.doc, sheet_id[0], view_id[6], point)









# # ////////////////////////////////////////////////////////////////////////
# # COLLECT WORKSETS NAME LINK
# collector_maq = DB.FilteredElementCollector(revit.doc)
# linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
# linkDoc = [links.GetLinkDocument() for links in linkInstances]
# collector_link = DB.FilteredWorksetCollector(linkDoc[0])

# collect_worksets_link = collector_link.OfKind(DB.WorksetKind.UserWorkset)
# worksets_name_link = [workset.Name for workset in collect_worksets_link]
# for workset in worksets_name_link:
#   print workset


# # CREATE WORKSETS IN MAQUETTE
# with db.Transaction('create worksets'):
#   for workset in worksets_name_link:
#     DB.Workset.Create(revit.doc, workset)
# # ///////////////////////////////////////////////////////////////////////






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
















