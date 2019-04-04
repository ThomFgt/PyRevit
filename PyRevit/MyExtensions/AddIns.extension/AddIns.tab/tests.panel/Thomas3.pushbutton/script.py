# -*- coding: utf-8 -*-

__title__ = 'Arborescence\nSynthese'

__doc__ = "Ce programme créé des vues à partir des niveaux existant et les range dans une arborescence. Les paramètres BPS_Arborescence_Lot, BPS_Arborescence_Niveau, BPS_Arborescence_Zone doivent être créés."


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



# COLLECT VIEW FAMILY TYPE ////////////////////////////////////////////////////////
family_type_category = db.Collector(of_class= 'ViewFamilyType', is_type = True)
family_type_elements = family_type_category.get_elements()
family_type_id = [views.Id for views in family_type_elements]
family_type_name = [views.LookupParameter('Nom du type').AsString() for views in family_type_elements]

plan_etage_id = []
plan_structure_id = []
plan_plafond_id = [] 

for i in range(len(family_type_name)):
	if "étage" in family_type_name[i]:
		plan_etage_id.append(family_type_id[i])
	if "plafond" in family_type_name[i]:
		plan_plafond_id.append(family_type_id[i])
	if "structure" in family_type_name[i]:
		plan_structure_id.append(family_type_id[i])
# /////////////////////////////////////////////////////////////////////////////////


# COLLECT LEVELS //////////////////////////////////////////////////////////////////
level_category = db.Collector(of_category='Levels', is_not_type=True)
level_elements = level_category.get_elements()
level_name = [levels.LookupParameter("Nom").AsString() for levels in level_elements]
level_id = [levels.Id for levels in level_elements]
# /////////////////////////////////////////////////////////////////////////////////



# CREATE VIEWS ////////////////////////////////////////////////////////////////////
with db.Transaction('create arborescence'):
	for j in range(len(level_elements)):

		newplanview1 = DB.ViewPlan.Create(revit.doc, plan_etage_id[0], level_id[j])
		newplanview1.Name = "Réseaux - "+level_name[j]
		newplanview1.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Technique")
		newplanview1.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newplanview1.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newplanview2 = DB.ViewPlan.Create(revit.doc, plan_etage_id[0], level_id[j])
		newplanview2.Name = "Réservation - "+level_name[j]
		newplanview2.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Technique")
		newplanview2.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newplanview2.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newplanview3 = DB.ViewPlan.Create(revit.doc, plan_etage_id[0], level_id[j])
		newplanview3.Name = "Terminaux - "+level_name[j]
		newplanview3.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Technique")
		newplanview3.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newplanview3.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newplanview4 = DB.ViewPlan.Create(revit.doc, plan_etage_id[0], level_id[j])
		newplanview4.Name = "ARC étage - "+level_name[j]
		newplanview4.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Architecturale")
		newplanview4.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newplanview4.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newfloorview1 = DB.ViewPlan.Create(revit.doc, plan_structure_id[0], level_id[j])
		newfloorview1.Name = "STR poutre - "+level_name[j]
		newfloorview1.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Technique")
		newfloorview1.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newfloorview1.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newfloorview2 = DB.ViewPlan.Create(revit.doc, plan_structure_id[0], level_id[j])
		newfloorview2.Name = "STR Poutre - "+level_name[j]
		newfloorview2.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Architecturale")
		newfloorview2.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newfloorview2.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newceilingview1 = DB.ViewPlan.Create(revit.doc, plan_plafond_id[0], level_id[j])
		newceilingview1.Name = "ARC Plafond - "+level_name[j]
		newceilingview1.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Technique")
		newceilingview1.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newceilingview1.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")

		newceilingview2 = DB.ViewPlan.Create(revit.doc, plan_plafond_id[0], level_id[j])
		newceilingview2.Name = "ARC plafond - "+level_name[j]
		newceilingview2.LookupParameter("BPS_Arborescence_Lot").Set("Synthèse Architecturale")
		newceilingview2.LookupParameter("BPS_Arborescence_Niveau").Set(level_name[j])
		newceilingview2.LookupParameter("BPS_Arborescence_Zone").Set("Toute Zone")
		 
# /////////////////////////////////////////////////////////////////////////////////


#COLLECT TEMPLATE //////////////////////////////////////////////////////////////
view_category = db.Collector(of_category = 'OST_Views').ToElements()
view_templates = [view for view in view_category if view.IsTemplate]
view_templates_id = [view.Id for view in view_templates]
view_templates_name = [view.Name for view in view_templates]

# SET TEMPLATE
view_category = db.Collector(of_category = 'OST_Views', is_not_type = True)
view_element = view_category.get_elements()
view_id = [views.Id for views in view_element]
# for view in view_element:
# 	view.ViewTemplateId = view_templates_id[0] #works only with not active view
#//////////////////////////////////////////////////////////////////////////////////






















