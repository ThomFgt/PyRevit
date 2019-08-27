# -*- coding: utf-8 -*-

__title__ = 'Link Workset Filter'

__doc__ = "Ce programme créé et ajoute des filtres de couleur sur les sous-projets des maquettes en lien. Filtres appliqués sur la vue active. Fonctionne avec Revit 2018"


e_a = str("\xe9")
a_a = str("\xe0")

# IMPORT
import math
import System
import rpw
import clr
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference('System')
from System.Collections.Generic import List
from rpw import revit, db, ui, DB, UI
from rpw.ui.forms import *
from System import Array
from System.Runtime.InteropServices import Marshal
import sys
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


	
# List of all category available for filters
Cat = [
			DB.BuiltInCategory.OST_DuctTerminal ,
			DB.BuiltInCategory.OST_BeamAnalytical ,
			DB.BuiltInCategory.OST_BraceAnalytical ,
			DB.BuiltInCategory.OST_ColumnAnalytical ,
			DB.BuiltInCategory.OST_FloorAnalytical ,
			DB.BuiltInCategory.OST_FoundationSlabAnalytical ,
			DB.BuiltInCategory.OST_IsolatedFoundationAnalytical,
			DB.BuiltInCategory.OST_LinksAnalytical ,
			DB.BuiltInCategory.OST_AnalyticalNodes ,
			# DB.BuiltInCategory.OST_AnalyticalPipeConnections ,
			DB.BuiltInCategory.OST_AnalyticSpaces ,
			DB.BuiltInCategory.OST_AnalyticSurfaces ,
			DB.BuiltInCategory.OST_WallFoundationAnalytical ,
			DB.BuiltInCategory.OST_WallAnalytical ,
			DB.BuiltInCategory.OST_Areas ,
			DB.BuiltInCategory.OST_Assemblies ,
			DB.BuiltInCategory.OST_CableTrayFitting ,
			DB.BuiltInCategory.OST_CableTray ,
			DB.BuiltInCategory.OST_Callouts ,
			DB.BuiltInCategory.OST_Casework ,
			DB.BuiltInCategory.OST_Ceilings ,
			DB.BuiltInCategory.OST_Columns ,
			DB.BuiltInCategory.OST_CommunicationDevices ,
			DB.BuiltInCategory.OST_ConduitFitting ,
			DB.BuiltInCategory.OST_Conduit ,
			DB.BuiltInCategory.OST_CurtainWallPanels ,
			# DB.BuiltInCategory.OST_CurtainGridsCurtaSystem , 
			DB.BuiltInCategory.OST_CurtainWallMullions ,
			DB.BuiltInCategory.OST_DataDevices ,
			DB.BuiltInCategory.OST_DetailComponents ,
			DB.BuiltInCategory.OST_Doors ,
			DB.BuiltInCategory.OST_DuctAccessory ,
			DB.BuiltInCategory.OST_DuctFitting ,
			DB.BuiltInCategory.OST_DuctInsulations ,
			DB.BuiltInCategory.OST_DuctLinings ,
			DB.BuiltInCategory.OST_PlaceHolderDucts ,
			DB.BuiltInCategory.OST_DuctSystem ,
			DB.BuiltInCategory.OST_DuctCurves ,
			DB.BuiltInCategory.OST_ElectricalEquipment ,
			DB.BuiltInCategory.OST_ElectricalFixtures ,
			# DB.BuiltInCategory.OST_ElevationMarks ,
			DB.BuiltInCategory.OST_Entourage ,
			DB.BuiltInCategory.OST_FireAlarmDevices ,
			DB.BuiltInCategory.OST_FlexDuctCurves ,
			DB.BuiltInCategory.OST_FlexPipeCurves ,
			DB.BuiltInCategory.OST_Floors ,
			DB.BuiltInCategory.OST_EdgeSlab ,
			DB.BuiltInCategory.OST_Furniture ,
			DB.BuiltInCategory.OST_FurnitureSystems ,
			DB.BuiltInCategory.OST_GenericModel ,
			DB.BuiltInCategory.OST_Grids ,
			DB.BuiltInCategory.OST_HVAC_Zones ,
			DB.BuiltInCategory.OST_Levels ,
			DB.BuiltInCategory.OST_LightingDevices ,
			DB.BuiltInCategory.OST_LightingFixtures ,
			DB.BuiltInCategory.OST_FabricationContainment ,
			DB.BuiltInCategory.OST_FabricationDuctwork ,
			DB.BuiltInCategory.OST_FabricationDuctworkInsulation ,
			DB.BuiltInCategory.OST_FabricationDuctworkLining , 
			DB.BuiltInCategory.OST_FabricationHangers ,
			DB.BuiltInCategory.OST_FabricationPipework ,
			DB.BuiltInCategory.OST_FabricationPipeworkInsulation ,
			DB.BuiltInCategory.OST_Mass ,
			DB.BuiltInCategory.OST_MassFloor ,
			DB.BuiltInCategory.OST_MassOpening ,
			DB.BuiltInCategory.OST_MassRoof ,
			DB.BuiltInCategory.OST_MassSkylights ,
			DB.BuiltInCategory.OST_MassZone ,
			DB.BuiltInCategory.OST_MassExteriorWall ,
			DB.BuiltInCategory.OST_MassGlazing ,
			DB.BuiltInCategory.OST_MassInteriorWall ,
			DB.BuiltInCategory.OST_MechanicalEquipment ,
			DB.BuiltInCategory.OST_NurseCallDevices ,
			DB.BuiltInCategory.OST_Parking ,
			DB.BuiltInCategory.OST_Parts ,
			DB.BuiltInCategory.OST_PipeAccessory ,
			DB.BuiltInCategory.OST_PipeFitting ,
			DB.BuiltInCategory.OST_PipeInsulations ,
			DB.BuiltInCategory.OST_PlaceHolderPipes ,
			DB.BuiltInCategory.OST_PipeCurves ,
			DB.BuiltInCategory.OST_PipingSystem ,
			DB.BuiltInCategory.OST_Planting ,
			DB.BuiltInCategory.OST_PlumbingFixtures ,
			# DB.BuiltInCategory.OST_Railings ,
			# DB.BuiltInCategory.OST_RailingBalusterRailCut ,
			DB.BuiltInCategory.OST_RailingHandRail  ,
			DB.BuiltInCategory.OST_RailingSupport ,
			DB.BuiltInCategory.OST_RailingTermination ,
			DB.BuiltInCategory.OST_RailingTopRail ,
			DB.BuiltInCategory.OST_Ramps ,
			DB.BuiltInCategory.OST_ReferenceLines ,
			# DB.BuiltInCategory.OST_ReferencePoints_Planes ,
			DB.BuiltInCategory.OST_Roads ,
			DB.BuiltInCategory.OST_Roofs ,
			DB.BuiltInCategory.OST_Fascia ,
			DB.BuiltInCategory.OST_Gutter , 
			DB.BuiltInCategory.OST_Rooms ,
			DB.BuiltInCategory.OST_Sections ,
			DB.BuiltInCategory.OST_SecurityDevices ,
			DB.BuiltInCategory.OST_ShaftOpening ,
			DB.BuiltInCategory.OST_Site ,
			DB.BuiltInCategory.OST_BuildingPad ,
			DB.BuiltInCategory.OST_SiteProperty ,
			# DB.BuiltInCategory.OST_AnalyticSpaces ,
			DB.BuiltInCategory.OST_SpecialityEquipment ,
			DB.BuiltInCategory.OST_Sprinklers ,
			DB.BuiltInCategory.OST_Stairs ,
			DB.BuiltInCategory.OST_AreaRein ,
			# DB.BuiltInCategory.OST_BeamAnalytical ,
			DB.BuiltInCategory.OST_StructuralColumns ,
			DB.BuiltInCategory.OST_StructConnections ,
			DB.BuiltInCategory.OST_FabricAreas ,
			DB.BuiltInCategory.OST_FabricReinforcement ,
			DB.BuiltInCategory.OST_StructuralFoundation ,
			DB.BuiltInCategory.OST_StructuralFraming ,
			DB.BuiltInCategory.OST_InternalAreaLoads  ,
			DB.BuiltInCategory.OST_InternalLineLoads ,
			DB.BuiltInCategory.OST_InternalPointLoads ,
			DB.BuiltInCategory.OST_AreaLoads ,
			DB.BuiltInCategory.OST_LineLoads ,
			DB.BuiltInCategory.OST_PointLoads ,
			DB.BuiltInCategory.OST_PathRein ,
			DB.BuiltInCategory.OST_Rebar ,
			DB.BuiltInCategory.OST_Coupler ,
			DB.BuiltInCategory.OST_StructuralStiffener ,
			DB.BuiltInCategory.OST_StructuralTruss ,
			DB.BuiltInCategory.OST_SwitchSystem ,
			DB.BuiltInCategory.OST_TelephoneDevices ,
			DB.BuiltInCategory.OST_Topography ,
			DB.BuiltInCategory.OST_Walls ,
			DB.BuiltInCategory.OST_Windows ,
			DB.BuiltInCategory.OST_Wire ,
			
			]
catlist = [ElementId(i) for i in Cat]
typeCatList = List[DB.ElementId](catlist)
# /////////////////////////////////////////////////////////////////////////

# /////////////////////////////////////////////////////////////////////////
# COLLECT LINK INSTANCE
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
none_index = [i for i, link in enumerate(linkDoc) if link==None]
linkDoc_name = [link.Name for link in linkInstances]
for index in none_index:
	linkDoc_name.pop(index) #remove unload link by finding None in linkDoc
linkDoc_name.insert(0, "Mon document")

#COLLECT WORKSETS IN DOC
collect = DB.FilteredWorksetCollector(revit.doc)
collect_worksets = collect.OfKind(DB.WorksetKind.UserWorkset)
worksets_name_doc = [workset.Name for workset in collect_worksets]

# COLLECT WORKSETS LINKS
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
worksets_name_link = []
for j in range(len(linkDoc)):
	try:
		collector_link = DB.FilteredWorksetCollector(linkDoc[j])
		collect_worksets_link = collector_link.OfKind(DB.WorksetKind.UserWorkset)
		worksets_name_link.append([workset.Name for workset in collect_worksets_link])
	except : pass 
worksets_name_link.insert(0, worksets_name_doc)	

# COLOR LIST
color_list = ["bleu","rouge","vert","jaune","orange","marron","rose","violet","bleu clair"]

# LIST OUT
data_out = [[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0]]
# /////////////////////////////////////////////////////////////////////////

# /////////////////////////////////////////////////////////////////////////
# START WINDOWS FORM
class MyListBox(Form):
	# docstring for ClassName
	def __init__(self):
		self.Text = "Plugin link workset filter"
		screenSize = Screen.GetWorkingArea(self)
		self.Height = screenSize.Height / 2
		self.Width = screenSize.Width / 3
		self.dico = {}

	# label for listbox 1 on the left
		label1 = Label()
		label1.Parent = self
		label1.Text = "Liens chargés"
		label1.Location = Point(5,5)
	# listbox 1 on the left
		self.listbox1 = ListBox()
		self.listbox1.Parent = self
		self.listbox1.Location = Point(5,30)
		self.listbox1.Size = Size(180,100)
		for k in linkDoc_name:
			self.listbox1.Items.Add(k)
		self.listbox1.SelectedIndexChanged += self.OnChanged

	# label for listbox 2 on the middle
		label2 = Label()
		label2.Parent = self
		label2.Text = "Sous-projets liens"
		label2.Location = Point(200,5)
	# listbox 2 on the middle
		self.listbox2 = ListBox() 
		self.listbox2.Parent = self
		self.listbox2.Location = Point(200,30)
		self.listbox2.Size = Size(180,100)
		self.listbox2.Items.Add(" ")

	# label for listbox 3 on the right
		label3 = Label()
		label3.Parent = self
		label3.Text = "Choix couleurs"
		label3.Location = Point(400,5)
	# listbox 3 on the right
		self.listbox3 = ListBox()
		self.listbox3.Parent = self
		self.listbox3.Location = Point(400,30)
		self.listbox3.Size = Size(180,100)
		for k in color_list:
			self.listbox3.Items.Add(k)

	# button add filter
		self.buttonAdd = Button()
		self.buttonAdd.Parent = self
		self.buttonAdd.Text = "Ajouter filtre"
		self.buttonAdd.Location = Point(450,150)
		self.buttonAdd.Size = Size(100,25)
		self.buttonAdd.Click += self.UpDate
	# button OK
		self.buttonOk = Button()
		self.buttonOk.Parent = self
		self.buttonOk.Text = "Ok Chef"
		self.buttonOk.Location = Point(500,420)
		self.buttonOk.Click += self.Ok
	# button Cancel
		self.buttonCancel = Button()
		self.buttonCancel.Parent = self
		self.buttonCancel.Text = "ANNULE"
		self.buttonCancel.Location = Point(400,420)
		self.buttonCancel.Click += self.Cancel

	# groupbox for visualize filter added
		self.groupBox1 = GroupBox()
		self.groupBox1.Parent = self
		self.groupBox1.Text = "Détail des filtres"
		self.groupBox1.Location = Point(50,200)
		self.groupBox1.Size = Size(500,200)

	# all empty label in the groupbox. 8 lines and 3 columns.
		for k in range(1,9):
			self.label1 = Label()
			self.label1.Name = "label1" + str(k)
			self.label1.Parent = self.groupBox1
			self.label1.Text = " "
			self.label1.Location = Point(40,20*k)
			self.label1.Size = Size(120,20)
			self.label1.BorderStyle = BorderStyle.Fixed3D
			self.dico[self.label1.Name] = self.label1
			
			self.label2 = Label()
			self.label2.Name = "label2" + str(k)
			self.label2.Parent = self.groupBox1
			self.label2.Text = " "
			self.label2.Location = Point(200,20*k)
			self.label2.Size = Size(120,20)
			self.label2.BorderStyle = BorderStyle.Fixed3D
			self.dico[self.label2.Name] = self.label2
			
			self.label3 = Label()
			self.label3.Name = "label3" + str(k)
			self.label3.Parent = self.groupBox1
			self.label3.Text = " "
			self.label3.Location = Point(350,20*k)
			self.label3.Size = Size(120,20)
			self.label3.BorderStyle = BorderStyle.Fixed3D
			self.dico[self.label3.Name] = self.label3

	# counter for changing line in groupbox
		self.counterDict={'1':1,'2':1,'3':1}
	# center the form on screen
		self.CenterToScreen()

	# EVENT ON CLICK //////
	# Event on changing selected link in first listbox
	def OnChanged(self,sender,event):
		self.listbox2.Items.Clear()
		index = sender.SelectedIndex
		for k in worksets_name_link[index]:
			self.listbox2.Items.Add(k)
		
	# Event after adding a filter in groupbox. Right on the next line in groupbox. Stop at 8 filters max.
	def UpDate(self,sender,event):
		label = self.dico["label1"+str(self.counterDict["1"])]
		label.Text = self.listbox1.SelectedItem
		label = self.dico["label2"+str(self.counterDict["2"])]
		label.Text = self.listbox2.SelectedItem
		label = self.dico["label3"+str(self.counterDict["3"])]
		label.Text = self.listbox3.SelectedItem
		for k in range(1,4):
			if self.counterDict[str(k)] < 8:
				self.counterDict[str(k)] += 1

	# Event on Cancel Button Click
	def Cancel(self,sender,event):
		self.Close()

	# Event on OK Button Click
	def Ok(self,sender,event):
		for i in range(0,8):
			data_out[i][0] = self.dico["label1"+str(i+1)].Text
			data_out[i][1] = self.dico["label2"+str(i+1)].Text
			data_out[i][2] = self.dico["label3"+str(i+1)].Text
		self.Close()

# run form
Application.Run(MyListBox())
# END OF WINDOWS FORM ///////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////////

# function for find in list of list
inL = lambda e, L: reduce(lambda b, l: b or e in l, L, False)
# dictionnary of color 
dico_couleur = {"bleu":DB.Color(0,0,255),"rouge":DB.Color(255,0,0),"vert":DB.Color(0,255,0),"jaune":DB.Color(255,255,0),"orange":DB.Color(255,127,0),"marron":DB.Color(118,71,23),"rose":DB.Color(255,0,255),"violet":DB.Color(155,0,255),"bleu clair":DB.Color(0,255,255)}

# CREATE WORKSET FILTER /////////////////////////////////////////////
okgo = True
if inL(0, data_out):
	okgo = False
	print "A+" # if these is a 0 in listout, it means user clicked on Cancel button
else :
	for i in range(len(data_out)):
		if data_out[i][0] == "" and data_out[i][1] != "" and data_out[i][2] != "":
			okgo = False
			print "Try again !\nMerci de séléctionner un élément de chaque liste"
			break
		elif data_out[i][0] != "" and data_out[i][1] == "" and data_out[i][2] != "":
			okgo = False
			print "Try again !\nMerci de séléctionner un élément de chaque liste"
			break
		elif data_out[i][0] != "" and data_out[i][1] != "" and data_out[i][2] == "":
			okgo = False
			print "Try again !\nMerci de séléctionner un élément de chaque liste"
			break
		elif data_out[i][0] == "" and data_out[i][1] == "" and data_out[i][2] != "":
			okgo = False
			print "Try again !\nMerci de séléctionner un élément de chaque liste"
			break
		elif data_out[i][0] != "" and data_out[i][1] == "" and data_out[i][2] == "":
			okgo = False
			print "Try again !\nMerci de séléctionner un élément de chaque liste"
			break
		elif data_out[i][0] == "" and data_out[i][1] != "" and data_out[i][2] == "":
			okgo = False
			print "Try again !\nMerci de séléctionner un élément de chaque liste"
			break
		elif data_out[0][0] == " ":
			okgo = False 
			print "Pas de Filtre ? \nA+"
			break

if okgo :
	for j in range(len(data_out)):
		if data_out[j][0] == " ":
			pass
		else : 
			EndWell = True
			with db.Transaction('create workset'):
				if not revit.doc.IsWorkshared :
					EndWell = False
					print "Erreur\n\nLe projet n'est pas en collaboration, impossible de créer un sous-projet."
					break 
				try:
					new_workset = DB.Workset.Create(revit.doc, str(data_out[j][1]))
					worksetId = str(new_workset.Id)
					elemId = DB.ElementId(int(worksetId))
					workset_parameter = DB.ElementId(DB.BuiltInParameter.ELEM_PARTITION_PARAM) #BuiltinParam for workset
				except :
					workset_parameter = DB.ElementId(DB.BuiltInParameter.ELEM_PARTITION_PARAM) #if workset name already exist
					collect = DB.FilteredWorksetCollector(revit.doc)
					collect_worksets = collect.OfKind(DB.WorksetKind.UserWorkset)
					worksets_name_doc = [workset.Name for workset in collect_worksets]
					worksets_id_doc = [workset.Id for workset in collect_worksets]
					elemId = DB.ElementId(int(str(worksets_id_doc[worksets_name_doc.index(str(data_out[j][1]))])))
			with db.Transaction('create rule'):
				rule = [DB.ParameterFilterRuleFactory.CreateEqualsRule(workset_parameter,elemId)]

			with db.Transaction('create filter'):
				try :
					new_filter = DB.ParameterFilterElement.Create(revit.doc,"Filtre " + str(data_out[j][1]),typeCatList,rule)
				except:
					EndWell = False 
					print "Un nom des nouveaux filtres existe déjà. le filtre ---%s--- n'a pas été créé ni ajouté à la vue mais le sous projet du même nom a été créé." %data_out[j][1]
					break

			with db.Transaction('add filter'):
				DB.View.AddFilter(revit.doc.ActiveView , new_filter.Id)

			with db.Transaction('set filter overrides'):
				ogs = DB.OverrideGraphicSettings()
				ogs.SetProjectionFillColor(dico_couleur[str(data_out[j][2])])
				try :
					ogs.SetProjectionFillPatternId((DB.FillPatternElement.GetFillPatternElementByName(revit.doc, DB.FillPatternTarget.Drafting, "Uni").Id))
				except :
					ogs.SetProjectionFillPatternId((DB.FillPatternElement.GetFillPatternElementByName(revit.doc, DB.FillPatternTarget.Drafting, "<Remplissage de solide>").Id))
				ogs.SetProjectionFillPatternVisible(True)
				ogs.SetSurfaceTransparency(80)
				ogs.SetCutFillColor(dico_couleur[str(data_out[j][2])])
				try :
					ogs.SetCutFillPatternId((DB.FillPatternElement.GetFillPatternElementByName(revit.doc, DB.FillPatternTarget.Drafting, "Uni").Id))
				except :
					ogs.SetCutFillPatternId((DB.FillPatternElement.GetFillPatternElementByName(revit.doc, DB.FillPatternTarget.Drafting, "<Remplissage de solide>").Id))
				ogs.SetCutFillPatternVisible(True)
				DB.View.SetFilterOverrides(revit.doc.ActiveView , new_filter.Id , ogs)
	if EndWell:
		print "Finito !\nSi les filtres ont été créés mais n'ont pas eu d'effet sur la vue active alors il faut parcourir la liste des filtres créés dans la fenêtre -ajouter/Modifier- les filtres puis cliquer sur OK."

# ////////////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////////