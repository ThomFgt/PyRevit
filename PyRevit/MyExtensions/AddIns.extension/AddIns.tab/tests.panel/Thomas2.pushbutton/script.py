# -*- coding: utf-8 -*-
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





# /////////////////////////////////////////////////////////////////////////
# COLLECT LINK INSTANCE
collector_maq = DB.FilteredElementCollector(revit.doc)
linkInstances = collector_maq.OfClass(DB.RevitLinkInstance)
linkDoc = [links.GetLinkDocument() for links in linkInstances]
linkDoc_name = [link.Name for link in linkInstances]


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
	
#COLOR
color_list = ["bleu","rouge","vert","jaune","orange","marron","rose","violet","bleu clair"]

data_out = [[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0],[0,0,0]]

# /////////////////////////////////////////////////////////////////////////
class MyListBox(Form):
	"""docstring for ClassName"""
	def __init__(self):
		self.Text = "Ma form cool"
		screenSize = Screen.GetWorkingArea(self)
		self.Height = screenSize.Height / 2
		self.Width = screenSize.Width / 3

		self.dico = {}



		label1 = Label()
		label1.Parent = self
		label1.Text = "la liste 1"
		label1.Location = Point(5,5)

		self.listbox1 = ListBox()
		self.listbox1.Parent = self
		self.listbox1.Location = Point(5,30)
		self.listbox1.Size = Size(180,100)
		for k in linkDoc_name:
			self.listbox1.Items.Add(k)
		self.listbox1.SelectedIndexChanged += self.OnChanged




		label2 = Label()
		label2.Parent = self
		label2.Text = "la liste 2"
		label2.Location = Point(200,5)

		self.listbox2 = ListBox() 
		self.listbox2.Parent = self
		self.listbox2.Location = Point(200,30)
		self.listbox2.Size = Size(180,100)
		self.listbox2.Items.Add(" ")


			


		label3 = Label()
		label3.Parent = self
		label3.Text = "la liste 3"
		label3.Location = Point(400,5)

		self.listbox3 = ListBox()
		self.listbox3.Parent = self
		self.listbox3.Location = Point(400,30)
		self.listbox3.Size = Size(180,100)
		for k in color_list:
			self.listbox3.Items.Add(k)




		self.buttonAdd = Button()
		self.buttonAdd.Parent = self
		self.buttonAdd.Text = "Ajouter filtre"
		self.buttonAdd.Location = Point(450,150)
		self.buttonAdd.Size = Size(100,25)
		self.buttonAdd.Click += self.UpDate


		self.buttonOk = Button()
		self.buttonOk.Parent = self
		self.buttonOk.Text = "Ok Chef"
		self.buttonOk.Location = Point(500,420)
		self.buttonOk.Click += self.Ok


		self.buttonCancel = Button()
		self.buttonCancel.Parent = self
		self.buttonCancel.Text = "ANNULE"
		self.buttonCancel.Location = Point(400,420)
		self.buttonCancel.Click += self.Cancel


		self.groupBox1 = GroupBox()
		self.groupBox1.Parent = self
		self.groupBox1.Text = "Détail des filtres"
		self.groupBox1.Location = Point(50,200)
		self.groupBox1.Size = Size(500,200)

		
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


		self.counterDict={'1':1,'2':1,'3':1}
		
		self.CenterToScreen()



	def OnChanged(self,sender,event):
		self.listbox2.Items.Clear()
		index = sender.SelectedIndex
		for k in worksets_name_link[index]:
			self.listbox2.Items.Add(k)
		


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



	def Cancel(self,sender,event):
		self.Close()


	def Ok(self,sender,event):
		for i in range(0,8):
			data_out[i][0] = self.dico["label1"+str(i+1)].Text
			data_out[i][1] = self.dico["label2"+str(i+1)].Text
			data_out[i][2] = self.dico["label3"+str(i+1)].Text
		self.Close()




Application.Run(MyListBox())

print data_out

# ///////////////////////////////////////////////////////////////////////
inL = lambda e, L: reduce(lambda b, l: b or e in l, L, False)

dico_couleur = {}

if inL(0, data_out):
	print "A+"
else:
# CREATE FILTER FOR WORKSET
	view = revit.doc.ActiveView 
	with db.Transaction('create filter'):
		for j in range(len(data_out)):
			if data_out[j][0] == " " :
				pass
			else:
				new_workset = DB.Workset.Create(revit.doc, str(data_out[j][1]))
				worksetId = str(new_workset.Id)
				elemId = DB.ElementId(int(worksetId))

				all_Categories_Id = System.Enum.GetValues(BuiltInCategory)
				catlist = [ElementId(i) for i in all_Categories_Id]
				typeCatList = List[DB.ElementId](catlist)

				workset_parameter = DB.ElementId(DB.BuiltInParameter.ELEM_PARTITION_PARAM)

				rule = [DB.ParameterFilterRuleFactory.CreateEqualsRule(workset_parameter,elemId)]

				new_filter = DB.ParameterFilterElement.Create(revit.doc,"Filtre " + str(data_out[j][1]),typeCatList,rule)

				DB.View.AddFilter(view , new_filter.Id)

				ogs = DB.OverrideGraphicSettings()
				ogs.SetProjectionFillColor(DB.Color(255,0,0))
				ogs.SetProjectionFillPatternId(ElementId(4))
				ogs.SetProjectionFillPatternVisible(True)
				ogs.SetSurfaceTransparency(80)

				DB.View.SetFilterOverrides(view, new_filter.Id , ogs)


# ////////////////////////////////////////////////////////////////////////

# # //////////////////////////////////////////////////////////////////////////
# # COLLECT VIEWS
# view_category = db.Collector(of_category = 'OST_Views', is_not_type = True)
# view_element = view_category.get_elements()
# view_id = [views.Id for views in view_element]
# # COLLECT SHEETS
# sheet_category = db.Collector(of_category = 'OST_Sheets', is_not_type = True)
# sheet_element = sheet_category.get_elements()
# sheet_id = [sheets.Id for sheets in sheet_element]
# # SET POINT 
# point = DB.XYZ(1,1,0)
# # PLACE VIEW IN SHEET 
# with db.Transaction('place view in sheet'):
# 	DB.Viewport.Create(revit.doc, sheet_id[0], view_id[6], point)
# # //////////////////////////////////////////////////////////////////////////




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



# class FailureHandler(IFailuresPreprocessor):
#                def __init__(self):
#                               self.ErrorMessage = ""
#                               self.ErrorSeverity = ""
#                def PreprocessFailures(self, failuresAccessor):
#                               # failuresAccessor.DeleteAllWarning()
#                               # return FailureProcessingResult.Continue
#                               failures = failuresAccessor.GetFailureMessages()
#                               rslt = ""
#                               for f in failures:
#                                             fseverity = failuresAccessor.GetSeverity()
#                                             if fseverity == FailureSeverity.Warning:
#                                                            failuresAccessor.DeleteWarning(f)
#                                             elif fseverity == FailureSeverity.Error:
#                                                            rslt = "Error"
#                                                            failuresAccessor.ResolveFailure(f)
#                               if rslt == "Error":
#                                             return FailureProcessingResult.ProceedWithCommit
#                                             # return FailureProcessingResult.ProceedWithRollBack
#                               else:
#                                             return FailureProcessingResult.Continue



# try:
#                t2 = Transaction(doc, 'Regroup group')
#                t2.Start()

#                print(t2.GetStatus())

#                failureHandlingOptions = t2.GetFailureHandlingOptions()
#                failureHandler = FailureHandler()
#                failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
#                failureHandlingOptions.SetClearAfterRollback(True)
#                t2.SetFailureHandlingOptions(failureHandlingOptions)

#                Regroup(groupname,groupmember)

#                t2.Commit()
#                print(t2.GetStatus())

# except:
#                t2.RollBack()
               
#                print(t2.GetStatus())




# # COLLECT GENERIC MODEL ///////////////////
# genericModel_category = db.Collector(of_category = 'OST_GenericModel' , is_type = True)
# genericModel_element = genericModel_category.get_elements()
# genericModel_element_CategoryId = [model.Category.Id for model in  genericModel_element]
# # CREATE SCHEDULE VIEW ///////////////////////////////////////
# newSchedule = []
# with db.Transaction('create schedule view'):
# 	newSchedule.append(DB.ViewSchedule.CreateSchedule(revit.doc, genericModel_element_CategoryId[0]))
# 	schedulableFields = [sc.Definition.GetSchedulableFields() for sc in newSchedule]
# # //////////////////////////////////////////////////////






