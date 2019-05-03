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



# # ///////////////////////////////////////////////////////////////////////
# # CREATE FILTER FOR WORKSET
# with db.Transaction('create filter'):
# 	new_workset = DB.Workset.Create(revit.doc, "nouveauSousProjet")
# 	worksetId = str(new_workset.Id)
# 	elemId = DB.ElementId(int(worksetId))

# 	all_Categories_Id = System.Enum.GetValues(BuiltInCategory)
# 	catlist = [ElementId(i) for i in all_Categories_Id]
# 	typeCatList = List[DB.ElementId](catlist)

# 	workset_parameter = DB.ElementId(DB.BuiltInParameter.ELEM_PARTITION_PARAM)

# 	rule = [DB.ParameterFilterRuleFactory.CreateEqualsRule(workset_parameter,elemId)]

# 	new_filter = DB.ParameterFilterElement.Create(revit.doc,"nouveauFiltre",typeCatList,rule)
# # ////////////////////////////////////////////////////////////////////////



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

for j in range(len(linkDoc)):
	try:
		collector_link = DB.FilteredWorksetCollector(linkDoc[j])
		collect_worksets_link = collector_link.OfKind(DB.WorksetKind.UserWorkset)
		worksets_name_link = [workset.Name for workset in collect_worksets_link]
	except : pass
	
#COLOR
color_list = ["bleu","rouge","vert","jaune","orange","marron","rose","violet","bleu clair"]


# /////////////////////////////////////////////////////////////////////////
class MyListBox(Form):
	"""docstring for ClassName"""
	def __init__(self):
		self.Text = "Ma form cool"
		screenSize = Screen.GetWorkingArea(self)
		self.Height = screenSize.Height / 2
		self.Width = screenSize.Width / 3

		label1 = Label()
		label1.Parent = self
		label1.Text = "la liste 1"
		label1.Location = Point(5,5)

		listbox1 = ListBox()
		listbox1.Parent = self
		listbox1.Location = Point(5,30)
		listbox1.Size = Size(180,100)
		for k in linkDoc_name:
			listbox1.Items.Add(k)



		label2 = Label()
		label2.Parent = self
		label2.Text = "la liste 2"
		label2.Location = Point(200,5)

		listbox2 = ListBox() 
		listbox2.Parent = self
		listbox2.Location = Point(200,30)
		listbox2.Size = Size(180,100)
		for k in worksets_name_link:
			listbox2.Items.Add(k)



		label3 = Label()
		label3.Parent = self
		label3.Text = "la liste 3"
		label3.Location = Point(400,5)

		listbox3 = ListBox()
		listbox3.Parent = self
		listbox3.Location = Point(400,30)
		listbox3.Size = Size(180,100)
		for k in color_list:
			listbox3.Items.Add(k)



		buttonAdd = Button()
		buttonAdd.Parent = self
		buttonAdd.Text = "Ajouter filtre"
		buttonAdd.Location = Point(450,150)
		buttonAdd.Size = Size(100,25)


		buttonOk = Button()
		buttonOk.Parent = self
		buttonOk.Text = "Ok Chef"
		buttonOk.Location = Point(500,420)


		buttonCancel = Button()
		buttonCancel.Parent = self
		buttonCancel.Text = "ANNULE"
		buttonCancel.Location = Point(400,420)


		groupBox1 = GroupBox()
		groupBox1.Parent = self
		groupBox1.Text = "Détail des filtres"
		groupBox1.Location = Point(50,200)
		groupBox1.Size = Size(500,200)

		
		for k in range(1,9):
			label1 = Label()
			label1.Parent = groupBox1
			label1.Text = " "
			label1.Location = Point(40,20*k)
			label1.Size = Size(120,20)
			label1.BorderStyle = BorderStyle.Fixed3D

			label2 = Label()
			label2.Parent = groupBox1
			label2.Text = " "
			label2.Location = Point(200,20*k)
			label2.Size = Size(120,20)
			label2.BorderStyle = BorderStyle.Fixed3D
			
			label3 = Label()
			label3.Parent = groupBox1
			label3.Text = " "
			label3.Location = Point(350,20*k)
			label3.Size = Size(120,20)
			label3.BorderStyle = BorderStyle.Fixed3D


		self.CenterToScreen()

Application.Run(MyListBox())
# ///////////////////////////////////////////////////////

# class HelloWorld3Form(Form):
#     def __init__(self):
#         self.Text = "Hello World 3"
#         self.FormBorderStyle = FormBorderStyle.FixedDialog

#         screenSize = Screen.GetWorkingArea(self)
#         self.Height = screenSize.Height / 3
#         self.Width = screenSize.Width / 3

#         self.panelHeight = self.ClientRectangle.Height / 2

#         self.setupPanel1()
#         self.setupPanel2()
#         self.setupCounters()

#         self.Controls.Add(self.panel1)
#         self.Controls.Add(self.panel2)

#     def setupPanel1(self):
#         self.panel1 = Panel()
#         self.panel1.BackColor = Color.LightSlateGray
#         self.panel1.ForeColor = Color.Blue
#         self.panel1.Width = self.Width
#         self.panel1.Height = self.panelHeight
#         self.panel1.Location = Point(0, 0)
#         self.panel1.BorderStyle = BorderStyle.FixedSingle

#         self.label1 = Label()
#         self.label1.Text = "Go On - Press Me"
#         self.label1.Location = Point(20, 20)
#         self.label1.Height = 25
#         self.label1.Width = 175

#         self.button1 = Button()
#         self.button1.Name = '1'
#         self.button1.Text = 'Press Me 1'
#         self.button1.Location = Point(20, 50)
#         self.button1.Click += self.update

#         self.panel1.Controls.Add(self.label1)
#         self.panel1.Controls.Add(self.button1)

#     def setupPanel2(self):
#         self.panel2 = Panel()
#         self.panel2.BackColor = Color.LightSalmon
#         self.panel2.Width = self.Width
#         self.panel2.Height = self.panelHeight
#         self.panel2.Location = Point(0, self.panelHeight)
#         self.panel2.BorderStyle = BorderStyle.FixedSingle

#         self.subpanel1 = Panel()
#         self.subpanel1.BackColor = Color.Wheat
#         self.subpanel1.Width = 175
#         self.subpanel1.Height = 100
#         self.subpanel1.Location = Point(25, 25)
#         self.subpanel1.BorderStyle = BorderStyle.Fixed3D

#         self.label2 = Label()
#         self.label2.Text = "Go On - Press Me"
#         self.label2.Location = Point(20, 20)
#         self.label2.Height = 25
#         self.label2.Width = 175

#         self.button2 = Button()
#         self.button2.Name = '2'
#         self.button2.Text = 'Press Me 2'
#         self.button2.Location = Point(20, 50)
#         self.button2.Click += self.update

#         self.subpanel1.Controls.Add(self.label2)
#         self.subpanel1.Controls.Add(self.button2)

#         self.subpanel2 = Panel()
#         self.subpanel2.BackColor = Color.Transparent
#         self.subpanel2.Width = 175
#         self.subpanel2.Height = 100
#         self.subpanel2.Location = Point(220, 25)
#         self.subpanel2.BorderStyle = BorderStyle.Fixed3D

#         self.label3 = Label()
#         self.label3.Text = "Go On - Press Me"
#         self.label3.Location = Point(20, 20)
#         self.label3.Height = 25
#         self.label3.Width = 175

#         self.button3 = Button()
#         self.button3.Name = '3'
#         self.button3.Text = 'Press Me 3'
#         self.button3.Location = Point(20, 50)
#         self.button3.Click += self.update

#         self.subpanel2.Controls.Add(self.label3)
#         self.subpanel2.Controls.Add(self.button3)

#         self.panel2.Controls.Add(self.subpanel1)
#         self.panel2.Controls.Add(self.subpanel2)

#     def setupCounters(self):
#         self.counterDict = {
#             '1': 0,
#             '2': 0,
#             '3': 0,
#         }

#     def update(self, sender, event):
#         name = sender.Name
#         self.counterDict[name] += 1
#         label = getattr(self, 'label' + name)
#         label.Text = "You have pressed me %s times." % self.counterDict[name]


# form = HelloWorld3Form()
# Application.Run(form)

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






