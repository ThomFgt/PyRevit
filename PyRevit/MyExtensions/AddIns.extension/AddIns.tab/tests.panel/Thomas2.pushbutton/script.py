# -*- coding: utf-8 -*-

__title__ = 'Révisions'

__doc__ = 'Ce programme ajoute les révisions disponibles sur les feuilles sélectionnées'

# IMPORT
# import math
import System
import rpw
import clr
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference('System')
from System.Collections.Generic import List
from rpw import revit, db, ui, DB, UI
# from rpw.ui.forms import *
# from System import Array
# from System.Runtime.InteropServices import Marshal
import sys
import Autodesk
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System.Drawing import *
from System.Windows.Forms import *


# COLLECT VIEWS MAQUETTE
view_category = DB.FilteredElementCollector(revit.doc).OfCategory(BuiltInCategory.OST_Sheets)
view_element = view_category.WhereElementIsNotElementType()
view_name = [views.ViewName for views in view_element]
view_id = [views.Id for views in view_element]

# COLLECT REVISION MAQUETTE
revisions_category = DB.FilteredElementCollector(revit.doc).OfCategory(BuiltInCategory.OST_Revisions)
revision_element = revisions_category.WhereElementIsNotElementType()
revision_name = [revisions.Name for revisions in revision_element]
revision_id = [revisions.Id for revisions in revision_element]
revision_description = [revisions.Description for revisions in revision_element]
revision_date = [revisions.RevisionDate for revisions in revision_element]


sheets_out = []
sheets_id_out = []
index = {}
sheets_id_name = {}
revision_id_name = {}

for k in range(len(view_name)):
	sheets_id_name[str(view_name[k])] = view_id[k]

for k in range(len(revision_name)):
	revision_id_name[str(revision_name[k])] = revision_id[k]


# /////////////////////////////////////////////////////////////////////////
# START WINDOWS FORM
class MyListBox(Form):
	# docstring for ClassName
	def __init__(self):
		self.Text = "Plugin revision on sheets"
		self.Size = Size(600,600)
		self.CenterToScreen()
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		

	# label for listbox 1 on the left
		label1 = Label()
		label1.Parent = self
		label1.Text = "Double clic sur les feuilles"
		label1.Location = Point(5,5)
		label1.Size = Size(180,15)
	# listbox 1 on the left
		self.listbox1 = ListBox()
		self.listbox1.Parent = self
		self.listbox1.Location = Point(5,30)
		self.listbox1.Size = Size(280,500)
		for k in view_name:
			self.listbox1.Items.Add(k)
		self.listbox1.DoubleClick += self.OnDoubleClick


	# label for listbox 2 on the right
		label2 = Label()
		label2.Parent = self
		label2.Size = Size(180,15)
		label2.Text = "Liste des feuilles sélectionnées"
		label2.Location = Point(300,5)
	# listbox 2 on the right
		self.listbox2 = ListBox()
		self.listbox2.Parent = self
		self.listbox2.Location = Point(300,30)
		self.listbox2.Size = Size(280,300)
		for k in sheets_out:
			self.listbox2.Items.Add(k)
		self.listbox2.DoubleClick += self.OnDoubleClick2



	# button OK
		self.buttonOk = Button()
		self.buttonOk.Parent = self
		self.buttonOk.Text = "Ok Chef"
		self.buttonOk.Location = Point(500,520)
		self.buttonOk.Click += self.Ok
	# button Cancel
		self.buttonCancel = Button()
		self.buttonCancel.Parent = self
		self.buttonCancel.Text = "ANNULE"
		self.buttonCancel.Location = Point(400,520)
		self.buttonCancel.Click += self.Cancel




	# EVENT ON CLICK //////
	# Event on double click on listbox1
	def OnDoubleClick(self,sender,event):
		item = sender.SelectedItem
		index[str(item)] = sender.SelectedIndex
		self.listbox2.Items.Add(item)
		self.listbox1.Items.Remove(item)


	# Event on double click on listbox2
	def OnDoubleClick2(self,sender,event):
		item = sender.SelectedItem
		self.listbox1.Items.Insert(index.get(item), item)
		self.listbox2.Items.Remove(item)
		

	# Event on Cancel Button Click
	def Cancel(self,sender,event):
		self.Close()

	# Event on OK Button Click
	def Ok(self,sender,event):
		for i in range(self.listbox2.Items.Count):
			sheets_out.append(self.listbox2.Items[i])
		self.Close()

# run form
Application.Run(MyListBox())
# END OF WINDOWS FORM ///////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////////


revision_out = []
# /////////////////////////////////////////////////////////////////////////
# START WINDOWS FORM
class MyListBox2(Form):
	# docstring for ClassName
	def __init__(self):
		self.Text = "Plugin revision on sheets"
		self.Size = Size(600,600)
		self.CenterToScreen()
		self.FormBorderStyle = FormBorderStyle.FixedToolWindow
		

	# label for listbox 1 on the left
		label1 = Label()
		label1.Parent = self
		label1.Text = "Choisir les révisions à faire apparaitre sur les feuilles sélectionnées précédemment"
		label1.Location = Point(5,5)
		label1.Size = Size(500,30)
	# listbox 1 on the left
		self.checkedlistbox1 = CheckedListBox()
		self.checkedlistbox1.Parent = self
		self.checkedlistbox1.Location = Point(30,40)
		self.checkedlistbox1.Size = Size(400,100)
		for k in revision_name:
			self.checkedlistbox1.Items.Add(k)


	# button OK
		self.buttonOk = Button()
		self.buttonOk.Parent = self
		self.buttonOk.Text = "Ok Chef"
		self.buttonOk.Location = Point(500,520)
		self.buttonOk.Click += self.Ok
	# button Cancel
		self.buttonCancel = Button()
		self.buttonCancel.Parent = self
		self.buttonCancel.Text = "ANNULE"
		self.buttonCancel.Location = Point(400,520)
		self.buttonCancel.Click += self.Cancel



	# EVENT ON CLICK //////
	# Event on Cancel Button Click
	def Cancel(self,sender,event):
		self.Close()

	# Event on OK Button Click
	def Ok(self,sender,event):
		for i in range(self.checkedlistbox1.CheckedItems.Count):
			revision_out.append(self.checkedlistbox1.CheckedItems[i])
		self.Close()

# run form
Application.Run(MyListBox2())
# END OF WINDOWS FORM ///////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////////


# Get sheet element selected by user
sheet_out_elem = []
for item in sheets_out:
	sheet_out_elem.append(DB.Document.GetElement(revit.doc, sheets_id_name[item]))

# Get collection element id of revision selected by user
revision_out_id = []
for item in revision_out:
	revision_out_id.append(revision_id_name[item])
collection = List[DB.ElementId](revision_out_id)

# Transaction
with db.Transaction("set new revision"):
	for sheets in sheet_out_elem:
		sheets.SetAdditionalRevisionIds(collection)



