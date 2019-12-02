# -*- coding: utf-8 -*-

__title__ = 'Add Prefix'

__doc__ = "Le programme ajoute un préfixe à tous les sous-projets de la maquette. "


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


# #COLLECT WORKSET IN DOC
collector_worksets = DB.FilteredWorksetCollector(revit.doc)
user_worksets = collector_worksets.OfKind(DB.WorksetKind.UserWorkset)

pfixe = []
#START FORM
class MyListBox(Form):
	# docstring for ClassName
	def __init__(self):
		self.Text = "Add prefix to worksets"
		self.CenterToScreen()
		screenSize = Screen.GetWorkingArea(self)
		self.Height = screenSize.Height / 4
		self.Width = screenSize.Width / 6

		self.lab = Label()
		self.lab.Parent = self
		self.lab.Text = "Ecrire le préfixe à ajouter devant tous les sous projets"
		self.lab.Location = Point(10,20)
		self.lab.Size = Size(280,30)

		self.tb = TextBox()
		self.tb.Parent = self
		self.tb.Location = Point(20,70)
		self.tb.Size = Size(200,30)

		self.btok = Button()
		self.btok.Parent = self
		self.btok.Text = "Ok Chef"
		self.btok.Location = Point(120,210)
		self.btok.Click += self.Ok

		self.btc = Button()
		self.btc.Parent = self
		self.btc.Text = "ANNULE"
		self.btc.Location = Point(200,210)
		self.btc.Click += self.Cancel

	# Event on Cancel Button Click
	def Cancel(self,sender,event):
		self.Close()

	# Event on OK Button Click
	def Ok(self,sender,event):
		pfixe.append(self.tb.Text)
		self.Close()

#Run form
Application.Run(MyListBox())
#End of form


#Transaction
if revit.doc.IsWorkshared :
	with db.Transaction('rename worksets'):
		for workset in user_worksets:
			DB.WorksetTable.RenameWorkset(revit.doc, workset.Id, pfixe[0]+workset.Name)
	print "Le préfixe a bien été ajouté à tous les sous-projets.\nCtrl + Z pour annuler"
else : print "Le projet n'est pas en collaboration"
