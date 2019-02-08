import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import TextInput

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()

value = TextInput('Title', default="")
print(value)

def setParam(param, value):
	if str(param.StorageType) == "Integer":
		


# import clr
# import System
# import math
# clr.AddReference('RevitAPI')
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from System.Collections.Generic import List
# from pyrevit import forms
# from rpw.ui.forms import TextInput
# from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button)


# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# options = __revit__.Application.Create.NewGeometryOptions()
# SEBoptions = SpatialElementBoundaryOptions()
# roomcalculator = SpatialElementGeometryCalculator(doc)

# def Ungroup(group):
# 	group.UngroupMembers()

# def Regroup(groupname,groupmember):
# 	newgroup = doc.Create.NewGroup(groupmember)
# 	newgroup.GroupType.Name = str(groupname)

# def convertStr(s):
# 	"""Convert string to either int or float."""
# 	try:
# 		ret = int(s)
# 	except ValueError:
# 		ret = 0
# 	return ret

# class FailureHandler(IFailuresPreprocessor):
# 	def __init__(self):
# 		self.ErrorMessage = ""
# 		self.ErrorSeverity = ""
# 	def PreprocessFailures(self, failuresAccessor):
# 		# failuresAccessor.DeleteAllWarning()
# 		# return FailureProcessingResult.Continue
# 		failures = failuresAccessor.GetFailureMessages()
# 		rslt = ""
# 		for f in failures:
# 			fseverity = failuresAccessor.GetSeverity()
# 			if fseverity == FailureSeverity.Warning:
# 				failuresAccessor.DeleteWarning(f)
# 			elif fseverity == FailureSeverity.Error:
# 				rslt = "Error"
# 				failuresAccessor.ResolveFailure(f)
# 		if rslt == "Error":
# 			return FailureProcessingResult.ProceedWithCommit
# 			# return FailureProcessingResult.ProceedWithRollBack
# 		else:
# 			return FailureProcessingResult.Continue

# td_button = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No

# res = TaskDialog.Show("Importation from Excel","Attention :\n- Les ids des elements doivent etre en colonne 1\n- Les noms exacts (avec majuscules) des parametres partages doivent etre en ligne 1\n- Aucun accent ou caractere special dans le fichier Excel", td_button)

# if res == TaskDialogResult.Yes:

# 	# t = Transaction(doc, 'Read Excel spreadsheet.') 
# 	# t.Start()

# 	#Accessing the Excel applications.
# 	xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
# 	count = 1

# 	dicWs = {}
# 	count = 1
# 	for i in xlApp.Worksheets:
# 		dicWs[i.Name] = i
# 		count += 1

# 	components = [Label('Enter the name of ID parameter:'),
# 			ComboBox('combobox', dicWs),
# 			Label('Enter the number of rows in Excel you want to integrate to Revit:'),
# 			TextBox('textbox', Text="600"),
# 			Label('Enter the number of colones in Excel you want to integrate to Revit:'),
# 			TextBox('textbox2', Text="20"),
# 			Separator(),
# 			Button('Select')]
# 	form = FlexForm('Title', components)
# 	form.show()

# 	worksheet = form.values['combobox']
# 	rowEnd = convertStr(form.values['textbox'])
# 	colEnd = convertStr(form.values['textbox2'])

# 	#Row, and Column parameters
# 	rowStart = 1
# 	column_id = 1
# 	colStart = 2

# 	# Using a loop to read a range of values and print them to the console.
# 	array = []
# 	param_names_excel = []
# 	data = {}
# 	for r in range(rowStart, rowEnd):
# 		data_id = worksheet.Cells(r, column_id).Text
# 		data_id_int = convertStr(data_id)
# 		if data_id_int != 0:
# 			data = {'id': data_id_int}
# 			for c in range(colStart, colEnd):
# 				data_param_value = worksheet.Cells(r, c).Text
# 				data_param_name = worksheet.Cells(1, c).Text
# 				if data_param_name != '':
# 					param_names_excel.append(data_param_name)
# 					if data_param_value != '':
# 						data[data_param_name] = data_param_value
# 			array.append(data)

# 	# t.Commit()

# 	#Recuperation des portes
# 	doors = FilteredElementCollector(doc)\
# 		.OfCategory(BuiltInCategory.OST_Doors)\
# 		.WhereElementIsNotElementType()\
# 		.ToElements()

# 	#Get parameters in the model
# 	params_door_set = doors[0].Parameters
# 	params_door_name = []
# 	for param_door in params_door_set:
# 		params_door_name.append(param_door.Definition.Name)

# 	unfounddoors = []




# 	# TROUVER ERREUR ICI ET SUPPRIMER PARAGRAPHE!
# 	# print(array)
# 	# for hash in array:
# 	# 	for param in param_names_excel:
# 	# 		print(param)
# 	# 		if (param in params_door_name) and (param in hash):
# 	# 			# door.LookupParameter(param).Set(hash[param])
# 	# 			print(hash[param])
# 	# TROUVER ERREUR ICI ET SUPPRIMER PARAGRAPHE!




# 	# t = Transaction(doc, 'Feed doors')
# 	# t.Start()
# 	tg = TransactionGroup(doc, 'Feed doors')
# 	tg.Start()

# 	for hash in array:
# 		idInt = int(hash['id'])
# 		try :
# 			door_id = ElementId(idInt)
# 			door = doc.GetElement(door_id)
# 			groupId = door.GroupId

# 			print("Door : " + str(idInt))
# 			print("Group id : " + str(groupId))

# 			if str(groupId) != "-1":

# 				t1 = Transaction(doc, 'Ungroup group')
# 				t1.Start()

# 				group = doc.GetElement(groupId)
# 				groupname = group.Name
# 				groupmember = group.GetMemberIds()
# 				Ungroup(group)

# 				t1.Commit()

# 				print(t1.GetStatus())





# 			# TROUVER ERREUR ICI ET SUPPRIMER PARAGRAPHE!
# 			# for param in param_names_excel:
# 			# 	if (param in params_door_name) and (param in hash):
# 			# 		door.LookupParameter(param).Set(hash[param])

# 			for param in param_names_excel:
# 				try:
# 					if (param in params_door_name) and (param in hash):
# 						door.LookupParameter(param).Set(hash[param])
# 						print(param + " : Done")
# 				except:
# 					print(param + " : Failed")
# 			# TROUVER ERREUR ICI ET SUPPRIMER PARAGRAPHE!





# 			if str(groupId) != "-1":

# 				try:
# 					t2 = Transaction(doc, 'Regroup group')
# 					t2.Start()

# 					print(t2.GetStatus())

# 					failureHandlingOptions = t2.GetFailureHandlingOptions()
# 					failureHandler = FailureHandler()
# 					failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
# 					failureHandlingOptions.SetClearAfterRollback(True)
# 					t2.SetFailureHandlingOptions(failureHandlingOptions)

# 					Regroup(groupname,groupmember)

# 					t2.Commit()
# 					print(t2.GetStatus())

# 				except:
# 					t2.RollBack()
# 					print(t2.GetStatus())

# 					for i in groupmember:
# 						IdsInLine = IdsInLine + str(i.IntegerValue) + ", "

# 					IdsInLine = IdsInLine[:len(IdsInLine)-3]

# 					print("Regrouping failed on group : " + str(groupId) + " / " + str(groupname))
# 					print("Grouped element ids was : " + IdsInLine)

# 			print("Door " + str(idInt) + " : OK")

# 		except:
# 			print(str(idInt) + " not in REVIT doc")
# 			unfounddoors.append(idInt)

# 	print("Job done!")

# 	# t.Commit()

# 	tg.Assimilate()

# 	if len(unfounddoors) != 0:
# 		print(str(len(unfounddoors)) + " doors not found : ")
# 		print(unfounddoors)

# else:
# 	"A plus tard!"


# # """Etiquetage de poutres CANOPY!"""

# # # -*- coding: utf-8 -*-
# # e_a = str("\xe9")
# # a_a = str("\xe0")

# # __title__ = 'Beam tagging\nCANOPY'

# # __doc__ = 'Ce programme remplit les arases inferieures mini et maxi des poutres '\
# # 			'(parametres AI_Min et AI_Max) de la vue active du projet CANOPY'

# # # from pyrevit import revit, DB, UI
# # # from pyrevit import script
# # # from pyrevit import forms

# # import clr
# # import math
# # clr.AddReference('RevitAPI') 
# # clr.AddReference('RevitAPIUI') 
# # from Autodesk.Revit.DB import *
# # from Autodesk.Revit.UI import * 

# # doc = __revit__.ActiveUIDocument.Document

# # options = __revit__.Application.Create.NewGeometryOptions()

# # BP_collector = FilteredElementCollector(doc)\
# #           .OfCategory(BuiltInCategory.OST_ProjectBasePoint)\
# #           .WhereElementIsNotElementType()\
# #           .ToElements()

# # beam_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# #           .OfCategory(BuiltInCategory.OST_StructuralFraming)\
# #           .WhereElementIsNotElementType()\
# #           .ToElements()
		  
# # # beam_ID_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# #           # .OfCategory(BuiltInCategory.OST_StructuralFraming)\
# #           # .WhereElementIsNotElementType()\
# #           # .ToElementIds()
		  
# # # beam_tag_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# # 						# .OfCategory(BuiltInCategory.OST_StructuralFramingTags)\
# # 						# .WhereElementIsNotElementType()\
# # 						# .ToElements()
						
# # # TaggegBeamsList = []					
# # # for bt in beam_tag_collector:
# # 	# if bt.TaggedLocalElementId in beam_ID_collector:
# # 		# TaggegBeamsList.append(bt.TaggedLocalElementId)
		
# # td_button = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Cancel | TaskDialogCommonButtons.No

# # res = TaskDialog.Show("Etiquetage de poutres","Voulez-vous lancer l'"+e_a+"tiquetage des poutres dans la vue active?",td_button)

# # if res == TaskDialogResult.Yes:

# # 	for BP in BP_collector:
# # 		# zBP = BP.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()/3.2808399
# # 		zBP = 0

# # 	print(zBP)

# # 	tg = TransactionGroup(doc, 'Tag all beams in active view')

# # 	tg.Start()

# # 	for beam in beam_collector:
# # 		# if beam.Id not in TaggegBeamsList:
# # 			try:
# # 				print("Nom de la poutre : " + beam.Name)
# # 				print("ID : " + str(beam.Id))
# # 				print("Niveau de r"+e_a+"f"+e_a+"rence : " + str(beam.LookupParameter('Niveau de r'+e_a+'f'+e_a+'rence').AsValueString()))
# # 				start = beam.Location.Curve.GetEndPoint(0)
# # 				print(start)
# # 				end = beam.Location.Curve.GetEndPoint(1)
# # 				print(end)
# # 				z_min = -zBP + ((beam.get_Geometry(options).GetBoundingBox().Min).Z)/3.2808399
# # 				print("AI_Inf : " + str(z_min) + "m")
				
# # 				# Si la poutre est en pente
# # 				if beam.LookupParameter('El'+e_a+'vation '+a_a+' la base').AsDouble() == 0:
# # 					print("EN PENTE")
# # 					t = Transaction(doc, 'Tag beam')
# # 					t.Start()
# # 					try:
# # 						delta = abs(float(beam.LookupParameter("D"+e_a+"calage du niveau de d"+e_a+"part").AsValueString())-float(beam.LookupParameter("D"+e_a+"calage du niveau d'arriv"+e_a+"e").AsValueString()))\
# # 									+abs(float(beam.LookupParameter("Valeur de d"+e_a+"calage de l'extr"+e_a+"mit"+e_a+" Z").AsValueString())-float(beam.LookupParameter("Valeur de d"+e_a+"calage Z de d"+e_a+"but").AsValueString()))
# # 					except:
# # 						delta = abs(float(beam.LookupParameter("D"+e_a+"calage du niveau de d"+e_a+"part").AsValueString())-float(beam.LookupParameter("D"+e_a+"calage du niveau d'arriv"+e_a+"e").AsValueString()))
# # 					z_max = z_min + delta
# # 					print("AI_Max : " + str(z_max) + "m")
# # 					beam.LookupParameter('AI_Min').Set(" ")
# # 					beam.LookupParameter('AI_Min').Set(str(round(z_min,2)))
# # 					beam.LookupParameter('AI_Max').Set(" ")
# # 					beam.LookupParameter('AI_Max').Set(str(round(z_max,2)))
# # 					cen=XYZ((start.X+end.X)/2,(start.Y+end.Y)/2,(z_min+z_max)/2)
# # 					print(cen)
# # 					print("\n")
# # 					beam_tag = doc.Create.NewTag(doc.ActiveView,beam,False,TagMode.TM_ADDBY_CATEGORY,TagOrientation.Horizontal,cen)
# # 					t.Commit()
# # 				# Si la poutre est horizontale
# # 				else :
# # 					print("HORIZONTALE")
# # 					t = Transaction(doc, 'Tag beam')
# # 					t.Start()
# # 					beam.LookupParameter('AI_Min').Set(" ")
# # 					beam.LookupParameter('AI_Min').Set(str(round(z_min,2)))
# # 					beam.LookupParameter('AI_Max').Set(" ")
# # 					cen=XYZ((start.X+end.X)/2,(start.Y+end.Y)/2,z_min)
# # 					print(cen)
# # 					print("\n")
# # 					beam_tag = doc.Create.NewTag(doc.ActiveView,beam,False,TagMode.TM_ADDBY_CATEGORY,TagOrientation.Horizontal,cen)
# # 					t.Commit()
			
# # 			except:
# # 				print(" ")
				
# # 	tg.Assimilate()
	
# # 	td_button2 = TaskDialogCommonButtons.Ok

# # 	res2 = TaskDialog.Show("Mise au propre","Veuillez checker la non superposition des "+e_a+"tiquettes pour un r"+e_a+"sultat plus lisible!",td_button2)
				
# # else:
# #   print("Une autre fois peut-etre...")