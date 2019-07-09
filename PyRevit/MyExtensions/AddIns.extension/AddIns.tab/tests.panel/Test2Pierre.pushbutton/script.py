# # # # # Clash detection 2 # # # # #

import clr
import System
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
copyOptions = CopyPasteOptions()

class Bubble:
	doc = __revit__.ActiveUIDocument.Document
	uidoc = __revit__.ActiveUIDocument
	options = __revit__.Application.Create.NewGeometryOptions()

	def __init__(self,bubble_name):
		self.familysymbol = None		
		self.presence = False
		
		gm_collector = FilteredElementCollector(doc)\
			  .OfClass(FamilySymbol)\
			  .OfCategory(BuiltInCategory.OST_GenericModel)
				
		for i in gm_collector:
			if i.FamilyName == str(bubble_name):
				self.familysymbol = i
				self.presence = True
				break
				
		return self.familysymbol
		return self.presence

def get_selected_elements(doc):
    try:
        # # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)

def DocToOriginTransform(doc):
	try:
		projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	except:
		projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero)
	translationVector = XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation)
	translationTransform = Transform.CreateTranslation(translationVector)
	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition.Angle, XYZ.Zero)
	finalTransform = translationTransform.Multiply(rotationTransform)
	return finalTransform

def OriginToDocTransform(doc):
	try:
		projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	except:
		projectPosition = doc.ActiveProjectLocation.GetProjectPosition(XYZ.Zero)
	translationVector = XYZ(-projectPosition.EastWest, -projectPosition.NorthSouth, -projectPosition.Elevation)
	translationTransform = Transform.CreateTranslation(translationVector)
	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, -projectPosition.Angle, XYZ.Zero)
	finalTransform = rotationTransform.Multiply(translationTransform)
	return finalTransform

def DocToDocTransform(doc1, doc2):
	docToOriginTrans = DocToOriginTransform(doc1)
	originToDocTrans = OriginToDocTransform(doc2)
	docToDocTrans = originToDocTrans.Multiply(docToOriginTrans)
	return docToDocTrans

def get_solid(element):
	solid_list = []
	for i in element.get_Geometry(options):
		if i.ToString() == "Autodesk.Revit.DB.Solid":
			solid_list.append(SolidUtils.CreateTransformed(i, DocToDocTransform(element.Document, __revit__.ActiveUIDocument.Document)))
		elif i.ToString() == "Autodesk.Revit.DB.GeometryInstance":
			for j in i.GetInstanceGeometry():
				if j.ToString() == "Autodesk.Revit.DB.Solid":
					solid_list.append(SolidUtils.CreateTransformed(j, DocToDocTransform(element.Document, __revit__.ActiveUIDocument.Document)))
	return solid_list

def get_intersection(el1, el2):
	bb1 = el1.get_Geometry(options).GetBoundingBox()
	bb2 = el2.get_Geometry(options).GetBoundingBox()

	trans1 = bb1.Transform
	trans2 = bb2.Transform

	min1 = trans1.OfPoint(bb1.Min)
	max1 = trans1.OfPoint(bb1.Max)
	min2 = trans2.OfPoint(bb2.Min)
	max2 = trans2.OfPoint(bb2.Max)

	outline1 = Outline(min1, max1)
	outline2 = Outline(min2, max2)

	# print(outline1.Intersects(outline2,0))
	# print(outline1.ContainsOtherOutline(outline2,0))
	# print(outline2.ContainsOtherOutline(outline1,0))

	solid1_list = get_solid(el1)
	solid2_list = get_solid(el2)

	for i in solid1_list:
		for j in solid2_list:
			try:
				inter = BooleanOperationsUtils.ExecuteBooleanOperation(i, j, BooleanOperationsType.Intersect)
				if inter.Volume != 0:
					interBb = inter.GetBoundingBox()
					interTrans = interBb.Transform
					interPoint = interTrans.OfPoint(interBb.Min)
					break
			except:
				"Oh god!"

	try:
		interPoint
		return interPoint
	except:
		return None

def get_elements_in_3Dview():
	cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
				BuiltInCategory.OST_CableTray, BuiltInCategory.OST_StructuralFraming,\
				BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting,\
				BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_FlexPipeCurves]
	
	rvtlink_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
			.OfCategory(BuiltInCategory.OST_RvtLinks)\
			.WhereElementIsNotElementType()\
			.ToElements()

	SectionBox = doc.ActiveView.GetSectionBox()
	trans = SectionBox.Transform
	bbmin = SectionBox.Min
	bbmax = SectionBox.Max

	dl_list = []
	for rvtlink in rvtlink_collector:
		dl_list.append(rvtlink)
	dl_list.append(doc)

	element_list = []
	for cat in cat_list:
		for dl in dl_list:
			if dl.GetType().ToString() == "Autodesk.Revit.DB.Document":
				try:
					outline = Outline(trans.OfPoint(bbmin),trans.OfPoint(bbmax))
					bbFilter = BoundingBoxIntersectsFilter(outline)
				except:
					bbmin2 = trans.OfPoint(bbmin)
					bbmax2 = trans.OfPoint(bbmax)
					outmin = XYZ(min(bbmin2.X,bbmax2.X), min(bbmin2.Y,bbmax2.Y), min(bbmin2.Z,bbmax2.Z))
					outmax = XYZ(max(bbmin2.X,bbmax2.X), max(bbmin2.Y,bbmax2.Y), max(bbmin2.Z,bbmax2.Z))
					outline = Outline(outmin, outmax)
					bbFilter = BoundingBoxIntersectsFilter(outline)
				element_collector = FilteredElementCollector(dl)\
					.OfCategory(cat)\
					.WherePasses(bbFilter)\
					.WhereElementIsNotElementType()\
					.ToElements()
				for e in element_collector:
					if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.Category.Id.IntegerValue == -2001320) \
							and (e.StructuralType.ToString() != "Beam"):
						pass
					else:
						element_list.append(e)
			elif dl.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
				docToOriginTrans = DocToOriginTransform(doc)
				originToDocTrans = OriginToDocTransform(dl.GetLinkDocument())
				docToDocTrans = originToDocTrans.Multiply(docToOriginTrans)
				try:
					outline = Outline(docToDocTrans.OfPoint(trans.OfPoint(bbmin)), docToDocTrans.OfPoint(trans.OfPoint(bbmax)))
					bbFilter = BoundingBoxIntersectsFilter(outline)
				except:
					bbmin2 = docToDocTrans.OfPoint(trans.OfPoint(bbmin))
					bbmax2 = docToDocTrans.OfPoint(trans.OfPoint(bbmax))
					outmin = XYZ(min(bbmin2.X,bbmax2.X), min(bbmin2.Y,bbmax2.Y), min(bbmin2.Z,bbmax2.Z))
					outmax = XYZ(max(bbmin2.X,bbmax2.X), max(bbmin2.Y,bbmax2.Y), max(bbmin2.Z,bbmax2.Z))
					outline = Outline(outmin, outmax)
					# outline = Outline(trans.OfPoint(bbmin), trans.OfPoint(bbmax))
					bbFilter = BoundingBoxIntersectsFilter(outline)
				element_collector = FilteredElementCollector(dl.GetLinkDocument())\
					.OfCategory(cat)\
					.WherePasses(bbFilter)\
					.WhereElementIsNotElementType()\
					.ToElements()
				for e in element_collector:
					if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.Category.Id.IntegerValue == -2001320) \
							and (e.StructuralType.ToString() != "Beam"):
						pass
					else:
						element_list.append(e)
	return element_list

def get_docsolid(element, elementdoc, activedoc):
	docToOriginTrans = DocToOriginTransform(elementdoc)
	originToDocTrans = OriginToDocTransform(activedoc)

	bb = element.get_Geometry(options).GetBoundingBox()
	trans = bb.Transform
	bbmin = originToDocTrans.OfPoint(docToOriginTrans.OfPoint(trans.OfPoint(bb.Min)))
	bbmax = originToDocTrans.OfPoint(docToOriginTrans.OfPoint(trans.OfPoint(bb.Max)))

	pt0 = XYZ(bbmin.X, bbmin.Y, bbmin.Z)
	pt1 = XYZ(bbmax.X, bbmin.Y, bbmin.Z)
	pt2 = XYZ(bbmax.X, bbmax.Y, bbmin.Z)
	pt3 = XYZ(bbmin.X, bbmax.Y, bbmin.Z)

	edge0 = Line.CreateBound(pt0, pt1)
	edge1 = Line.CreateBound(pt1, pt2)
	edge2 = Line.CreateBound(pt2, pt3)
	edge3 = Line.CreateBound(pt3, pt0)

	iCurveCollection = List[Curve]()
	iCurveCollection.Add(edge0)
	iCurveCollection.Add(edge1)
	iCurveCollection.Add(edge2)
	iCurveCollection.Add(edge3)
	height = bbmax.Z - bbmin.Z
	baseLoop = CurveLoop.Create(iCurveCollection)

	iCurveLoopCollection = List[CurveLoop]()
	iCurveLoopCollection.Add(baseLoop)

	solid = GeometryCreationUtilities.CreateExtrusionGeometry(iCurveLoopCollection, XYZ.BasisZ, height)
	
	return solid

def get_docoutline(element, elementdoc, activedoc):
	docToOriginTrans = DocToOriginTransform(elementdoc)
	originToDocTrans = OriginToDocTransform(activedoc)

	bb = element.get_Geometry(options).GetBoundingBox()
	trans = bb.Transform
	bbmin = originToDocTrans.OfPoint(docToOriginTrans.OfPoint(trans.OfPoint(bb.Min)))
	bbmax = originToDocTrans.OfPoint(docToOriginTrans.OfPoint(trans.OfPoint(bb.Max)))

	outline = Outline(bbmin, bbmax)

	return outline

bname = "OBSERVATIONS_2017"
# bname = "Observation_Libelle SYA"

element_list = get_elements_in_3Dview()
print(element_list)
tuple_list = []
point_list = []
k = 0
l = -1
for i in element_list:
	k = k + 1
	for j in range(k, len(element_list)):
		if (i.Category.Id.IntegerValue == -2001320) and (element_list[j].Category.Id.IntegerValue == -2001320):
			"Beam intersection"
		else:
			interPoint = get_intersection(i, element_list[j])
			if interPoint is None:
				pass
			else:
				if interPoint.ToString() not in tuple_list:
					point_list.append(interPoint)
					tuple_list.append(interPoint.ToString())

t = Transaction(doc, 'Place bubbles')
t.Start()
for m in point_list:
	print("Bubble placed")
	instance = doc.Create.NewFamilyInstance(m, Bubble(bname).familysymbol, Structure.StructuralType.NonStructural)
t.Commit()






# bname = "OBSERVATIONS_2017"

# if Bubble(bname).presence is True:
# 	cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
# 				BuiltInCategory.OST_CableTray, BuiltInCategory.OST_StructuralFraming,\
# 				BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting,\
# 				BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_FlexPipeCurves]
	
# 	rvtlink_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 			.OfCategory(BuiltInCategory.OST_RvtLinks)\
# 			.WhereElementIsNotElementType()\
# 			.ToElements()

# 	SectionBox = doc.ActiveView.GetSectionBox()
# 	trans = SectionBox.Transform
# 	bbmin = SectionBox.Min
# 	bbmax = SectionBox.Max
	
# 	# outline = Outline(trans.OfPoint(min),trans.OfPoint(max))
	
# 	dl_list = []
# 	for rvtlink in rvtlink_collector:
# 		dl_list.append(rvtlink)
# 	dl_list.append(doc)

# 	element_list = []
# 	for cat in cat_list:
# 		for dl in dl_list:
# 			if dl.GetType().ToString() == "Autodesk.Revit.DB.Document":
# 				outline = Outline(trans.OfPoint(bbmin),trans.OfPoint(bbmax))
# 				bbFilter = BoundingBoxIntersectsFilter(outline)
# 				element_collector = FilteredElementCollector(dl)\
# 					.OfCategory(cat)\
# 					.WherePasses(bbFilter)\
# 					.WhereElementIsNotElementType()\
# 					.ToElements()
# 				for e in element_collector:
# 					if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
# 						pass
# 					else:
# 						element_list.append(e)
# 			elif dl.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
# 				docToOriginTrans = DocToOriginTransform(doc)
# 				originToDocTrans = OriginToDocTransform(dl.GetLinkDocument())
# 				docToDocTrans = originToDocTrans.Multiply(docToOriginTrans)
# 				try:
# 					outline = Outline(docToDocTrans.OfPoint(trans.OfPoint(bbmin)), docToDocTrans.OfPoint(trans.OfPoint(bbmax)))
# 					bbFilter = BoundingBoxIntersectsFilter(outline)
# 				except:
# 					a = "Same base point"
# 					outline = Outline(trans.OfPoint(bbmin), trans.OfPoint(bbmax))
# 					bbFilter = BoundingBoxIntersectsFilter(outline)
# 				element_collector = FilteredElementCollector(dl.GetLinkDocument())\
# 					.OfCategory(cat)\
# 					.WherePasses(bbFilter)\
# 					.WhereElementIsNotElementType()\
# 					.ToElements()
# 				for e in element_collector:
# 					if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
# 						pass
# 					else:
# 						element_list.append(e)

# 	print(element_list)
	
# 	tuple_list = []
# 	point_list = []
# 	k = 0
# 	l = -1
# 	for i in element_list:
# 		l = l + 1
# 		k = k + 1
# 		for j in range(k, len(element_list)):
# 			if interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document) is None:
# 				pass
# 			else:
# 				if interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document).ToString() not in tuple_list: 
# 					point_list.append(interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document))
# 					tuple_list.append(interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document).ToString())
	
# 	t = Transaction(doc, 'Place bubbles')
# 	t.Start()
# 	for m in point_list:
# 		print("Bubble placed")
# 		instance = doc.Create.NewFamilyInstance(m, Bubble(bname).familysymbol, Structure.StructuralType.NonStructural)
# 	t.Commit()

# # # # # Clash detection 2 # # # # #











# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# el = get_selected_elements(doc)[0]
# print("el:")
# print(el)
# print("dir(el)")
# print(dir(el))
# print("dir(el.location)")
# print(dir(el.Location))
# bb = el.get_Geometry(options).GetBoundingBox()
# print(bb)
# print("bbmin")
# print(bb.Min)
# print("bbmax")
# print(bb.Max)
# print("dir(bb)")
# print(dir(bb))
# print("category")
# print(el.Category.Name)
# print("el geometry object from reference")
# a = el.GetGeometryObjectFromReference(Reference(el))
# print(el.GetGeometryObjectFromReference(Reference(el)))
# print("dir(a)")
# print(dir(a))
# print(bb.Bounds)

# ceiling_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# 	.OfCategory(BuiltInCategory.OST_Ceilings)\
# 	.WhereElementIsNotElementType()\
# 	.ToElements()

# hsfp_list = []
# for i in ceiling_collector:
# 	z = round(i.LookupParameter("D"+"\xe9"+"calage par rapport au niveau").AsDouble()/3.2808399, 2)
# 	if (z not in hsfp_list) and (i.LookupParameter("Sous-projet").AsValueString() == "BPS_FXP"):
# 		print(i.LookupParameter("Sous-projet").AsValueString())
# 		print(z)
# 		hsfp_list.append(z)
	
# for j in hsfp_list:
# 	print(j)

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


# >>>>>>>>>>>>>>>


# beam_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# 	.OfCategory(BuiltInCategory.OST_StructuralFraming)\
# 	.WhereElementIsNotElementType()\
# 	.ToElements()

# icollection = List[ElementId]()
	
# uidoc.Selection.SetElementIds(icollection)

# for beam in beam_collector:
# 	if beam.StructuralType.ToString() == "Beam":
# 		bb = beam.get_Geometry(options).GetBoundingBox()
# 		z = bb.Min.Z/3.2808399
# 		if z-114.72<2.79 :
# 			print(z)
# 			print(beam.Id)
# 			icollection.Add(beam.Id)
# 		# "FABRICATION_LEVEL_PARAM"
# 		# param = BuiltInParameter.STRUCTURAL_REFERENCE_LEVEL_ELEVATION
# 		# print(param)
# 		# zlevel = beam.get_Parameter(param).AsDouble()/3.2808399
# 		# print(zlevel)
# uidoc.Selection.SetElementIds(icollection)


# <<<<<<<<<<<<<<<<<<<<<


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