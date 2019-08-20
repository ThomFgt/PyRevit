# import clr
# import math
# clr.AddReference('RevitAPI') 
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from System.Collections.Generic import List

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# options = __revit__.Application.Create.NewGeometryOptions()
# SEBoptions = SpatialElementBoundaryOptions()
# roomcalculator = SpatialElementGeometryCalculator(doc)

# class FailureHandler(IFailuresPreprocessor):
#   def __init__(self):
#     self.ErrorMessage = ""
#     self.ErrorSeverity = ""
#   def PreprocessFailures(self, failuresAccessor):
#   	# failuresAccessor.DeleteAllWarning()
#   	# return FailureProcessingResult.Continue
#   	failures = failuresAccessor.GetFailureMessages()
#   	rslt = ""
#   	for f in failures:
#   		fseverity = failuresAccessor.GetSeverity()
#   		if fseverity == FailureSeverity.Warning:
#   			failuresAccessor.DeleteWarning(f)
#   		elif fseverity == FailureSeverity.Error:
#   			rslt = "Error"
#   			failuresAccessor.ResolveFailure(f)
#   	if rslt == "Error":
#   		return FailureProcessingResult.ProceedWithCommit
#   		# return FailureProcessingResult.ProceedWithRollBack
#   	else:
#   		return FailureProcessingResult.Continue

# group_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# 	  .OfClass(Group)

# def get_selected_elements(doc):
#     try:
#         # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def Ungroup(group):
# 	group.UngroupMembers()
	
# def Regroup(groupname,groupmember):
# 	newgroup = doc.Create.NewGroup(groupmember)
# 	newgroup.GroupType.Name = str(groupname)
	
# # group = get_selected_elements(doc)[0]

# # print(group.Id)
# # print(group.GetType())
# # print(group.GroupId)

# # groupname = group.Name
# # groupmember = group.GetMemberIds()

# # IdsInLine = ""
# # for i in groupmember:
# # 	IdsInLine = IdsInLine + str(i.IntegerValue) + ", "

# # IdsInLine = IdsInLine[:len(IdsInLine)-3]

# # print(IdsInLine)

# # status = ""
# # t1 = Transaction(doc, 'Ungroup group')
# # t1.Start()
# # Ungroup(group)
# # print("Group ungrouped")
# # t1.Commit()
# # try:
# # 	t2 = Transaction(doc, 'Regroup group')
# # 	t2.Start()

# # 	failureHandlingOptions = t2.GetFailureHandlingOptions()
# # 	failureHandler = FailureHandler()
# # 	failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
# # 	failureHandlingOptions.SetClearAfterRollback(True)
# # 	t2.SetFailureHandlingOptions(failureHandlingOptions)

# # 	Regroup(groupname,groupmember)
# # 	print("Group regrouped")
# # 	status = "Yeah!"
# # 	t2.Commit()

# # except:
# # 	t2.RollBack()
# # 	print("Regrouping fail")
# # 	status = "Fuck!"

# # print(status + "\n")
# # print(t2.GetStatus())





# tg = TransactionGroup(doc, 'Ungroup/regroup all groups')
# tg.Start()

# for group in group_collector:
# 	print(group.Id)
# 	print(group.GroupId)

# 	groupname = group.Name
# 	print("Group name : " + groupname)
# 	groupmember = group.GetMemberIds()

# 	t1 = Transaction(doc, 'Ungroup group')
# 	t1.Start()
# 	Ungroup(group)
# 	print("Group ungrouped")
# 	t1.Commit()
# 	try:
# 		t2 = Transaction(doc, 'Regroup group')
# 		t2.Start()

# 		failureHandlingOptions = t2.GetFailureHandlingOptions()
# 		failureHandler = FailureHandler()
# 		failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
# 		failureHandlingOptions.SetClearAfterRollback(True)
# 		t2.SetFailureHandlingOptions(failureHandlingOptions)

# 		Regroup(groupname,groupmember)
# 		print("Group regrouped")
# 		t2.Commit()

# 	except:
# 		t2.RollBack()
# 		print("Regrouping fail")
# 		IdsInLine = ""
# 		for i in groupmember:
# 			IdsInLine = IdsInLine + str(i.IntegerValue) + ", "

# 		IdsInLine = IdsInLine[:len(IdsInLine)-3]

# 		print("Grouped element ids was : " + IdsInLine)

# print("Done")

# tg.Assimilate()







# import clr
# import math
# clr.AddReference('RevitAPI') 
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from System.Collections.Generic import List
# from pyrevit import forms
# from rpw.ui.forms import TextInput

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# options = __revit__.Application.Create.NewGeometryOptions()

# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def DocToOriginTransform(doc):
# 	projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
# 	translationVector = XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation)
# 	translationTransform = Transform.CreateTranslation(translationVector)
# 	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition.Angle, XYZ.Zero)
# 	finalTransform = translationTransform.Multiply(rotationTransform)
# 	return finalTransform

# def OriginToDocTransform(doc):
# 	projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
# 	translationVector = XYZ(-projectPosition.EastWest, -projectPosition.NorthSouth, -projectPosition.Elevation)
# 	translationTransform = Transform.CreateTranslation(translationVector)
# 	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, -projectPosition.Angle, XYZ.Zero)
# 	finalTransform = rotationTransform.Multiply(translationTransform)
# 	return finalTransform

# def DocToDocTransform(doc1, doc2):
# 	projectPosition1 = doc1.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
# 	projectPosition2 = doc2.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
# 	translationVector = XYZ(projectPosition1.EastWest-projectPosition2.EastWest, \
# 		projectPosition1.NorthSouth-projectPosition2.NorthSouth, \
# 		projectPosition1.Elevation-projectPosition2.Elevation)
# 	translationTransform = Transform.CreateTranslation(translationVector)
# 	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition1.Angle-projectPosition2.Angle, XYZ.Zero)
# 	finalTransform = translationTransform.Multiply(rotationTransform)
# 	return finalTransform

# def get_solid(element):
# 	solid_list = []
# 	for i in element.get_Geometry(options):
# 		if i.ToString() == "Autodesk.Revit.DB.Solid":
# 			solid_list.append(SolidUtils.CreateTransformed(i, DocToDocTransform(element.Document, __revit__.ActiveUIDocument.Document)))
# 		elif i.ToString() == "Autodesk.Revit.DB.GeometryInstance":
# 			for j in i.GetInstanceGeometry():
# 				if j.ToString() == "Autodesk.Revit.DB.Solid":
# 					solid_list.append(SolidUtils.CreateTransformed(j, DocToDocTransform(element.Document, __revit__.ActiveUIDocument.Document)))
# 	return solid_list


# def get_docsolid(element, elementdoc, activedoc):
# 	docToOriginTrans = DocToOriginTransform(elementdoc)
# 	originToDocTrans = OriginToDocTransform(activedoc)

# 	bb = element.get_Geometry(options).GetBoundingBox()
# 	trans = bb.Transform
# 	bbmin = originToDocTrans.OfPoint(docToOriginTrans.OfPoint(trans.OfPoint(bb.Min)))
# 	bbmax = originToDocTrans.OfPoint(docToOriginTrans.OfPoint(trans.OfPoint(bb.Max)))

# 	pt0 = XYZ(bbmin.X, bbmin.Y, bbmin.Z)
# 	pt1 = XYZ(bbmax.X, bbmin.Y, bbmin.Z)
# 	pt2 = XYZ(bbmax.X, bbmax.Y, bbmin.Z)
# 	pt3 = XYZ(bbmin.X, bbmax.Y, bbmin.Z)

# 	edge0 = Line.CreateBound(pt0, pt1)
# 	edge1 = Line.CreateBound(pt1, pt2)
# 	edge2 = Line.CreateBound(pt2, pt3)
# 	edge3 = Line.CreateBound(pt3, pt0)

# 	iCurveCollection = List[Curve]()
# 	iCurveCollection.Add(edge0)
# 	iCurveCollection.Add(edge1)
# 	iCurveCollection.Add(edge2)
# 	iCurveCollection.Add(edge3)
# 	height = bbmax.Z - bbmin.Z
# 	baseLoop = CurveLoop.Create(iCurveCollection)

# 	iCurveLoopCollection = List[CurveLoop]()
# 	iCurveLoopCollection.Add(baseLoop)

# 	solid = GeometryCreationUtilities.CreateExtrusionGeometry(iCurveLoopCollection, XYZ.BasisZ, height)
	
# 	return solid

# def get_intersection(el1, el2):
# 	bb1 = el1.get_Geometry(options).GetBoundingBox()
# 	bb2 = el2.get_Geometry(options).GetBoundingBox()

# 	trans1 = bb1.Transform
# 	trans2 = bb2.Transform

# 	min1 = trans1.OfPoint(bb1.Min)
# 	max1 = trans1.OfPoint(bb1.Max)
# 	min2 = trans2.OfPoint(bb2.Min)
# 	max2 = trans2.OfPoint(bb2.Max)

# 	outline1 = Outline(min1, max1)
# 	outline2 = Outline(min2, max2)

# 	solid1_list = get_solid(el1)
# 	solid2_list = get_solid(el2)

# 	for i in solid1_list:
# 		for j in solid2_list:
# 			try:
# 				inter = BooleanOperationsUtils.ExecuteBooleanOperation(i, j, BooleanOperationsType.Intersect)
# 				if inter.Volume != 0:
# 					interBb = inter.GetBoundingBox()
# 					interTrans = interBb.Transform
# 					interPoint = interTrans.OfPoint(interBb.Min)
# 					break
# 			except:
# 				"Oh god!"

# 	try:
# 		interPoint
# 		return interPoint
# 	except:
# 		return None

# el = get_selected_elements(doc)[0]
# print(el.Id)

# bubbleBb = el.get_Geometry(options).GetBoundingBox()
# bubbleTrans = bubbleBb.Transform
# bubbleOutline = Outline(bubbleTrans.OfPoint(bubbleBb.Min), bubbleTrans.OfPoint(bubbleBb.Max))
# bbFilter = BoundingBoxIntersectsFilter(bubbleOutline)
# gm_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 		.OfCategory(BuiltInCategory.OST_GenericModel)\
# 		.WherePasses(bbFilter)\
# 		.WhereElementIsNotElementType()\
# 		.ToElements()

# for i in gm_collector:
# 	print(i)
# 	print(i.Id)







# # # # # ATTENTION MODULE FORM DANS REVIT
import clr
import System
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
from System.Drawing import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
# app = __revit__.Application
app = uidoc.Application.Application

# # if doc.ActiveView.GetType().ToString() == "Autodesk.Revit.DB.ViewSheet":
# # 	viewId_list = doc.ActiveView.GetAllPlacedViews()
# # 	for viewId in viewId_list:
# # 		print(viewId)
# # 		view = doc.GetElement(viewId)
# # 		print(view)
# # 		for i in view.GetExternalFileReference():
# # 			print(i)
# # else:
# # 	print("The active view is not a view sheet!")

# dwgSettingsFilter = ElementClassFilter(ExportDWGSettings)
# settings = FilteredElementCollector(doc)
# settings = settings.WherePasses(dwgSettingsFilter)

# xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
# # worksheet = xlApp.Worksheets[1]
# worksheet = xlApp.ActiveSheet

# i = 0
# for setting in settings:
# 	# if setting.Name == "Export NTW_TESTbis":
# 	if setting.Name == "Export CITY":
# 		dwgOptions = setting.GetDWGExportOptions()
# 		layerTable = dwgOptions.GetExportLayerTable()
# 		for layerItem in layerTable:
# 			cat = layerItem.Key
# 			layerInfo = layerItem.Value
# 			categoryType = layerInfo.CategoryType
# 			catName = cat.CategoryName
# 			subCatName = cat.SubCategoryName
# 			layerName = layerInfo.LayerName
# 			cutLayerName = layerInfo.CutLayerName
# 			layerModifiers = layerInfo.GetLayerModifiers()
# 			layerModifier = ""
# 			cutLayerModifiers = layerInfo.GetCutLayerModifiers()
# 			cutLayerModifier = ""

# 			if (".dwg" not in catName) and ((categoryType.ToString() == "Model") or (categoryType.ToString() == "Annotation")) :
# 				i = i + 1
# 				print(dir(layerItem))
# 				break
# 				print("--")

# # k = 0
# # l = 0
# # for cat in doc.Settings.Categories:
# # 	catName = cat.Name
# # 	categoryType = cat.CategoryType
# # 	if (".dwg" not in catName) and ((categoryType.ToString() == "Model") or (categoryType.ToString() == "Annotation")) :
# # 		print(cat.SubCategories.Size)
# # 		k = k + 1
# # 		for j in cat.SubCategories:
# # 			l = l + 1

# # print(k+l)



# # CHERCHER ZONE REMPLIES!

# # OST_FillPatterns

# # collector = FilteredElementCollector(doc)\
# # 		.OfClass(FillPatternElement)\
# # 		.WhereElementIsNotElementType()

# collector = FilteredElementCollector(doc).OfClass(FilledRegionType)

# # collector = FilteredElementCollector(doc)\
# # 		.OfClass(FillPatternElement)

# # collector = FilteredElementCollector(doc)\
# # 		.OfCategory(BuiltInCategory.OST_FillPatterns)



# for i in collector:
# 	print(i.Name)

# # CHERCHER ZONE REMPLIES!








# # # # # USE REVIT FORMS
# a = TextBox
# a.Name = "TB"
# a.Visible = True
# a.Enabled = True
# # # # # USE REVIT FORMS









# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# e = get_selected_elements(doc)
# for i in e:
# 	print(i.LookupParameter("Ai NVP").AsDouble())




# insulationId = InsulationLiningBase.GetInsulationIds(doc, e.Id)[0]
# e1 = doc.GetElement(insulationId)
# print(e1.LookupParameter("Sous-projet").AsValueString())

# wsparam = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
# wsparamId = wsparam.Id


# provider = ParameterValueProvider(wsparamId)
# evaluator = FilterStringEquals()
# wsparamValue = "13_CVP_CVC_Aeraulique"
# rule = FilterStringRule(provider, evaluator, wsparamValue, False)
# wsfilter = ElementParameterFilter(rule)

# collector = FilteredElementCollector(doc, doc.ActiveView.Id).WherePasses(wsfilter).ToElementIds()

# uidoc.Selection.SetElementIds(collector)

# print(wsparam.StorageType)
# print(wsparam.AsInteger())
# print(wsparam.AsValueString())








# collector = FilteredElementCollector(doc).OfClass(ParameterValue)
# for i in collector:
# 	print(i)
# 	print(i.GetType())
# 	print(i.ToString())
# 	break


# # # # # NUMEROTE BULLES
# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)


# e = get_selected_elements(doc)[0]
# level = e.LookupParameter("Niveau").AsValueString()
# name = e.Name

# collector = FilteredElementCollector(doc)\
# 				.OfCategory(BuiltInCategory.OST_GenericModel)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# num_list = []
# j = 0
# for i in collector:
# 	level_i = i.LookupParameter("Niveau").AsValueString()
# 	name_i = i.Name
# 	if (level_i == level) and (name_i == name):
# 		j = j + 1
# 		num_i = i.LookupParameter("BPS - Num"+str("\xe9")+"ro").AsString()
# 		num_list.append(int(num_i))
# 		if num_i != str(j):
# 			print(j)
# 			t = Transaction(doc, 'Number bubble')
# 			t.Start()
# 			# i.LookupParameter("BPS - Num"+str("\xe9")+"ro").Set(str(j))
# 			e.LookupParameter("BPS - Num"+str("\xe9")+"ro").Set(str(j))
# 			t.Commit()
# 			break

# print(num_list)
# num_sorted_list = sorted(num_list)
# print(num_sorted_list)
# # # # # NUMEROTE BULLES





# # # # # ZOOM SUR ELEMENT SELECTIONNE
def get_selected_elements(doc):
    try:
        # # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)


e = get_selected_elements(doc)[0]

# # # # # ZOOM SUR ELEMENT SELECTIONNE





# # # # # COLORISE UN RESEAU SOUS UNE CERTAINE AI
# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def ColorElement(e):
# 	bb = e.get_Geometry(options).GetBoundingBox()
# 	trans = bb.Transform
# 	minPoint = trans.OfPoint(bb.Min)
# 	zMin = minPoint.Z

# 	i = 0
# 	for el in elevation_list:
# 		i = i + 1
# 		if zMin >= el:
# 			pass
# 		else:
# 			Ai = zMin - elevation_list[i-2]
# 			break

# 	try:
# 		Aim = str(round(UnitUtils.ConvertFromInternalUnits(Ai, DisplayUnitType.DUT_METERS),3))

# 		if float(Aim) < 2.05:
# 			ogs = OverrideGraphicSettings()
# 			color = Color(255, 0, 0)
# 			ogs.SetProjectionFillColor(color)
# 			ogs.SetCutFillColor(color)
# 			fillPatternElement = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "Uni")
# 			ogs.SetProjectionFillPatternId(fillPatternElement.Id)
# 			ogs.SetCutFillPatternId(fillPatternElement.Id)
# 			ogs.SetSurfaceTransparency(50)
# 			ogs.SetProjectionFillPatternVisible(True)
# 			ogs.SetCutFillPatternVisible(True)
# 			# t = Transaction(doc, 'Color element')
# 			# t.Start()
# 			doc.ActiveView.SetElementOverrides(e.Id, ogs)
# 			# t.Commit()
# 		else:
# 			"Ok"
# 	except:
# 		print(e)
# 		print("Didn't end well")

# # e = get_selected_elements(doc)[0]

# level_collector = FilteredElementCollector(doc)\
# 				.OfCategory(BuiltInCategory.OST_Levels)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# level_list = {}
# for l in level_collector:
# 	level_list[l.Name] = l.Elevation

# elevation_list = sorted(level_list.values())


# icatcollection = List[BuiltInCategory]()
# icatcollection.Add(BuiltInCategory.OST_PipeCurves)
# icatcollection.Add(BuiltInCategory.OST_DuctCurves)
# icatcollection.Add(BuiltInCategory.OST_CableTray)
# icatcollection.Add(BuiltInCategory.OST_PipeInsulations)
# icatcollection.Add(BuiltInCategory.OST_DuctInsulations)

# multiCatFilter = ElementMulticategoryFilter(icatcollection)

# collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 				.WherePasses(multiCatFilter)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# t = Transaction(doc, 'Color elements')
# t.Start()
# for e in collector:
# 	ColorElement(e)
# t.Commit()
# # # # # COLORISE UN RESEAU SOUS UNE CERTAINE AI








# import_collector = FilteredElementCollector(doc)\
# 	  .OfClass(ImportInstance)\
# 	  .WhereElementIsNotElementType()\

# for i in import_collector:
# 	print(i.Category.Name)
# 	view = doc.GetElement(i.OwnerViewId).Name
# 	print(view)
# 	print("----------")




# # # # # Check RSO

# def get_solid(element):
# 	solid_list = []
# 	for i in element.get_Geometry(options):
# 		if i.ToString() == "Autodesk.Revit.DB.Solid":
# 			solid_list.append(i)
# 		elif i.ToString() == "Autodesk.Revit.DB.GeometryInstance":
# 			for j in i.GetInstanceGeometry():
# 				if j.ToString() == "Autodesk.Revit.DB.Solid":
# 					solid_list.append(j)
# 	return solid_list

# def get_intersection(el1, el2):
# 	bb1 = el1.get_Geometry(options).GetBoundingBox()
# 	bb2 = el2.get_Geometry(options).GetBoundingBox()

# 	trans1 = bb1.Transform
# 	trans2 = bb2.Transform

# 	min1 = trans1.OfPoint(bb1.Min)
# 	max1 = trans1.OfPoint(bb1.Max)
# 	min2 = trans2.OfPoint(bb2.Min)
# 	max2 = trans2.OfPoint(bb2.Max)

# 	outline1 = Outline(min1, max1)
# 	outline2 = Outline(min2, max2)

# 	solid1_list = get_solid(el1)
# 	solid2_list = get_solid(el2)

# 	for i in solid1_list:
# 		for j in solid2_list:
# 			try:
# 				inter = BooleanOperationsUtils.ExecuteBooleanOperation(i, j, BooleanOperationsType.Intersect)
# 				if inter.Volume != 0:
# 					interBb = inter.GetBoundingBox()
# 					interTrans = interBb.Transform
# 					interPoint = interTrans.OfPoint(interBb.Min)
# 					break
# 			except:
# 				"Oh god!"

# 	try:
# 		interPoint
# 		return interPoint
# 	except:
# 		return None

# def ColorElement(e):
# 	ogs = OverrideGraphicSettings()
# 	color = Color(255, 0, 0)
# 	ogs.SetProjectionFillColor(color)
# 	ogs.SetCutFillColor(color)
# 	try:
# 		fillPatternElement = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "Uni")
# 		fillPatternElement.GetFillPattern().IsSolidFill
# 	except:
# 		fillPatternElement = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "<Remplissage de solide>")
# 	ogs.SetProjectionFillPatternId(fillPatternElement.Id)
# 	ogs.SetCutFillPatternId(fillPatternElement.Id)
# 	ogs.SetSurfaceTransparency(50)
# 	ogs.SetProjectionFillPatternVisible(True)
# 	ogs.SetCutFillPatternVisible(True)
# 	# t = Transaction(doc, 'Color element')
# 	# t.Start()
# 	doc.ActiveView.SetElementOverrides(e.Id, ogs)
# 	# t.Commit()



# cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
# 			BuiltInCategory.OST_CableTray, BuiltInCategory.OST_FlexPipeCurves]

# icatcollection = List[BuiltInCategory]()
# for cat in cat_list:
# 	icatcollection.Add(cat)

# multiCatFilter = ElementMulticategoryFilter(icatcollection)

# collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 				.WherePasses(multiCatFilter)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# t = Transaction(doc, 'Check RSO')
# t.Start()
# for e1 in collector:
# 	try:
# 		insulationId = InsulationLiningBase.GetInsulationIds(doc, e1.Id)[0]
# 		e1 = doc.GetElement(insulationId)
# 	except:
# 		"e1 not insulated"

# 	bb = e1.get_Geometry(options).GetBoundingBox()
# 	trans = bb.Transform
# 	outline = Outline(trans.OfPoint(bb.Min), trans.OfPoint(bb.Max))
# 	bbFilter = BoundingBoxIntersectsFilter(outline)

# 	totalFilter = LogicalAndFilter(multiCatFilter, bbFilter)

# 	collector2 = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 					.WherePasses(totalFilter)\
# 					.WhereElementIsNotElementType()\
# 					.ToElements()

# 	for e2 in collector2:
# 		try:
# 			insulationId2 = InsulationLiningBase.GetInsulationIds(doc, e2.Id)[0]
# 			e2 = doc.GetElement(insulationId2)
# 		except:
# 			"e2 not insulated"
# 		if e1.Id.IntegerValue != e2.Id.IntegerValue:
# 			if get_intersection(e1, e2) is not None:
# 				ColorElement(e1)
# 				ColorElement(e2)

# t.Commit()

# # # # # Check RSO






# # # # # Check BEAM

# def get_solid(element):
# 	solid_list = []
# 	for i in element.get_Geometry(options):
# 		if i.ToString() == "Autodesk.Revit.DB.Solid":
# 			solid_list.append(i)
# 		elif i.ToString() == "Autodesk.Revit.DB.GeometryInstance":
# 			for j in i.GetInstanceGeometry():
# 				if j.ToString() == "Autodesk.Revit.DB.Solid":
# 					solid_list.append(j)
# 	return solid_list

# def get_intersection(el1, el2):
# 	bb1 = el1.get_Geometry(options).GetBoundingBox()
# 	bb2 = el2.get_Geometry(options).GetBoundingBox()

# 	trans1 = bb1.Transform
# 	trans2 = bb2.Transform

# 	min1 = trans1.OfPoint(bb1.Min)
# 	max1 = trans1.OfPoint(bb1.Max)
# 	min2 = trans2.OfPoint(bb2.Min)
# 	max2 = trans2.OfPoint(bb2.Max)

# 	outline1 = Outline(min1, max1)
# 	outline2 = Outline(min2, max2)

# 	solid1_list = get_solid(el1)
# 	solid2_list = get_solid(el2)

# 	for i in solid1_list:
# 		for j in solid2_list:
# 			try:
# 				inter = BooleanOperationsUtils.ExecuteBooleanOperation(i, j, BooleanOperationsType.Intersect)
# 				if inter.Volume != 0:
# 					interBb = inter.GetBoundingBox()
# 					interTrans = interBb.Transform
# 					interPoint = interTrans.OfPoint(interBb.Min)
# 					break
# 			except:
# 				"Oh god!"

# 	try:
# 		interPoint
# 		return interPoint
# 	except:
# 		return None

# def ColorElement(e):
# 	ogs = OverrideGraphicSettings()
# 	color = Color(255, 0, 0)
# 	ogs.SetProjectionFillColor(color)
# 	ogs.SetCutFillColor(color)
# 	try:
# 		fillPatternElement = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "Uni")
# 		fillPatternElement.GetFillPattern().IsSolidFill
# 	except:
# 		fillPatternElement = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, "<Remplissage de solide>")
# 	ogs.SetProjectionFillPatternId(fillPatternElement.Id)
# 	ogs.SetCutFillPatternId(fillPatternElement.Id)
# 	ogs.SetSurfaceTransparency(50)
# 	ogs.SetProjectionFillPatternVisible(True)
# 	ogs.SetCutFillPatternVisible(True)
# 	# t = Transaction(doc, 'Color element')
# 	# t.Start()
# 	doc.ActiveView.SetElementOverrides(e.Id, ogs)
# 	# t.Commit()

# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# STRcollector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 				.OfCategory(BuiltInCategory.OST_StructuralFraming)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# t = Transaction(doc, 'Check BEAM')
# t.Start()
# for e in STRcollector:
# 	# e = get_selected_elements(doc)[0]

# 	bb = e.get_Geometry(options).GetBoundingBox()
# 	trans = bb.Transform
# 	outline = Outline(trans.OfPoint(bb.Min), trans.OfPoint(bb.Max))
# 	bbFilter = BoundingBoxIntersectsFilter(outline)

# 	cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
# 				BuiltInCategory.OST_CableTray, BuiltInCategory.OST_FlexPipeCurves,\
# 				BuiltInCategory.OST_PipeInsulations, BuiltInCategory.OST_DuctInsulations]
# 	icatcollection = List[BuiltInCategory]()
# 	for cat in cat_list:
# 		icatcollection.Add(cat)
# 	multiCatFilter = ElementMulticategoryFilter(icatcollection)

# 	totalFilter = LogicalAndFilter(multiCatFilter, bbFilter)

# 	collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 					.WherePasses(totalFilter)\
# 					.WhereElementIsNotElementType()\
# 					.ToElements()

# 	for i in collector:
# 		if get_intersection(e, i) is not None:
# 			ColorElement(i)
# t.Commit()

# # # # # Check BEAM









# # # # # Check connection

# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def CheckConnection(e1, e2):
# 	if (e1.ToString() == "Autodesk.Revit.DB.FamilyInstance") or (e2.ToString() == "Autodesk.Revit.DB.FamilyInstance"):
# 	# 	try:
# 	# 		connector = e1
# 	# 		connectorList = e2.ConnectorManager.Connectors
# 	# 	except:
# 	# 		connector = e2
# 	# 		connectorList = e1.ConnectorManager.Connectors
# 	# 	print(dir(connector.MEPModel.ConnectorManager.Connectors))
# 	# 	print(connectorList)
# 	# 	for i in connectorList:
# 	# 		print(i)
# 	# 		if i.Id == connector.Id:
# 	# 			return True
# 	# 			break
# 	# 	else :
# 	# 		return False
# 	# else:
# 	# 	return False
# 		try:
# 			connectorList1 = e1.MEPModel.ConnectorManager.Connectors
# 			connectorList2 = e2.ConnectorManager.Connectors
# 		except:
# 			connectorList2 = e2.MEPModel.ConnectorManager.Connectors
# 			connectorList1 = e1.ConnectorManager.Connectors
# 		for c1 in connectorList1:
# 			print(e1)
# 			print(c1)
# 			if c1 in connectorList2:
# 				return True
# 				break
# 			else:
# 				return False
# 		for c2 in connectorList2:
# 			print(e2)
# 			print(c2)
# 			if c2 in connectorList1:
# 				return True
# 				break
# 			else:
# 				pass

# e1 = get_selected_elements(doc)[0]
# e2 = get_selected_elements(doc)[1]

# print(CheckConnection(e1, e2))

# # # # # Check connection








# el = get_selected_elements(doc)[0]
# print(el)
# print(el.Category.Name)
# print(el.GetType())

# collector = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_FilledRegion)

# for i in collector:
# 	print(i.FamilyName)



# # # TRY TO GET ELEMENT FROM LINK MODEL
# sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element, "Please select an object")

# el = get_selected_elements(doc)[0]
# print(el)
# # # TRY TO GET ELEMENT FROM LINK MODEL




# # # TRY TO REFRESH BUBBLES IN VIEW LOOK THE BUILDING CODER
# t = Transaction(doc, 'tg')
# t.Start()
# doc.Regenerate()
# t.Commit()
# uidoc.RefreshActiveView()
# # # TRY TO REFRESH BUBBLES IN VIEW LOOK THE BUILDING CODER




# # # GET DEFAULT VALUE OF FAMILY SYMBOL BEFORE PLACING IT
# el = get_selected_elements(doc)[0]

# gm_collector = FilteredElementCollector(doc)\
# 	.OfClass(FamilySymbol)\
# 	.OfCategory(BuiltInCategory.OST_GenericModel)

# for i in gm_collector:
# 	print(i.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM))
# 	print(i.LevelId)

# print(el.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM).AsString())
# # # GET DEFAULT VALUE OF FAMILY SYMBOL BEFORE PLACING IT