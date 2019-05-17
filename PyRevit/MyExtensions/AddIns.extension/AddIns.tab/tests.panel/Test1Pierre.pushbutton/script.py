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








# # import clr
# # import math
# # clr.AddReference('RevitAPI') 
# # clr.AddReference('RevitAPIUI') 
# # from Autodesk.Revit.DB import *
# # from Autodesk.Revit.UI import *
# # from System.Collections.Generic import List

# # doc = __revit__.ActiveUIDocument.Document
# # uidoc = __revit__.ActiveUIDocument
# # options = __revit__.Application.Create.NewGeometryOptions()
# # SEBoptions = SpatialElementBoundaryOptions()
# # roomcalculator = SpatialElementGeometryCalculator(doc)

# # # def get_selected_elements(doc):
# # #     try:
# # #         # Revit 2016
# # #         return [doc.GetElement(id)
# # #                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
# # #     except:
# # #         # old method
# # #         return list(__revit__.ActiveUIDocument.Selection.Elements)
	
# # # el = get_selected_elements(doc)[0]
# # # print(el.Id)
# # # print(el.GroupId)

# # # t = Transaction(doc, 'Set Param')
# # # t.Start()

# # # el.LookupParameter('ECO_lot').Set("God no")

# # # t.Commit()



# # cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
# # 			BuiltInCategory.OST_CableTray]

# # cat = BuiltInCategory.OST_DuctCurves

# # # element_collector = FilteredElementCollector(doc)\
# # # 	.OfCategory(cat)\
# # # 	.WhereElementIsNotElementType()\
# # # 	.ToElements()

# # element_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# # 	.OfCategory(cat)\
# # 	.WhereElementIsNotElementType()\
# # 	.ToElements()

# # icollection = List[ElementId]()

# # # param = "Niveau de r"+"\xe9"+"f"+"\xe9"+"rence"
# # param = "Type de syst"+"\xe8"+"me"

# # for element in element_collector:
# # 	# print(element.LookupParameter(param).AsValueString()+" !")
# # 	# if element.LookupParameter(param).AsValueString() == "RDC-OLD":
# # 	if element.LookupParameter(param).AsValueString() == "Soufflage":
# # 		icollection.Add(element.Id)
	
# # uidoc.Selection.SetElementIds(icollection)











# # import clr
# # import math
# # clr.AddReference('RevitAPI') 
# # clr.AddReference('RevitAPIUI') 
# # from Autodesk.Revit.DB import *
# # from Autodesk.Revit.UI import *
# # from System.Collections.Generic import List

# # doc = __revit__.ActiveUIDocument.Document
# # uidoc = __revit__.ActiveUIDocument
# # options = __revit__.Application.Create.NewGeometryOptions()
# # SEBoptions = SpatialElementBoundaryOptions()
# # roomcalculator = SpatialElementGeometryCalculator(doc)


# # group_collector = FilteredElementCollector(doc)\
# # 	  .OfClass(Group)

# # def get_selected_elements(doc):
# #     try:
# #         # Revit 2016
# #         return [doc.GetElement(id)
# #                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
# #     except:
# #         # old method
# #         return list(__revit__.ActiveUIDocument.Selection.Elements)

# # def Ungroup(group):
# # 	group.UngroupMembers()

# # for group in group_collector:
# # 	print(group.Id)
# # 	print(group.GroupId)

# # 	groupname = group.Name
# # 	print("Group name : " + groupname)
# # 	groupmember = group.GetMemberIds()

# # 	try:
# # 		t1 = Transaction(doc, 'Ungroup group')
# # 		t1.Start()
# # 		Ungroup(group)
# # 		print("Group ungrouped")
# # 		t1.Commit()
# # 	except:
# # 		t1.RollBack()





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

def get_selected_elements(doc):
    try:
        # # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)

def DocToOriginTransform(doc):
	projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	translationVector = XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation)
	translationTransform = Transform.CreateTranslation(translationVector)
	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition.Angle, XYZ.Zero)
	finalTransform = translationTransform.Multiply(rotationTransform)
	return finalTransform

def OriginToDocTransform(doc):
	projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	translationVector = XYZ(-projectPosition.EastWest, -projectPosition.NorthSouth, -projectPosition.Elevation)
	translationTransform = Transform.CreateTranslation(translationVector)
	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, -projectPosition.Angle, XYZ.Zero)
	finalTransform = rotationTransform.Multiply(translationTransform)
	return finalTransform

def DocToDocTransform(doc1, doc2):
	projectPosition1 = doc1.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	projectPosition2 = doc2.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	translationVector = XYZ(projectPosition1.EastWest-projectPosition2.EastWest, \
		projectPosition1.NorthSouth-projectPosition2.NorthSouth, \
		projectPosition1.Elevation-projectPosition2.Elevation)
	translationTransform = Transform.CreateTranslation(translationVector)
	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition1.Angle-projectPosition2.Angle, XYZ.Zero)
	finalTransform = translationTransform.Multiply(rotationTransform)
	return finalTransform

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

# el1 = get_selected_elements(doc)[0]
# el2 = get_selected_elements(doc)[1]

# print(get_intersection(el1, el2))

el = get_selected_elements(doc)[0]
print(el.Id)

bubbleBb = el.get_Geometry(options).GetBoundingBox()
bubbleTrans = bubbleBb.Transform
bubbleOutline = Outline(bubbleTrans.OfPoint(bubbleBb.Min), bubbleTrans.OfPoint(bubbleBb.Max))
bbFilter = BoundingBoxIntersectsFilter(bubbleOutline)
gm_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
		.OfCategory(BuiltInCategory.OST_GenericModel)\
		.WherePasses(bbFilter)\
		.WhereElementIsNotElementType()\
		.ToElements()

for i in gm_collector:
	print(i)
	print(i.Id)
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

# print(dir(Transform))

# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def GetProjectLocationTransform(doc):
# 	projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
# 	translationVector = XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation)
# 	translationTransform = Transform.CreateTranslation(translationVector)
# 	rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition.Angle, XYZ.Zero)
# 	finalTransform = translationTransform.Multiply(rotationTransform)
# 	return finalTransform

# bpfilter = ElementCategoryFilter(BuiltInCategory.OST_ProjectBasePoint)
# bp = FilteredElementCollector(doc).WherePasses(bpfilter).ToElements()[0]

# xbp = bp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
# ybp = bp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
# zbp = bp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
# print("xbp = " + str(xbp))
# print("ybp = " + str(ybp))
# print("zbp = " + str(zbp))

# plTrans = GetProjectLocationTransform(doc)

# el = get_selected_elements(doc)[0]
# print(el.UniqueId)
# bb = el.get_Geometry(options).GetBoundingBox()
# trans = bb.Transform
# print(bb.Max)
# print(trans.OfPoint(bb.Max))
# print(plTrans.OfPoint(trans.OfPoint(bb.Max)))






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

# class CheckBoxOption:
# 	def __init__(self, name, default_state=False):
# 		self.name = name
# 		self.state = default_state

# 	def __nonzero__(self):
# 		return self.state

# 	def __bool__(self):
# 		return self.state

# view_list = ["La vue active", "Tout le document"]
# res1 = forms.SelectFromList.show(view_list,
# 									multiselect = False,
# 									name_attr = "Vue",
# 									button_name = "OK")

# if res1 != None:
# 	if res1[0] == "La vue active":
# 		collector = FilteredElementCollector(doc, doc.ActiveView.Id)
# 	else:
# 		collector = FilteredElementCollector(doc)

# 	categories = doc.Settings.Categories
# 	cat_dir = {}
# 	for cat in categories:
# 		cat_dir[cat.Name] = cat

# 	# categories = doc.Settings.Categories
# 	# cat_dir = {}
# 	# bic_dir = {}
# 	# for bic in BuiltInCategory.GetValues(BuiltInCategory):
# 	# 	cat = categories.get_Item(bic)
# 	# 	print(cat)
# 	# 	catName = cat.Name
# 	# 	cat_dir[catName] = cat
# 	# 	bic_dir[catName] = bic

# 	# print(cat_dir.values()[0])
# 	# print(bic_dir.values()[0])

# 	cat_list = cat_dir.keys()
# 	options = [CheckBoxOption(j) for j in cat_list]

# 	all_checkboxes = forms.SelectFromCheckBoxes.show(options)

# 	res2 = []
# 	for checkbox in all_checkboxes:
# 		if checkbox.state == True:
# 			res2.append(checkbox.name)

# 	chosenCat_list = [cat_dir[x] for x in res2]

# 	icatcollection = List[ElementId]()
# 	elementFilter = List[ElementFilter](len(chosenCat_list))
# 	for i in chosenCat_list:
# 		icatcollection.Add(i.Id)
# 		elementFilter.Add(ElementCategoryFilter(i.Id))

# 	paramFilter = ParameterFilterUtilities.GetFilterableParametersInCommon(doc,icatcollection)

# 	bip_dir = {}
# 	bipName_dir = {}
# 	for bip in BuiltInParameter.GetValues(BuiltInParameter):
# 		bipId = str(ElementId(bip).IntegerValue)
# 		bip_dir[bipId] = bip
# 		try:
# 			bipName = LabelUtils.GetLabelFor(bip)
# 		except:
# 			bipName = bip.ToString()
# 		bipName_dir[bipId] = bipName

# 	param_dir = {}
# 	for paramId in paramFilter:
# 		if paramId.IntegerValue < 0:
# 			paramName = bipName_dir[str(paramId.IntegerValue)]
# 			param = bip_dir[str(paramId.IntegerValue)]
# 			param_dir[paramName] = param
# 		else:
# 			paramName = doc.GetElement(paramId).Name
# 			param = doc.GetElement(paramId)
# 			param_dir[paramName] = param


# 	res3 = forms.SelectFromList.show(param_dir.keys(),
# 										multiselect = False,
# 										name_attr = "Parametre",
# 										button_name = "OK")

# 	if res3 != None:
# 		value = TextInput('Value', default="Parameter value")
# 		chosenParam = param_dir[res3[0]]
# 		categoryFilter = LogicalOrFilter(elementFilter)
# 		# element_collector = FilteredElementCollector(searchRange)\
# 		# 	.OfCategoryId(cat.Id)\
# 		# 	.WhereElementIsNotElementType()\
# 		# 	.ToElements()
# 		element_collector = collector\
# 			.WhereElementIsNotElementType()\
# 			.WherePasses(categoryFilter)\
# 			.ToElements()

# 		icollection = List[ElementId]()
# 		for e in element_collector:
# 			# print(e)
# 			try:
# 				p = e.get_Parameter(chosenParam.GuidValue)
# 				# print(p)
# 				# print(p.StorageType)
# 				if (str(p.StorageType) == "Integer") and (str(p.AsString()) == str(value)):
# 					icollection.Add(e.Id)
# 				elif (str(p.StorageType) == "String") and (str(p.AsString()) == str(value)):
# 					icollection.Add(e.Id)
# 				elif (str(p.StorageType) == "Double") and (str(p.AsDouble()) == str(value)):
# 					icollection.Add(e.Id)
# 				elif (str(p.StorageType) == "ElementId") and (str(p.AsValueString()) == str(value)):
# 					icollection.Add(e.Id)
# 			except:
# 				p = e.get_Parameter(chosenParam)
# 				# print(p)
# 				# print(p.StorageType)
# 				if (str(p.StorageType) == "Integer") and (str(p.AsString()) == str(value)):
# 					icollection.Add(e.Id)
# 				elif (str(p.StorageType) == "String") and (str(p.AsString()) == str(value)):
# 					icollection.Add(e.Id)
# 				elif (str(p.StorageType) == "Double") and (str(p.AsDouble()) == str(value)):
# 					icollection.Add(e.Id)
# 				elif (str(p.StorageType) == "ElementId") and (str(p.AsValueString()) == str(value)):
# 					icollection.Add(e.Id)

# 		uidoc.Selection.SetElementIds(icollection)
# 	else:
# 		"bye"
# else:
# 	"bye"