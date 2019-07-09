# # # # # Clash detection 2 # # # # #

__title__ = 'Place bubbles\nin active 3D view'

__doc__ = 'Ce programme place les bulles dans la vue 3D active'

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

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# options = __revit__.Application.Create.NewGeometryOptions()
# copyOptions = CopyPasteOptions()

# class Bubble:
	# doc = __revit__.ActiveUIDocument.Document
	# uidoc = __revit__.ActiveUIDocument
	# options = __revit__.Application.Create.NewGeometryOptions()

	# def __init__(self,bubble_name):
		# self.familysymbol = None		
		# self.presence = False
		
		# gm_collector = FilteredElementCollector(doc)\
			  # .OfClass(FamilySymbol)\
			  # .OfCategory(BuiltInCategory.OST_GenericModel)
				
		# for i in gm_collector:
			# if i.FamilyName == str(bubble_name):
				# self.familysymbol = i
				# self.presence = True
				# break
				
		# return self.familysymbol
		# return self.presence

# def DocToOriginTransform(doc):
	# projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	# translationVector = XYZ(projectPosition.EastWest, projectPosition.NorthSouth, projectPosition.Elevation)
	# translationTransform = Transform.CreateTranslation(translationVector)
	# rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPosition.Angle, XYZ.Zero)
	# finalTransform = translationTransform.Multiply(rotationTransform)
	# return finalTransform

# def OriginToDocTransform(doc):
	# projectPosition = doc.ActiveProjectLocation.get_ProjectPosition(XYZ.Zero)
	# translationVector = XYZ(-projectPosition.EastWest, -projectPosition.NorthSouth, -projectPosition.Elevation)
	# translationTransform = Transform.CreateTranslation(translationVector)
	# rotationTransform = Transform.CreateRotationAtPoint(XYZ.BasisZ, -projectPosition.Angle, XYZ.Zero)
	# finalTransform = rotationTransform.Multiply(translationTransform)
	# return finalTransform

# def interpoint(el1, el2, doc1, doc2):

	# doc = __revit__.ActiveUIDocument.Document
	# uidoc = __revit__.ActiveUIDocument
	# options = __revit__.Application.Create.NewGeometryOptions()
	# intersectOptions = SolidCurveIntersectionOptions()

	# if (el1.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and ((el2.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance")):
		# return None
	# else:
		# x = 0
		# if doc1.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
			# x = x + 1
			# trans1 = doc1.GetTotalTransform()
			# bb1 = el1.get_Geometry(options).GetBoundingBox()
			# outline1 = Outline(trans1.OfPoint(bb1.Min),trans1.OfPoint(bb1.Max))
		# else:
			# bb1 = el1.get_Geometry(options).GetBoundingBox()
			# outline1 = Outline(bb1.Min,bb1.Max)
			
		# if doc2.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
			# x = x + 1
			# trans2 = doc2.GetTotalTransform()
			# bb2 = el2.get_Geometry(options).GetBoundingBox()
			# outline2 = Outline(trans2.OfPoint(bb2.Min),trans2.OfPoint(bb2.Max))
		# else:
			# bb2 = el2.get_Geometry(options).GetBoundingBox()
			# outline2 = Outline(bb2.Min,bb2.Max)
	
		# filter = BoundingBoxIntersectsFilter(outline1)
	
		# result = outline1.Intersects(outline2, 0)
		
		# if result is True:
			# el1_solids_list = []
			# el1_curves_list = []
			# for i in el1.get_Geometry(options):
				# if i.GetType().ToString() == "Autodesk.Revit.DB.Solid":
					# el1_solids_list.append(i)
					# for j in i.Edges:
						# el1_curves_list.append(j.AsCurve())	
		
			# try:
				# el1_curves_list.append(el1.Location.Curve)
			# except:
				# "Oh God"
				
			# try:
				# if "Autodesk.Revit.DB.Arc" in el1_curves_list[0].GetType().ToString():
					# try:
						# c1 = Line.CreateBound(trans1.OfPoint(el1_curves_list[0].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[6].GetEndPoint(0)))
						# c2 = Line.CreateBound(trans1.OfPoint(el1_curves_list[1].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[7].GetEndPoint(0)))
						# c3 = Line.CreateBound(trans1.OfPoint(el1_curves_list[2].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[4].GetEndPoint(0)))
						# c4 = Line.CreateBound(trans1.OfPoint(el1_curves_list[3].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[5].GetEndPoint(0)))
					# except:
						# c1 = Line.CreateBound(el1_curves_list[0].GetEndPoint(0),el1_curves_list[6].GetEndPoint(0))
						# c2 = Line.CreateBound(el1_curves_list[1].GetEndPoint(0),el1_curves_list[7].GetEndPoint(0))
						# c3 = Line.CreateBound(el1_curves_list[2].GetEndPoint(0),el1_curves_list[4].GetEndPoint(0))
						# c4 = Line.CreateBound(el1_curves_list[3].GetEndPoint(0),el1_curves_list[5].GetEndPoint(0))
					# el1_curves_list.extend([c1, c2, c3, c4])
			# except:
				# "Oh God"
		
			
			# el2_solids_list = []
			# el2_curves_list = []
			# for i in el2.get_Geometry(options):
				# if i.GetType().ToString() == "Autodesk.Revit.DB.Solid":
					# el2_solids_list.append(i)
					# for j in i.Edges:
						# el2_curves_list.append(j.AsCurve())	
		
			# try:
				# el2_curves_list.append(el2.Location.Curve)
			# except:
				# "Oh God"
				
			# try:
				# if "Autodesk.Revit.DB.Arc" in el2_curves_list[0].GetType().ToString():
					# try:
						# c1 = Line.CreateBound(trans2.OfPoint(el2_curves_list[0].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[6].GetEndPoint(0)))
						# c2 = Line.CreateBound(trans2.OfPoint(el2_curves_list[1].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[7].GetEndPoint(0)))
						# c3 = Line.CreateBound(trans2.OfPoint(el2_curves_list[2].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[4].GetEndPoint(0)))
						# c4 = Line.CreateBound(trans2.OfPoint(el2_curves_list[3].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[5].GetEndPoint(0)))
					# except:
						# c1 = Line.CreateBound(el2_curves_list[0].GetEndPoint(0),el2_curves_list[6].GetEndPoint(0))
						# c2 = Line.CreateBound(el2_curves_list[1].GetEndPoint(0),el2_curves_list[7].GetEndPoint(0))
						# c3 = Line.CreateBound(el2_curves_list[2].GetEndPoint(0),el2_curves_list[4].GetEndPoint(0))
						# c4 = Line.CreateBound(el2_curves_list[3].GetEndPoint(0),el2_curves_list[5].GetEndPoint(0))
					# el2_curves_list.extend([c1, c2, c3, c4])
			# except:
				# "Oh God"
			
			# if len(el1_solids_list)>len(el2_solids_list):
				# solids_list = el1_solids_list
				# curves_list = el2_curves_list
			# else:
				# solids_list = el2_solids_list
				# curves_list = el1_curves_list
			
			# for k in solids_list:
				# for l in curves_list:
					# try:
						# solidcurveintersection = k.IntersectWithCurve(l,intersectOptions)
						# seg_nb = solidcurveintersection.SegmentCount
						# if seg_nb != 0:
							# if x == 0:
								# point = solidcurveintersection.GetCurveSegment(0).GetEndPoint(0)
								# break
							# else:
								# try:
									# point = trans1.OfPoint(solidcurveintersection.GetCurveSegment(0).GetEndPoint(0))
								# except:
									# point = trans2.OfPoint(solidcurveintersection.GetCurveSegment(0).GetEndPoint(0))
								# break
					# except:
						# pass
			
			# try:
				# return point
			# except:
				# return None
		# else:
			# return None

# bname = "OBSERVATIONS_2017"

# if Bubble(bname).presence is True:
	# cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
				# BuiltInCategory.OST_CableTray, BuiltInCategory.OST_StructuralFraming,\
				# BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctFitting,\
				# BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_FlexPipeCurves]
	
	# rvtlink_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
			# .OfCategory(BuiltInCategory.OST_RvtLinks)\
			# .WhereElementIsNotElementType()\
			# .ToElements()

	# SectionBox = doc.ActiveView.GetSectionBox()
	# trans = SectionBox.Transform
	# bbmin = SectionBox.Min
	# bbmax = SectionBox.Max
	
	# dl_list = []
	# for rvtlink in rvtlink_collector:
		# dl_list.append(rvtlink)
	# dl_list.append(doc)

	# element_list = []
	# for cat in cat_list:
		# for dl in dl_list:
			# if dl.GetType().ToString() == "Autodesk.Revit.DB.Document":
				# outline = Outline(trans.OfPoint(bbmin),trans.OfPoint(bbmax))
				# bbFilter = BoundingBoxIntersectsFilter(outline)
				# element_collector = FilteredElementCollector(dl)\
					# .OfCategory(cat)\
					# .WherePasses(bbFilter)\
					# .WhereElementIsNotElementType()\
					# .ToElements()
				# for e in element_collector:
					# if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
						# pass
					# else:
						# element_list.append(e)
			# elif dl.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
				# docToOriginTrans = DocToOriginTransform(doc)
				# originToDocTrans = OriginToDocTransform(dl.GetLinkDocument())
				# docToDocTrans = originToDocTrans.Multiply(docToOriginTrans)
				# try:
					# outline = Outline(docToDocTrans.OfPoint(trans.OfPoint(bbmin)), docToDocTrans.OfPoint(trans.OfPoint(bbmax)))
					# bbFilter = BoundingBoxIntersectsFilter(outline)
				# except:
					# a = "Same base point"
					# outline = Outline(trans.OfPoint(bbmin), trans.OfPoint(bbmax))
					# bbFilter = BoundingBoxIntersectsFilter(outline)
				# element_collector = FilteredElementCollector(dl.GetLinkDocument())\
					# .OfCategory(cat)\
					# .WherePasses(bbFilter)\
					# .WhereElementIsNotElementType()\
					# .ToElements()
				# for e in element_collector:
					# if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
						# pass
					# else:
						# element_list.append(e)

	# print(element_list)
	
	# tuple_list = []
	# point_list = []
	# k = 0
	# l = -1
	# for i in element_list:
		# l = l + 1
		# k = k + 1
		# for j in range(k, len(element_list)):
			# if interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document) is None:
				# pass
			# else:
				# if interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document).ToString() not in tuple_list: 
					# point_list.append(interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document))
					# tuple_list.append(interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document).ToString())
	
	# t = Transaction(doc, 'Place bubbles')
	# t.Start()
	# for m in point_list:
		# print("Bubble placed")
		# instance = doc.Create.NewFamilyInstance(m, Bubble(bname).familysymbol, Structure.StructuralType.NonStructural)
	# t.Commit()
	
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
from System.Drawing import *
from System.Windows.Forms import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
copyOptions = CopyPasteOptions()

gm_collector = FilteredElementCollector(doc)\
	  .OfClass(FamilySymbol)\
	  .OfCategory(BuiltInCategory.OST_GenericModel)

gm_dict = {}
for i in gm_collector:
	gm_dict[i.FamilyName] = i

blist = ["init"]

class MaForm(Form):
	def __init__(self):
		self.blist = blist
		self.Text = 'Select the bubble family'
		screenSize = Screen.GetWorkingArea(self)
		self.Width = screenSize.Width / 3
		self.Height = screenSize.Height / 3
		self.CenterToScreen()

		self.listbox = ListBox()
		self.listbox.Parent = self
		self.listbox.Location = Point(0, 0)
		self.listbox.Size = Size(self.Width - 10, self.Height - 100)
		self.listbox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

		for i in gm_dict:
			self.listbox.Items.Add(i)

		self.listbox.SelectedIndexChanged += self.OnChanged

		self.button = Button()
		self.button.Parent = self
		self.button.Location = Point(10, self.listbox.Bottom + 10)
		self.button.Name = "1"
		self.button.Text = "OK"
		self.button.Click += self.OkButton

		self.sb = StatusBar()
		self.sb.Parent = self

	def OnChanged(self, sender, event):
		self.sb.Text = sender.SelectedItem

	def OkButton(self, sender, event):
		t = self.listbox.SelectedItem
		self.blist.append(t)
		self.Close()

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

Application.Run(MaForm())

if len(MaForm().blist) == 1:
	print("Please select a family")
	

else:
	if doc.ActiveView.GetType().ToString() == "Autodesk.Revit.DB.View3D":
		bname = MaForm().blist[1]
		# bname = "OBSERVATIONS_2017"
		# bname = "Observation_Libelle SYA"

		if Bubble(bname).presence is True:

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
			
		else:
			print("Oh Lord!")
		
	else:
		print("Please active a 3D view with a straight section box!")