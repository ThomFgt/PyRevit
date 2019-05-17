# # # # # Clash detection 2 # # # # #

__title__ = 'Place bubbles\nin active 3D view'

__doc__ = 'Ce programme place les bulles dans la vue 3D active'

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

def interpoint(el1, el2, doc1, doc2):

	doc = __revit__.ActiveUIDocument.Document
	uidoc = __revit__.ActiveUIDocument
	options = __revit__.Application.Create.NewGeometryOptions()
	intersectOptions = SolidCurveIntersectionOptions()

	if (el1.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and ((el2.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance")):
		return None
	else:
		x = 0
		if doc1.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
			x = x + 1
			trans1 = doc1.GetTotalTransform()
			bb1 = el1.get_Geometry(options).GetBoundingBox()
			outline1 = Outline(trans1.OfPoint(bb1.Min),trans1.OfPoint(bb1.Max))
		else:
			bb1 = el1.get_Geometry(options).GetBoundingBox()
			outline1 = Outline(bb1.Min,bb1.Max)
			
		if doc2.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
			x = x + 1
			trans2 = doc2.GetTotalTransform()
			bb2 = el2.get_Geometry(options).GetBoundingBox()
			outline2 = Outline(trans2.OfPoint(bb2.Min),trans2.OfPoint(bb2.Max))
		else:
			bb2 = el2.get_Geometry(options).GetBoundingBox()
			outline2 = Outline(bb2.Min,bb2.Max)
	
		filter = BoundingBoxIntersectsFilter(outline1)
	
		result = outline1.Intersects(outline2, 0)
		
		if result is True:
			el1_solids_list = []
			el1_curves_list = []
			for i in el1.get_Geometry(options):
				if i.GetType().ToString() == "Autodesk.Revit.DB.Solid":
					el1_solids_list.append(i)
					for j in i.Edges:
						el1_curves_list.append(j.AsCurve())	
		
			try:
				el1_curves_list.append(el1.Location.Curve)
			except:
				"Oh God"
				
			try:
				if "Autodesk.Revit.DB.Arc" in el1_curves_list[0].GetType().ToString():
					try:
						c1 = Line.CreateBound(trans1.OfPoint(el1_curves_list[0].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[6].GetEndPoint(0)))
						c2 = Line.CreateBound(trans1.OfPoint(el1_curves_list[1].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[7].GetEndPoint(0)))
						c3 = Line.CreateBound(trans1.OfPoint(el1_curves_list[2].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[4].GetEndPoint(0)))
						c4 = Line.CreateBound(trans1.OfPoint(el1_curves_list[3].GetEndPoint(0)),trans1.OfPoint(el1_curves_list[5].GetEndPoint(0)))
					except:
						c1 = Line.CreateBound(el1_curves_list[0].GetEndPoint(0),el1_curves_list[6].GetEndPoint(0))
						c2 = Line.CreateBound(el1_curves_list[1].GetEndPoint(0),el1_curves_list[7].GetEndPoint(0))
						c3 = Line.CreateBound(el1_curves_list[2].GetEndPoint(0),el1_curves_list[4].GetEndPoint(0))
						c4 = Line.CreateBound(el1_curves_list[3].GetEndPoint(0),el1_curves_list[5].GetEndPoint(0))
					el1_curves_list.extend([c1, c2, c3, c4])
			except:
				"Oh God"
		
			
			el2_solids_list = []
			el2_curves_list = []
			for i in el2.get_Geometry(options):
				if i.GetType().ToString() == "Autodesk.Revit.DB.Solid":
					el2_solids_list.append(i)
					for j in i.Edges:
						el2_curves_list.append(j.AsCurve())	
		
			try:
				el2_curves_list.append(el2.Location.Curve)
			except:
				"Oh God"
				
			try:
				if "Autodesk.Revit.DB.Arc" in el2_curves_list[0].GetType().ToString():
					try:
						c1 = Line.CreateBound(trans2.OfPoint(el2_curves_list[0].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[6].GetEndPoint(0)))
						c2 = Line.CreateBound(trans2.OfPoint(el2_curves_list[1].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[7].GetEndPoint(0)))
						c3 = Line.CreateBound(trans2.OfPoint(el2_curves_list[2].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[4].GetEndPoint(0)))
						c4 = Line.CreateBound(trans2.OfPoint(el2_curves_list[3].GetEndPoint(0)),trans2.OfPoint(el2_curves_list[5].GetEndPoint(0)))
					except:
						c1 = Line.CreateBound(el2_curves_list[0].GetEndPoint(0),el2_curves_list[6].GetEndPoint(0))
						c2 = Line.CreateBound(el2_curves_list[1].GetEndPoint(0),el2_curves_list[7].GetEndPoint(0))
						c3 = Line.CreateBound(el2_curves_list[2].GetEndPoint(0),el2_curves_list[4].GetEndPoint(0))
						c4 = Line.CreateBound(el2_curves_list[3].GetEndPoint(0),el2_curves_list[5].GetEndPoint(0))
					el2_curves_list.extend([c1, c2, c3, c4])
			except:
				"Oh God"
			
			if len(el1_solids_list)>len(el2_solids_list):
				solids_list = el1_solids_list
				curves_list = el2_curves_list
			else:
				solids_list = el2_solids_list
				curves_list = el1_curves_list
			
			for k in solids_list:
				for l in curves_list:
					try:
						solidcurveintersection = k.IntersectWithCurve(l,intersectOptions)
						seg_nb = solidcurveintersection.SegmentCount
						if seg_nb != 0:
							if x == 0:
								point = solidcurveintersection.GetCurveSegment(0).GetEndPoint(0)
								break
							else:
								try:
									point = trans1.OfPoint(solidcurveintersection.GetCurveSegment(0).GetEndPoint(0))
								except:
									point = trans2.OfPoint(solidcurveintersection.GetCurveSegment(0).GetEndPoint(0))
								break
					except:
						pass
			
			try:
				return point
			except:
				return None
		else:
			return None

bname = "OBSERVATIONS_2017"

if Bubble(bname).presence is True:
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
				outline = Outline(trans.OfPoint(bbmin),trans.OfPoint(bbmax))
				bbFilter = BoundingBoxIntersectsFilter(outline)
				element_collector = FilteredElementCollector(dl)\
					.OfCategory(cat)\
					.WherePasses(bbFilter)\
					.WhereElementIsNotElementType()\
					.ToElements()
				for e in element_collector:
					if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
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
					a = "Same base point"
					outline = Outline(trans.OfPoint(bbmin), trans.OfPoint(bbmax))
					bbFilter = BoundingBoxIntersectsFilter(outline)
				element_collector = FilteredElementCollector(dl.GetLinkDocument())\
					.OfCategory(cat)\
					.WherePasses(bbFilter)\
					.WhereElementIsNotElementType()\
					.ToElements()
				for e in element_collector:
					if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
						pass
					else:
						element_list.append(e)

	print(element_list)
	
	tuple_list = []
	point_list = []
	k = 0
	l = -1
	for i in element_list:
		l = l + 1
		k = k + 1
		for j in range(k, len(element_list)):
			if interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document) is None:
				pass
			else:
				if interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document).ToString() not in tuple_list: 
					point_list.append(interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document))
					tuple_list.append(interpoint(i, element_list[j], element_list[l].Document, element_list[j].Document).ToString())
	
	t = Transaction(doc, 'Place bubbles')
	t.Start()
	for m in point_list:
		print("Bubble placed")
		instance = doc.Create.NewFamilyInstance(m, Bubble(bname).familysymbol, Structure.StructuralType.NonStructural)
	t.Commit()