"""Bulle la vue 3D active"""

__title__ = 'Place bubbles\nin active 3D view'

__doc__ = 'Ce programme place les bulles dans la vue 3D active'

import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List

# # # # # ATTENTION : SI BUBBLE NON PLACEE, MARCHE PAS
class Bubble:

	doc = __revit__.ActiveUIDocument.Document
	uidoc = __revit__.ActiveUIDocument
	options = __revit__.Application.Create.NewGeometryOptions()
	intersectOptions = SolidCurveIntersectionOptions()

	def __init__(self,bubble_name):
		self.familysymbol = None		
		self.presence = False
		
		gm_collector = FilteredElementCollector(doc)\
			  .OfClass(FamilySymbol)\
			  .OfCategory(BuiltInCategory.OST_GenericModel)
				
		for i in gm_collector:
			if i.FamilyName == str(bubble_name):
				#self.familyinstance = i
				#self.familysymbol = i.Symbol
				self.familysymbol = i
				self.presence = True
				break
				
		return self.familysymbol
		return self.presence

# # # # # ATTENTION : TROUVER LE MOYEN DE VIRER LES ELEMENTS NON VISIBLES DANS LA VUE
class ListElements:

	def __init__(self):
		self.element_list = []
		self.doc_list = []

		doc = __revit__.ActiveUIDocument.Document
		uidoc = __revit__.ActiveUIDocument
		options = __revit__.Application.Create.NewGeometryOptions()
		
		SectionBox = doc.ActiveView.GetSectionBox()
		trans = SectionBox.Transform
		min = SectionBox.Min
		max = SectionBox.Max
		
		outline = Outline(trans.OfPoint(min),trans.OfPoint(max))
		
		# cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting, BuiltInCategory.OST_DuctCurves,\
					# BuiltInCategory.OST_DuctFitting, BuiltInCategory.OST_CableTray,\
					# BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_StructuralFraming]
					
		# cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
		# 			BuiltInCategory.OST_CableTray]
					
		cat_list = [BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_DuctCurves,\
					BuiltInCategory.OST_CableTray, BuiltInCategory.OST_StructuralFraming]
		
		rvtlink_collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
				  .OfCategory(BuiltInCategory.OST_RvtLinks)\
				  .WhereElementIsNotElementType()\
				  .ToElements()
		
		dl_list = []
		for rvtlink in rvtlink_collector:
			dl_list.append(rvtlink)
		dl_list.append(doc)

		for cat in cat_list:
			for dl in dl_list:
				if dl.GetType().ToString() == "Autodesk.Revit.DB.Document":
					element_collector = FilteredElementCollector(dl)\
						.OfCategory(cat)\
						.WhereElementIsNotElementType()\
						.ToElements()
					for e in element_collector:
						if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
							pass
						else:
							try:
								bb2 = e.get_Geometry(options).GetBoundingBox()
								outline2 = Outline(bb2.Min,bb2.Max)
								filter1 = outline.ContainsOtherOutline(outline2,0)
								filter2 = outline.Intersects(outline2,0)
								if (filter1 is True) or (filter2 is True):
									self.element_list.append(e)
									self.doc_list.append(dl)
							except:
								pass
				elif dl.GetType().ToString() == "Autodesk.Revit.DB.RevitLinkInstance":
					element_collector = FilteredElementCollector(dl.GetLinkDocument())\
						.OfCategory(cat)\
						.WhereElementIsNotElementType()\
						.ToElements()
					for e in element_collector:
						if (e.GetType().ToString() == "Autodesk.Revit.DB.FamilyInstance") and (e.StructuralType.ToString() != "Beam"):
							pass
						else:
							try:
								trans2 = dl.GetTotalTransform()
								bb2 = e.get_Geometry(options).GetBoundingBox()
								outline2 = Outline(trans2.OfPoint(bb2.Min),trans2.OfPoint(bb2.Max))
								filter1 = outline.ContainsOtherOutline(outline2,0)
								filter2 = outline.Intersects(outline2,0)
								if (filter1 is True) or (filter2 is True):
									self.element_list.append(e)
									self.doc_list.append(dl)
							except:
								pass
						
		return self.element_list
		return self.doc_list
		
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
				
			# # # # # TEST
			
			for k in solids_list:
				for l in curves_list:
					try:
						solidcurveintersection = k.IntersectWithCurve(l,intersectOptions)
						seg_nb = solidcurveintersection.SegmentCount
						# if seg_nb <> 0:
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
			
			# for k in solids_list:
				# for l in curves_list:
					# solidcurveintersection = k.IntersectWithCurve(l,intersectOptions)
					# seg_nb = solidcurveintersection.SegmentCount
					# if seg_nb <> 0:
						# if x == 0:
							# point = solidcurveintersection.GetCurveSegment(0).GetEndPoint(0)
							# break
						# else:
							# try:
								# point = trans1.OfPoint(solidcurveintersection.GetCurveSegment(0).GetEndPoint(0))
							# except:
								# point = trans2.OfPoint(solidcurveintersection.GetCurveSegment(0).GetEndPoint(0))
							# break
				
			# # # # # TEST
			try:
				return point
			except:
				return None
		else:
			return None

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
intersectOptions = SolidCurveIntersectionOptions()

# # # # # ATTENTION : AJOUTER CHECK SI VUE ACTIVE IS 3DVIEW

# bname = "OBSERVATIONS_Test_NTA"
bname = "Observation_Libelle SYA"

if Bubble(bname).presence is True:
	
	e_list = ListElements().element_list
	d_list = ListElements().doc_list
	
	tuple_list = []
	point_list = []
	k = 0
	l = -1
	for i in e_list:
		l = l + 1
		k = k + 1
		for j in range(k, len(e_list)):
			if interpoint(i, e_list[j], d_list[l], d_list[j]) is None:
				pass
			else:
				if interpoint(i, e_list[j], d_list[l], d_list[j]).ToString() not in tuple_list: 
					point_list.append(interpoint(i, e_list[j], d_list[l], d_list[j]))
					tuple_list.append(interpoint(i, e_list[j], d_list[l], d_list[j]).ToString())
	
	t = Transaction(doc, 'Place bubbles')
	t.Start()
	for m in point_list:
		print("Bubble placed")
		instance = doc.Create.NewFamilyInstance(m, Bubble(bname).familysymbol, Structure.StructuralType.NonStructural)
	t.Commit()
	
else:
	print("Oh Lord!")