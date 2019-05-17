"""Paste elements from link model"""

__title__ = 'Paste element\nfrom link model'

__doc__ = 'This program copies and pastes elements from link model'

import clr
import System
import math
import sys
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import TextInput
# from pyautocad import Autocad

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
copyOptions = CopyPasteOptions()

# acad = Autocad()
# acad.prompt("Hello Autocad")
# print(acad.doc.Name)

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

def get_selected_elements(doc):
    try:
        # # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)

if doc.ActiveView.GetType().ToString() == "Autodesk.Revit.DB.View3D":

	try:
		el = get_selected_elements(doc)[0]
		linkdoc = el.GetLinkDocument()
	except:
		print("Please select a link model")
		sys.exit()

	# bpfilter = ElementCategoryFilter(BuiltInCategory.OST_ProjectBasePoint)
	# bp = FilteredElementCollector(doc).WherePasses(bpfilter).ToElements()[0]

	# xbp = bp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
	# ybp = bp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
	# zbp = bp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
	# print("xbp = " + str(xbp))
	# print("ybp = " + str(ybp))
	# print("zbp = " + str(zbp))

	try:
		bb = doc.ActiveView.GetSectionBox()
	except:
		bb = doc.ActiveView.SectionBox

	trans = bb.Transform
	docToOriginTrans = DocToOriginTransform(doc)
	originToDocTrans = OriginToDocTransform(linkdoc)
	docToDocTrans = originToDocTrans.Multiply(docToOriginTrans)
	print(docToDocTrans.OfPoint(trans.OfPoint(bb.Min)))
	print(docToDocTrans.OfPoint(trans.OfPoint(bb.Max)))
	print(trans.OfPoint(bb.Min))
	print(trans.OfPoint(bb.Max))
	try:
		a = ""
		outline = Outline(docToDocTrans.OfPoint(trans.OfPoint(bb.Min)), docToDocTrans.OfPoint(trans.OfPoint(bb.Max)))
		bbFilter = BoundingBoxIntersectsFilter(outline)
	except:
		a = "Same base point"
		outline = Outline(trans.OfPoint(bb.Min), trans.OfPoint(bb.Max))
		bbFilter = BoundingBoxIntersectsFilter(outline)

	# ElementOwnerViewFilter

	# element_collector = FilteredElementCollector(linkdoc)\
	# 	.OfClass(FamilyInstance)\
	# 	.WherePasses(bbFilter)\
	# 	.WhereElementIsNotElementType()\
	# 	.ToElements()

	element_collector = FilteredElementCollector(linkdoc)\
		.WherePasses(bbFilter)\
		.WhereElementIsNotElementType()\
		.ToElements()

	# elementId_collector = FilteredElementCollector(linkdoc)\
	# 	.OfClass(FamilyInstance)\
	# 	.WherePasses(bbFilter)\
	# 	.WhereElementIsNotElementType()\
	# 	.ToElementIds()

	elementId_collector = FilteredElementCollector(linkdoc)\
		.WherePasses(bbFilter)\
		.WhereElementIsNotElementType()\
		.ToElementIds()

	for e in element_collector:
		print(e)
		print(e.ViewSpecific)
		if e.ViewSpecific is True:
			elementId_collector.Remove(e.Id)

	t = Transaction(doc,"Copy and paste from selected link")
	t.Start()
	# copyOptions.SetDuplicateTypeNamesHandler(CopyUseDestination())
	if a == "Same base point":
		ElementTransformUtils.CopyElements(linkdoc, elementId_collector, doc, None, copyOptions)
	else:
		ElementTransformUtils.CopyElements(linkdoc, elementId_collector, doc, OriginToDocTransform(doc)\
			.Multiply(DocToOriginTransform(linkdoc)), copyOptions)
	doc.Regenerate()
	t.Commit()
else:
	print("Please active a 3D view with a straight section box!")