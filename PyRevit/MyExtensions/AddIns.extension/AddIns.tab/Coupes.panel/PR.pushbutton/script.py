"""Rotation de 90 degres la coupe"""

__title__ = 'Perpendicular\nsection rotation'

__doc__ = 'Ce programme fait une rotation de 90 degres de la coupe selectionnee'
	
import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()

def get_selected_elements(doc):
    try:
        # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)

el = get_selected_elements(doc)[0]

view_collector = FilteredElementCollector(doc)\
  .OfCategory(BuiltInCategory.OST_Views)\
  .WhereElementIsNotElementType()\
  .ToElements()
  
for i in view_collector:
	if i.Name == el.Name:
		myview = i
		
try:
	if myview.GetType().ToString() == "Autodesk.Revit.DB.ViewSection":
		origin = myview.Origin
		axis = Line.CreateBound(origin,origin.Add(XYZ(0,0,1)))

		t = Transaction(doc, 'Invert section')
		t.Start()

		ElementTransformUtils.RotateElement(doc,el.Id,axis,math.pi/2)
		
		t.Commit()
		
except:
	print("Ceci n'est pas une coupe!")