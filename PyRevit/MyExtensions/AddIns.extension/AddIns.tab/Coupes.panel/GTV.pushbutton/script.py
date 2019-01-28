"""Aller dans la vue"""

__title__ = 'Go to\nselected view'

__doc__ = 'Ce programme permet de basculer dans la vue selectionnee'
	
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

try:
	el = get_selected_elements(doc)[0]

	view_collector = FilteredElementCollector(doc)\
	  .OfCategory(BuiltInCategory.OST_Views)\
	  .WhereElementIsNotElementType()\
	  .ToElements()
	  
	for view in view_collector:
		if view.Name == el.Name:
			myview = view
			
	uidoc.RequestViewChange(myview)
except:
	print("L'element selectionne n'est pas une vue")