"""Bouge coupe de travail"""

__title__ = 'Move\nwork section'

__doc__ = 'Ce programme bouge la coupe de travail'

import os
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
        # # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)

try:
	el = get_selected_elements(doc)[0]
	start = el.Location.Curve.GetEndPoint(0)
	end = el.Location.Curve.GetEndPoint(1)
	cen = XYZ((start.X+end.X)/2,(start.Y+end.Y)/2,(start.Z+end.Z)/2)
	try:
		angle = math.atan((end.Y-start.Y)/(end.X-start.X))
	except:
		angle = math.pi/2

	view_collector = FilteredElementCollector(doc)\
		  .OfCategory(BuiltInCategory.OST_Views)\
		  .WhereElementIsNotElementType()\
		  .ToElements()
		  
	viewer_collector = FilteredElementCollector(doc)\
	  .OfCategory(BuiltInCategory.OST_Viewers)\
	  .WhereElementIsNotElementType()\
	  .ToElements()
		  
	user = os.path.expanduser('~').replace('C:\\Users\\','')	  
		  
	i = 0
	for view in view_collector:
		if ("Coupe" in view.Name) and (user in view.Name) and ("travail" in view.Name):
			myview = view
			i=i+1
			
	for viewer in viewer_collector:
		if ("Coupe" in viewer.Name) and (user in viewer.Name) and ("travail" in viewer.Name):
			myviewer = viewer

	if i==1:
		origine = myview.Origin
		
		try:
			dirangle = math.atan(myview.ViewDirection.Y/myview.ViewDirection.X)
		except:
			dirangle = math.pi/2
			
		axis = Line.CreateBound(origine,origine.Add(XYZ(0,0,1)))			
		t = Transaction(doc, 'Rotation & translation')
		t.Start()
		

		ElementTransformUtils.RotateElement(doc,myviewer.Id,axis,float(angle-dirangle)+math.pi/2)
		ElementTransformUtils.MoveElement(doc,myviewer.Id,XYZ((cen.X-origine.X),(cen.Y-origine.Y),0))
		
		t.Commit()
				
		myviewer_icollection = List[ElementId](1)
		myviewer_icollection.Add(myviewer.Id)
	
		uidoc.Selection.SetElementIds(myviewer_icollection)
		
	elif i==0:
		print('Pas de coupe nomm'+str("\xe9")+'e "Coupe travail '+ user + '"!')

	else :
		print("Il ne faut qu'une coupe de travail!") 
	
except:
	try:
		t.Commit()
		print("Il faut selectionner un reseau! Ou une poutre...\n")
	except:
		print("Il faut selectionner un reseau! Ou une poutre...\n")