"""Pin all unpinned DWGs and REVIT links in doc"""

__title__ = 'Pin all unpinned DWGs\nand REVIT links in doc'

__doc__ = 'This program pin all unpinned DWGs and REVIT links in the current doc'

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

DWGimport_collector = FilteredElementCollector(doc)\
	  .OfClass(ImportInstance)\
	  .WhereElementIsNotElementType()\
	  .ToElements()

RVTimport_collector = FilteredElementCollector(doc)\
	  .OfClass(RevitLinkInstance)\
	  .WhereElementIsNotElementType()\
	  .ToElements()

t = Transaction(doc, 'Pin all DWGs and REVIT links')
t.Start()
j = 0
for i in DWGimport_collector:
	if i.Pinned is False:
		j = j + 1
		print(i.Category.Name + " --> pinned")
		i.Pinned = True
		
l = 0
for k in RVTimport_collector:
	if k.Pinned is False:
		l = l + 1
		print(k.Name + " --> pinned")
		k.Pinned = True
t.Commit()

print("----------")

if j == 0:
	print("     DWGs alreday pinned in doc")
else:
	print("     " + str(j) + " DWGs have been pinned")
	
if l == 0:
	print("     REVIT links alreday pinned in doc")
else:
	print("     " + str(l) + " REVIT links have been pinned")