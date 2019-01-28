"""Pin all unpinned DWGs in doc"""

__title__ = 'Pin all unpinned\nDWGs in doc'

__doc__ = 'This program pin all unpinned DWGs in the current doc'

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

import_collector = FilteredElementCollector(doc)\
	  .OfClass(ImportInstance)\
	  .WhereElementIsNotElementType()\
	  .ToElements()

t = Transaction(doc, 'Pin all DWGs')
t.Start()
j = 0
for i in import_collector:
	if i.Pinned is False:
		j = j + 1
		print(i.Category.Name)
		i.Pinned = True
t.Commit()

if j == 0:
	print("     DWGs alreday pinned in doc")
else:
	print("     " + str(j) + " DWGs have been pinned")