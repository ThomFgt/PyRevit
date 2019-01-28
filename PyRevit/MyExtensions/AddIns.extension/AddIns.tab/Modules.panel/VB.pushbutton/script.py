"""Opens requested view"""

__title__ = 'Browser that opens\nrequested view'

__doc__ = 'Browser that opens requested view'

import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
SEBoptions = SpatialElementBoundaryOptions()
roomcalculator = SpatialElementGeometryCalculator(doc)

logger = script.get_logger()

view_collector = FilteredElementCollector(doc)\
	.OfCategory(BuiltInCategory.OST_Views)\
	.WhereElementIsNotElementType()\
	.ToElements()

# view_list = []
# for i in view_collector:
# 	view_list.append(str(i.Name))

view_dir = {}
i = 0
for view in view_collector:
	view_dir[view.Name] = str(i)
	i = i + 1

view_list = view_dir.keys()

# matched_str = 'target1'
matched_str = view_list
args = ['--help', '--branch', 'branchname']
# switches = {'/switch1': True, '/switch2': False}
OK_switch = '/OK'
Cancel_switch = '/Cancel'
switches = [OK_switch, Cancel_switch]

sb = matched_str, args, switches
sb = forms.SearchPrompt.show(
	search_db = view_list,
	switches= [OK_switch,
	Cancel_switch],
	search_tip = 'View browser'
	)

if sb[0] != None:
	# for view in view_collector:
	# 	if view.Name == sb[0]:
	# 		myview = view
	myview = view_collector[int(view_dir[sb[0]])]
	uidoc.RequestViewChange(myview)