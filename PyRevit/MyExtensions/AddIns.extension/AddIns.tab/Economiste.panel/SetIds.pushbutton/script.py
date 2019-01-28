"""Remplit les IDs des portes"""

__title__ = 'Set IDs'

__doc__ = 'Ce programme remplit les IDs des portes ou des pieces'

import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from rpw.ui.forms import TextInput
from pyrevit import forms
from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
SEBoptions = SpatialElementBoundaryOptions()

components = [Label('Pick a category:'),
              ComboBox('combobox', {'Doors': 0, 'Rooms': 1}),
              Label('Enter the name of ID parameter'),
              Label('(must be created like a integer - nombre entier - ) :'),
              TextBox('textbox', Text="ID Revit"),
              Separator(),
              Button('OK')]
form = FlexForm('Title', components)
form.show()
# User selects `Opt 1`, types 'Wood' in TextBox, and select Checkbox

parameterName = form.values['textbox']
category = form.values['combobox']

if category == 0:
	collector = FilteredElementCollector(doc)\
		 		.OfCategory(BuiltInCategory.OST_Doors)\
		  		.WhereElementIsNotElementType()\
		  		.ToElements()
else:
	collector = FilteredElementCollector(doc)\
		 		.OfCategory(BuiltInCategory.OST_Rooms)\
		  		.WhereElementIsNotElementType()\
		  		.ToElements()

t = Transaction(doc, 'Fill Ids')
t.Start()

for i in collector:
	
	print(i.Id.IntegerValue)
	i.LookupParameter(parameterName).Set(i.Id.IntegerValue)
	
t.Commit()