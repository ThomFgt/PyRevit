"""Import a list of points in order to edit the shape of a slab"""

__title__ = 'Edit Slab Shape'

__doc__ = 'Needed : an Excel list of points\' coordinates (in centimeters), with for first colomn x, second y and third z'


import clr
import System
from System.Collections.Generic import List
import Autodesk
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitNodes')
clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
import Revit
clr.ImportExtensions(Revit.Elements)
from RevitServices.Transactions import TransactionManager
from rpw.ui.forms import (select_file, Alert, TextInput, FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button, CommandLink, TaskDialog)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

commands = [CommandLink('OK', return_value = True), CommandLink('No, let me change my Excel please', return_value = False)]

dialog = TaskDialog("Warning Form", content = 'The Excel sheet to import must be open with for first colomn x, second y and third z', commands = commands)


if dialog.show():
  def convertStr(s):
    try:
      ret = int(s)
    except ValueError:
      ret = None
    return ret

  # Get the Excel file
  t = Transaction(doc, 'Read Excel spreadsheet.') 
  t.Start()

  #Accessing the Excel applications.
  try:
    xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
  except:
    Alert('The Excel sheet must be open (the program failed because of a problem with Excel)', title = "Failed", exit = True)

  dicWs = {}
  count = 1
  for i in xlApp.Worksheets:
      dicWs[i.Name] = i
      count += 1

  components = [Label('Select the name of Excel sheet to import :'),
            ComboBox('combobox', dicWs),
            Label('Choose the units you used in your Excel sheet :'),
            ComboBox('combobox2', {'Meters': 3.048, 'Decimeters': 30.48, 'Centimeters': 304.8, 'Millimeters': 3048}),
            Label('Enter the number of points:'),
            TextBox('textbox', Text="11"),
            Separator(),
            Button('OK')]
  form = FlexForm('Title', components)
  form.show()

  worksheet = form.values['combobox']
  rowEnd = convertStr(form.values['textbox'])
  array = []
  fails = []
  for r in range(1, rowEnd):
    try:
      x = convertStr(worksheet.Cells(r, 1).Text)/form.values['combobox2']
      y = convertStr(worksheet.Cells(r, 2).Text)/form.values['combobox2']
      z = convertStr(worksheet.Cells(r, 3).Text)/form.values['combobox2']
      array.append(XYZ(x,y,z))
    except:
      fails.append(str(r))
  t.Commit()

  commands = [CommandLink('OK', return_value = True), CommandLink('No, let me change my values please', return_value = False)]

  dialog = TaskDialog("Warning values unsupported", content = 'The units in Excel spreadsheet at lines : '+ ','.join(fails) + ' are not numbers', commands = commands)

  if dialog.show():
    Alert('Pick a slab please (clicking on it)', title = "Select a slab", exit = False)
    # Pick an element
    sel = uidoc.Selection
    obType = Selection.ObjectType.Element
    ref = sel.PickObject(obType, "Select Element.")
    element = doc.GetElement(ref.ElementId)

    t = Transaction(doc, 'Points')
    t.Start()

    for p in array:
      element.SlabShapeEditor.DrawPoint(p)
    t.Commit()

