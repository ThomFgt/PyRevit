import clr
import System
import math
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import (Alert, TextInput, FlexForm, Label, ComboBox, TextBox, TextBox, Separator, Button)


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
SEBoptions = SpatialElementBoundaryOptions()
roomcalculator = SpatialElementGeometryCalculator(doc)

td_button = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No

res = TaskDialog.Show("Importation from Excel","Attention :\n- Les ids des elements doivent etre en colonne 1\n- Les noms exacts (avec majuscules) des parametres partages doivent etre en ligne 1\n- Aucun accent ou caractere special dans le fichier Excel", td_button)

if res == TaskDialogResult.Yes:
               
  # Ungroup function 
  def Ungroup(group):
    group.UngroupMembers()
                  
  # Regroup function 
  def Regroup(groupname,groupmember):
    newgroup = doc.Create.NewGroup(groupmember)
    newgroup.GroupType.Name = str(groupname)
   
  def convertStr(s):
    try:
      ret = int(s)
    except ValueError:
      ret = 0
    return ret


  unfoundelements = []
  def feeding_parameter(array, param_names_excel, params_element_name, unfoundelements, missing_param):
    t = Transaction(doc, 'Feed elements')
    t.Start()
    count = 0
    for elementhash in array:
      idInt = int(elementhash['id'])
      try :
        element_id = ElementId(idInt)
        element = doc.GetElement(element_id)
        groupId = element.GroupId
        if str(groupId) != "-1":
          group = doc.GetElement(groupId)
          groupname = group.Name
          groupmember = group.GetMemberIds()
          Ungroup(group)                        

        for param in param_names_excel:
          if (param in params_element_name) and (param in elementhash):
            element.LookupParameter(param).Set(elementhash[param])
                                       
        if str(groupId) != "-1":
          Regroup(groupname,groupmember)
          # if "(membre exclu)" in group.GroupType.Name:
          # group.GroupType.Name = groupname
        count += 1
        print("element " + str(idInt) + " : OK - Avancement :" + str(count) + "/" + str(len(array)))
      except:
        count += 1
        print(str(idInt) + " not in REVIT doc")
        unfoundelements.append(idInt)                               
    print("Job done!")
    t.Commit()

    if len(unfoundelements) != 0:
      print(str(len(unfoundelements)) + " elements not found : ")
      print(unfoundelements)

    if nbr_missing_param != 0:
      print(str(len(missing_param)) + " parameters from Excel not found in Revit (wrong spelling in Excel or just not in this Revit document?) : ")
      print(missing_param)

  # Reading Excel file
  t = Transaction(doc, 'Read Excel spreadsheet.') 
  t.Start()
   
  #Accessing the Excel applications.
  xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
  count = 1

  dicWs = {}
  count = 1
  for i in xlApp.Worksheets:
    dicWs[i.Name] = i
    count += 1

  components = [Label('Pick a category:'),
            ComboBox('combobox2', {'Doors': 0, 'Rooms': 1}),
            Label('Select the name of Excel sheet to import:'),
            ComboBox('combobox', dicWs),
            Label('Enter the number of rows in Excel you want to integrate to Revit:'),
            TextBox('textbox', Text="60"),
            Label('Enter the number of colones in Excel you want to integrate to Revit:'),
            TextBox('textbox2', Text="2"),
            Separator(),
            Button('OK')]
  form = FlexForm('Title', components)
  form.show()

  worksheet = form.values['combobox']
  rowEnd = convertStr(form.values['textbox'])
  colEnd = convertStr(form.values['textbox2'])
  category = form.values['combobox2']

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

  #Row, and Column parameters
  rowStart = 1
  column_id = 1
  colStart = 2

  #Get parameters in the model
  params_element_set = collector[0].Parameters
  params_element_name = []
  for param_element in params_element_set:
    params_element_name.append(param_element.Definition.Name)

  # Using a loop to read a range of values and print them to the console.
  array = []
  param_names_excel = []
  data = {}
  for r in range(rowStart, rowEnd):
      data_id = worksheet.Cells(r, column_id).Text
      data_id_int = convertStr(data_id)
      if data_id_int != 0:
          data = {'id': data_id_int}
          for c in range(colStart, colEnd):
              data_param_value = worksheet.Cells(r, c).Text
              data_param_name = worksheet.Cells(1, c).Text
              if data_param_name != '':
                  param_names_excel.append(data_param_name)
                  if data_param_value != '':
                      data[data_param_name] = data_param_value
          array.append(data)
  t.Commit()

  #Check if parameters from Excel are in the model
  missing_param = []
  for param in param_names_excel:
    if (param not in params_element_name) and param not in (missing_param):
      missing_param.append(param)
  
  #Ask to continue if some parameters are missing
  nbr_missing_param = len(missing_param)          
  if nbr_missing_param != 0:
    miss_button = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
    miss = TaskDialog.Show("Parameters missing", "Some parameters in Excel document are missing in Revit : " + ", ".join(missing_param) + ".\nThe plugin cannot write in this parameters.\nDo you want to continue?" , miss_button)
    if miss == TaskDialogResult.Yes:
      feeding_parameter(array, param_names_excel, params_element_name, unfoundelements, missing_param)
    else:
      Alert('Come back when your changes are done', title = "End of the program", exit = True)
  else:
    feeding_parameter(array, param_names_excel, params_element_name, unfoundelements, missing_param)
else:
  Alert('A plus tard!', title = "End of the program", exit = True)