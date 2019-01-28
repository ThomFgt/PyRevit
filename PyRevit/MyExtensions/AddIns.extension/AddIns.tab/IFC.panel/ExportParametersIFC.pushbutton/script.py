"""Importe les parametres partages"""

__title__ = 'Export Parameters To IFC'

__doc__ = 'Configure the file .txt required to export the selected Revit parameter to an IFC file '


import clr
import System
import math
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI') 
import Autodesk
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import (Alert, select_file, TextInput, CommandLink, TaskDialog)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()

app = doc.Application

#Forms preparation
class CheckBoxOption:
    def __init__(self, name, default_state=False):
        self.name = name
        self.state = default_state

  # define the __nonzero__ method so you can use your objects in an 
  # if statement. e.g. if checkbox_option:
    def __nonzero__(self):
        return self.state

  # __bool__ is same as __nonzero__ but is for python 3 compatibility
    def __bool__(self):
        return self.state

def checkbox(itemsName):
  options = []
  for i in itemsName:
    options.append(CheckBoxOption(i))

  all_checkboxes = forms.SelectFromCheckBoxes.show(options)

  # now you can check the state of checkboxes in your program
  i = 0
  index = []
  for checkbox in all_checkboxes:
    if checkbox:
      index.append(i)
    i += 1

  itemsChoosed = []
  for i in index:
    itemsChoosed.append(itemsName[i])
  return itemsChoosed


#SharedParametersFile openning
spfile = app.OpenSharedParameterFile()

#Traduction of SharedParametersFile
spGroups = spfile.Groups
spGroupsDef = [g.Definitions for g in spGroups]
filepath = select_file('(.txt)|*.txt', restore_directory = True, title = "Select the mappage file (text file)")

spParamNames = []
spParamNamesChoosed = []
catIfcNames = ['IfcAlarmType', 'IfcAnnotation', 'IfcAirTerminal', 'IfcBeam', 'IfcBuildingElementProxy', 'IfcBuildingStorey', 'IfcColumn', 'IfcCovering', 'IfcCurtainWall', 'IfcDoor', 'IfcDuctSegment', 'IfcDuctFitting', 'IfcFlowTerminal', 'IfcFooting', 'IfcFurniture', 'IfcLightFixtureType', 'IfcMember', 'IfcPile', 'IfcPipeFitting', 'IfcPlate', 'IfcRailing', 'IfcRamp', 'IfcRoof', 'IfcSlab', 'IfcSpace', 'IfcSite', 'IfcStair', 'IfcValveType', 'IfcWall', 'IfcWindow']
catIfcNamesChoosed = []
commands = [CommandLink('Yes (you will choose the IFC categories for those parameters)', return_value = True), CommandLink('No', return_value = False)]
# Write the config parameter exporter
file = open(filepath, "w")

for spGroup in spGroups:
	for spParamDef in spGroup.Definitions:
		spParamNames.append(spParamDef.Name)
	dialog = TaskDialog("PropertySet : Ifc" + spGroup.Name, content = '\nDo you want to export for those parameters to the IFC model :\n' + '\n'.join(spParamNames), commands = commands)
	if dialog.show():
		# Alert('\nPlease choose the IFC categories for those parameters:\n' + '\n'.join(spParamNames), header = "PropertySet : Ifc" + spGroup.Name, title = "IFC category", exit = False)
		# spParamNamesChoosed = checkbox(spParamNames)
		catIfcNamesChoosed = checkbox(catIfcNames)
		spParamNames = []
		file.write("\n\n" + "PropertySet:" + "\tIFC" + spGroup.Name + "\t" + "I" + "\t" +  ','.join(catIfcNamesChoosed))
		for spParamDef in spGroup.Definitions:
			pType = str(spParamDef.ParameterType)
			if pType == "YesNo":
				pType = "Boolean"
			elif pType == "ElectricalPotential":
				pType = "ElectrialVoltage"
			elif pType == "ElectricalLuminousFlux":
				pType = "LuminousFlux"
			elif pType == "ElectricalLuminousIntensity":
				pType = "LuminousIntensity"
			elif pType == "PlaneAngle":
				pType = "Angle"
			elif pType == "'PipingPressure":
				pType = "Pressure"
			elif pType == "Number":
				pType = "Count"
			else:
				pType
	 		file.write("\n" + "\t" + "IFC" + spParamDef.Name + "\t" + pType + "\t" + spParamDef.Name)
file.close()