"""Importe les parametres partages"""

__title__ = 'Import Shared Parameters'

__doc__ = 'Ce programme importe les parametres partages dans un projet'


import clr
import System
import math
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import (SelectFromList, Alert)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()

app = doc.Application

#SharedParametersFile openning
spfile = app.OpenSharedParameterFile()

#Traduction of SharedParametersFile
spGroups = spfile.Groups
spGroupsDef = [g.Definitions for g in spGroups]
spParametersDef = [x for l in spGroupsDef for x in l]
#Names of groups
spGroupsName = [x.Name for x in spGroups]
#Names of SharedParameters
spParametersName = [x.Name for x in spParametersDef]

categoriesDef = [i for i in doc.Settings.Categories]
categoriesName = [i.Name for i in doc.Settings.Categories]

gp = {}
gpNames = [LabelUtils.GetLabelFor(i) for i in BuiltInParameterGroup.GetValues(BuiltInParameterGroup)]
for i in BuiltInParameterGroup.GetValues(BuiltInParameterGroup):
  gp[LabelUtils.GetLabelFor(i)] = i

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

def checkbox(itemsDef, itemsName):
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
    itemsChoosed.append(itemsDef[i])
  return itemsChoosed

#Save the choices of the user
categoriesChoosed = checkbox(categoriesDef, categoriesName)
spParametersChoosed = checkbox(spParametersDef, spParametersName)
gpStringChoosed = SelectFromList('Parameter Group', gpNames)

#creating category set
catset = app.Create.NewCategorySet()
[catset.Insert(j) for j in categoriesChoosed]

#Import the parameter in the parameter group choosed by user
group = gp[gpStringChoosed]

#Instances parameters
bind = app.Create.NewInstanceBinding(catset)

t = Transaction(doc, 'Fill door Ids')
t.Start()

bindmap = doc.ParameterBindings
for p in spParametersChoosed:
  try:
    bindmap.Insert(p, bind, group)
  except:
    continue
  
t.Commit()