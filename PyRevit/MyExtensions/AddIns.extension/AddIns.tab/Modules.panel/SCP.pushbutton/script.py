"""Select elements that match chosen categories and parameter value"""

__title__ = 'Select by category\nand parameter'

__doc__ = "Ce programme selection les elements de une ou plusieurs categories en fonction de la valeur d'un parametre choisi"

import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import TextInput

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class CheckBoxOption:
	def __init__(self, name, default_state=False):
		self.name = name
		self.state = default_state

	def __nonzero__(self):
		return self.state

	def __bool__(self):
		return self.state

view_list = ["La vue active", "Tout le document"]
res1 = forms.SelectFromList.show(view_list,
									multiselect = False,
									name_attr = "Vue",
									button_name = "OK")

if res1 != None:
	if res1[0] == "La vue active":
		collector = FilteredElementCollector(doc, doc.ActiveView.Id)
	else:
		collector = FilteredElementCollector(doc)

	categories = doc.Settings.Categories
	bic_dir = {}
	for bic in BuiltInCategory.GetValues(BuiltInCategory):
		try:
			cat = categories.get_Item(bic)
			bicName = cat.Name
			bic_dir[bicName] = bic
		except:
			"shit"

	bic_list = bic_dir.keys()
	options = [CheckBoxOption(j) for j in bic_list]

	all_checkboxes = forms.SelectFromCheckBoxes.show(options)

	res2 = []
	for checkbox in all_checkboxes:
		if checkbox.state == True:
			res2.append(checkbox.name)

	chosenCat_list = [bic_dir[x] for x in res2]

	icatcollection = List[ElementId]()
	elementFilter = List[ElementFilter](len(chosenCat_list))
	for i in chosenCat_list:
		# icatcollection.Add(i.Id)
		icatcollection.Add(categories.get_Item(i).Id)
		elementFilter.Add(ElementCategoryFilter(i))

	paramFilter = ParameterFilterUtilities.GetFilterableParametersInCommon(doc,icatcollection)

	bip_dir = {}
	bipName_dir = {}
	for bip in BuiltInParameter.GetValues(BuiltInParameter):
		bipId = str(ElementId(bip).IntegerValue)
		bip_dir[bipId] = bip
		try:
			bipName = LabelUtils.GetLabelFor(bip)
		except:
			bipName = bip.ToString()
		bipName_dir[bipId] = bipName

	param_dir = {}
	for paramId in paramFilter:
		if paramId.IntegerValue < 0:
			paramName = bipName_dir[str(paramId.IntegerValue)]
			param = bip_dir[str(paramId.IntegerValue)]
			param_dir[paramName] = param
		else:
			paramName = doc.GetElement(paramId).Name
			param = doc.GetElement(paramId)
			param_dir[paramName] = param


	res3 = forms.SelectFromList.show(param_dir.keys(),
										multiselect = False,
										name_attr = "Parametre",
										button_name = "OK")

	if res3 != None:
		value = TextInput('Value', default="Parameter value")
		chosenParam = param_dir[res3[0]]
		categoryFilter = LogicalOrFilter(elementFilter)
		element_collector = collector\
			.WhereElementIsNotElementType()\
			.WherePasses(categoryFilter)\
			.ToElements()

		icollection = List[ElementId]()
		for e in element_collector:
			# print(e)
			try:
				p = e.get_Parameter(chosenParam.GuidValue)
				# print(p)
				# print(p.StorageType)
				if (str(p.StorageType) == "Integer") and (str(p.AsString()) == str(value)):
					icollection.Add(e.Id)
				elif (str(p.StorageType) == "String") and (str(p.AsString()) == str(value)):
					icollection.Add(e.Id)
				elif (str(p.StorageType) == "Double") and (str(p.AsDouble()) == str(value)):
					icollection.Add(e.Id)
				elif (str(p.StorageType) == "ElementId") and (str(p.AsValueString()) == str(value)):
					icollection.Add(e.Id)
			except:
				p = e.get_Parameter(chosenParam)
				# print(p)
				# print(p.StorageType)
				if (str(p.StorageType) == "Integer") and (str(p.AsString()) == str(value)):
					icollection.Add(e.Id)
				elif (str(p.StorageType) == "String") and (str(p.AsString()) == str(value)):
					icollection.Add(e.Id)
				elif (str(p.StorageType) == "Double") and (str(p.AsDouble()) == str(value)):
					icollection.Add(e.Id)
				elif (str(p.StorageType) == "ElementId") and (str(p.AsValueString()) == str(value)):
					icollection.Add(e.Id)

		uidoc.Selection.SetElementIds(icollection)
	else:
		"bye"
else:
	"bye"