# import clr
# import math
# clr.AddReference('RevitAPI') 
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from System.Collections.Generic import List

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument
# options = __revit__.Application.Create.NewGeometryOptions()
# SEBoptions = SpatialElementBoundaryOptions()
# roomcalculator = SpatialElementGeometryCalculator(doc)

# class FailureHandler(IFailuresPreprocessor):
#   def __init__(self):
#     self.ErrorMessage = ""
#     self.ErrorSeverity = ""
#   def PreprocessFailures(self, failuresAccessor):
#   	# failuresAccessor.DeleteAllWarning()
#   	# return FailureProcessingResult.Continue
#   	failures = failuresAccessor.GetFailureMessages()
#   	rslt = ""
#   	for f in failures:
#   		fseverity = failuresAccessor.GetSeverity()
#   		if fseverity == FailureSeverity.Warning:
#   			failuresAccessor.DeleteWarning(f)
#   		elif fseverity == FailureSeverity.Error:
#   			rslt = "Error"
#   			failuresAccessor.ResolveFailure(f)
#   	if rslt == "Error":
#   		return FailureProcessingResult.ProceedWithCommit
#   		# return FailureProcessingResult.ProceedWithRollBack
#   	else:
#   		return FailureProcessingResult.Continue

# group_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
# 	  .OfClass(Group)

# def get_selected_elements(doc):
#     try:
#         # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def Ungroup(group):
# 	group.UngroupMembers()
	
# def Regroup(groupname,groupmember):
# 	newgroup = doc.Create.NewGroup(groupmember)
# 	newgroup.GroupType.Name = str(groupname)
	
# # group = get_selected_elements(doc)[0]

# # print(group.Id)
# # print(group.GetType())
# # print(group.GroupId)

# # groupname = group.Name
# # groupmember = group.GetMemberIds()

# # IdsInLine = ""
# # for i in groupmember:
# # 	IdsInLine = IdsInLine + str(i.IntegerValue) + ", "

# # IdsInLine = IdsInLine[:len(IdsInLine)-3]

# # print(IdsInLine)

# # status = ""
# # t1 = Transaction(doc, 'Ungroup group')
# # t1.Start()
# # Ungroup(group)
# # print("Group ungrouped")
# # t1.Commit()
# # try:
# # 	t2 = Transaction(doc, 'Regroup group')
# # 	t2.Start()

# # 	failureHandlingOptions = t2.GetFailureHandlingOptions()
# # 	failureHandler = FailureHandler()
# # 	failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
# # 	failureHandlingOptions.SetClearAfterRollback(True)
# # 	t2.SetFailureHandlingOptions(failureHandlingOptions)

# # 	Regroup(groupname,groupmember)
# # 	print("Group regrouped")
# # 	status = "Yeah!"
# # 	t2.Commit()

# # except:
# # 	t2.RollBack()
# # 	print("Regrouping fail")
# # 	status = "Fuck!"

# # print(status + "\n")
# # print(t2.GetStatus())





# tg = TransactionGroup(doc, 'Ungroup/regroup all groups')
# tg.Start()

# for group in group_collector:
# 	print(group.Id)
# 	print(group.GroupId)

# 	groupname = group.Name
# 	print("Group name : " + groupname)
# 	groupmember = group.GetMemberIds()

# 	t1 = Transaction(doc, 'Ungroup group')
# 	t1.Start()
# 	Ungroup(group)
# 	print("Group ungrouped")
# 	t1.Commit()
# 	try:
# 		t2 = Transaction(doc, 'Regroup group')
# 		t2.Start()

# 		failureHandlingOptions = t2.GetFailureHandlingOptions()
# 		failureHandler = FailureHandler()
# 		failureHandlingOptions.SetFailuresPreprocessor(failureHandler)
# 		failureHandlingOptions.SetClearAfterRollback(True)
# 		t2.SetFailureHandlingOptions(failureHandlingOptions)

# 		Regroup(groupname,groupmember)
# 		print("Group regrouped")
# 		t2.Commit()

# 	except:
# 		t2.RollBack()
# 		print("Regrouping fail")
# 		IdsInLine = ""
# 		for i in groupmember:
# 			IdsInLine = IdsInLine + str(i.IntegerValue) + ", "

# 		IdsInLine = IdsInLine[:len(IdsInLine)-3]

# 		print("Grouped element ids was : " + IdsInLine)

# print("Done")

# tg.Assimilate()














# # # # # ATTENTION MODULE FORM DANS REVIT
import sys
import clr
import System
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
from System.Drawing import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from System.Windows.Forms import *
from System.Collections.Generic import List
from pyrevit import forms
from rpw.ui.forms import TextInput
from pyrevit import framework
from pyrevit.framework import Interop
from pyrevit.api import AdWindows
from pyrevit.framework import wpf, Forms, Controls, Media

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
# app = __revit__.Application
app = uidoc.Application.Application



# # # # # FORM CHELOU
# class WPFWindow(framework.Windows.Window):
#     def __init__(self, xaml_source, literal_string=False):
#         # self.Parent = self
#         wih = Interop.WindowInteropHelper(self)
#         wih.Owner = AdWindows.ComponentManager.ApplicationWindow

#         if not literal_string:
#             if not op.exists(xaml_source):
#                 wpf.LoadComponent(self,
#                                   os.path.join(EXEC_PARAMS.command_path,
#                                                xaml_source)
#                                   )
#             else:
#                 wpf.LoadComponent(self, xaml_source)
#         else:
#             wpf.LoadComponent(self, framework.StringReader(xaml_source))

#         #2c3e50
#         self.Resources['pyRevitDarkColor'] = \
#             Media.Color.FromArgb(0xFF, 0x2c, 0x3e, 0x50)

#         #23303d
#         self.Resources['pyRevitDarkerDarkColor'] = \
#             Media.Color.FromArgb(0xFF, 0x23, 0x30, 0x3d)

#         #ffffff
#         self.Resources['pyRevitButtonColor'] = \
#             Media.Color.FromArgb(0xFF, 0xff, 0xff, 0xff)

#         #f39c12
#         self.Resources['pyRevitAccentColor'] = \
#             Media.Color.FromArgb(0xFF, 0xf3, 0x9c, 0x12)

#         self.Resources['pyRevitDarkBrush'] = \
#             Media.SolidColorBrush(self.Resources['pyRevitDarkColor'])
#         self.Resources['pyRevitAccentBrush'] = \
#             Media.SolidColorBrush(self.Resources['pyRevitAccentColor'])

#         self.Resources['pyRevitDarkerDarkBrush'] = \
#             Media.SolidColorBrush(self.Resources['pyRevitDarkerDarkColor'])

#         self.Resources['pyRevitButtonForgroundBrush'] = \
#             Media.SolidColorBrush(self.Resources['pyRevitButtonColor'])

#     def show(self, modal=False):
#         if modal:
#             return self.ShowDialog()
#         self.Show()

#     def show_dialog(self):
#         return self.ShowDialog()

#     def set_image_source(self, element_name, image_file):
#         wpfel = getattr(self, element_name)
#         if not op.exists(image_file):
#             # noinspection PyUnresolvedReferences
#             wpfel.Source = \
#                 framework.Imaging.BitmapImage(
#                     framework.Uri(os.path.join(EXEC_PARAMS.command_path,
#                                                image_file))
#                     )
#         else:
#             wpfel.Source = \
#                 framework.Imaging.BitmapImage(framework.Uri(image_file))

#     @staticmethod
#     def hide_element(*wpf_elements):
#         for wpfel in wpf_elements:
#             wpfel.Visibility = framework.Windows.Visibility.Collapsed

#     @staticmethod
#     def show_element(*wpf_elements):
#         for wpfel in wpf_elements:
#             wpfel.Visibility = framework.Windows.Visibility.Visible

#     @staticmethod
#     def toggle_element(*wpf_elements):
#         for wpfel in wpf_elements:
#             if wpfel.Visibility == framework.Windows.Visibility.Visible:
#                 self.hide_element(wpfel)
#             elif wpfel.Visibility == framework.Windows.Visibility.Collapsed:
#                 self.show_element(wpfel)

# class TemplateUserInputWindow(WPFWindow):
#     layout = """
#     <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
#             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
#             ShowInTaskbar="False" ResizeMode="NoResize"
#             WindowStartupLocation="CenterScreen"
#             HorizontalContentAlignment="Center">
#     </Window>
#     """
#     def __init__(self, context, title, width, height, **kwargs):
#         WPFWindow.__init__(self, self.layout, literal_string=True)
#         self.Title = title
#         self.Width = width
#         self.Height = height

#         self._context = context
#         self.response = None
#         self.PreviewKeyDown += self.handle_input_key

#         self._setup(**kwargs)

#     def _setup(self, **kwargs):
#         pass

#     def handle_input_key(self, sender, args):
#         if args.Key == framework.Windows.Input.Key.Escape:
#             self.Close()

#     @classmethod
#     def show(cls, context,
#              title='User Input', width=300, height=400, **kwargs):
#         dlg = cls(context, title, width, height, **kwargs)
#         dlg.ShowDialog()
#         return dlg.response

# class SelectFromCheckBoxes(TemplateUserInputWindow):
#     layout = """
#     <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
#             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
#             ShowInTaskbar="False" ResizeMode="CanResize"
#             WindowStartupLocation="CenterScreen"
#             HorizontalContentAlignment="Center">
#             <Window.Resources>
#                 <Style x:Key="ClearButton" TargetType="Button">
#                     <Setter Property="Background" Value="White"/>
#                 </Style>
#             </Window.Resources>
#             <DockPanel Margin="10">
#                 <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
#                     <TextBlock FontSize="14" Margin="0,3,10,0" \
#                                DockPanel.Dock="Left">Filter:</TextBlock>
#                     <StackPanel>
#                         <TextBox x:Name="search_tb" Height="25px"
#                                  TextChanged="search_txt_changed"/>
#                         <Button Style="{StaticResource ClearButton}"
#                                 HorizontalAlignment="Right"
#                                 x:Name="clrsearch_b" Content="x"
#                                 Margin="0,-25,5,0" Padding="0,-4,0,0"
#                                 Click="clear_search"
#                                 Width="15px" Height="15px"/>
#                     </StackPanel>
#                 </DockPanel>
#                 <StackPanel DockPanel.Dock="Bottom">
#                     <Grid>
#                         <Grid.RowDefinitions>
#                             <RowDefinition Height="Auto" />
#                         </Grid.RowDefinitions>
#                         <Grid.ColumnDefinitions>
#                             <ColumnDefinition Width="*" />
#                             <ColumnDefinition Width="*" />
#                             <ColumnDefinition Width="*" />
#                         </Grid.ColumnDefinitions>
#                         <Button x:Name="checkall_b"
#                                 Grid.Column="0" Grid.Row="0"
#                                 Content="Check" Click="check_all"
#                                 Margin="0,10,3,0"/>
#                         <Button x:Name="uncheckall_b"
#                                 Grid.Column="1" Grid.Row="0"
#                                 Content="Uncheck" Click="uncheck_all"
#                                 Margin="3,10,3,0"/>
#                         <Button x:Name="toggleall_b"
#                                 Grid.Column="2" Grid.Row="0"
#                                 Content="Toggle" Click="toggle_all"
#                                 Margin="3,10,0,0"/>
#                     </Grid>
#                     <Button x:Name="select_b" Content=""
#                             Click="button_select" Margin="0,10,0,0"/>
#                 </StackPanel>
#                 <ListView x:Name="list_lb">
#                     <ListView.ItemTemplate>
#                          <DataTemplate>
#                            <StackPanel>
#                              <CheckBox Content="{Binding name}"
#                                        IsChecked="{Binding state}"/>
#                            </StackPanel>
#                          </DataTemplate>
#                    </ListView.ItemTemplate>
#                 </ListView>
#             </DockPanel>
#     </Window>
#     """
#     def _setup(self, **kwargs):
#         self.hide_element(self.clrsearch_b)
#         self.clear_search(None, None)
#         self.search_tb.Focus()

#         button_name = kwargs.get('button_name', None)
#         if button_name:
#             self.select_b.Content = button_name

#         self._list_options()

#     def _list_options(self, checkbox_filter=None):
#         if checkbox_filter:
#             self.checkall_b.Content = 'Check'
#             self.uncheckall_b.Content = 'Uncheck'
#             self.toggleall_b.Content = 'Toggle'
#             checkbox_filter = checkbox_filter.lower()
#             self.list_lb.ItemsSource = \
#                 [checkbox for checkbox in self._context
#                  if checkbox_filter in checkbox.name.lower()]
#         else:
#             self.checkall_b.Content = 'Check All'
#             self.uncheckall_b.Content = 'Uncheck All'
#             self.toggleall_b.Content = 'Toggle All'
#             self.list_lb.ItemsSource = self._context

#     def _set_states(self, state=True, flip=False):
#         current_list = self.list_lb.ItemsSource
#         for checkbox in current_list:
#             if flip:
#                 checkbox.state = not checkbox.state
#             else:
#                 checkbox.state = state

#         # push list view to redraw
#         self.list_lb.ItemsSource = None
#         self.list_lb.ItemsSource = current_list

#     def toggle_all(self, sender, args):
#         self._set_states(flip=True)

#     def check_all(self, sender, args):
#         self._set_states(state=True)

#     def uncheck_all(self, sender, args):
#         self._set_states(state=False)

#     def button_select(self, sender, args):
#         self.response = self._context
#         self.Close()

#     def search_txt_changed(self, sender, args):
#         if self.search_tb.Text == '':
#             self.hide_element(self.clrsearch_b)
#         else:
#             self.show_element(self.clrsearch_b)

#         self._list_options(checkbox_filter=self.search_tb.Text)

#     def clear_search(self, sender, args):
#         self.search_tb.Text = ' '
#         self.search_tb.Clear()
#         self.list_lb.ItemsSource = self._context

# class CheckBoxOption:
# 	def __init__(self, name, default_state=False):
# 		self.name = name
# 		self.state = default_state

# 	def __nonzero__(self):
# 		return self.state

# 	def __bool__(self):
# 		return self.state

# value = TextInput('Value', default="Parameter value")
# ops = ['option1', 'option2', 'option3', 'option4']
# res4 = forms.CommandSwitchWindow.show(ops, message='Select Option')
# print(res4)
# # # # # FORM CHELOU






# # # # # TURN OFF A WORKSET IN ALL VIEWS
# from Autodesk.Revit.DB import *

# activeWorksetId = doc.GetWorksetTable().GetActiveWorksetId()
# print(doc.ActiveView.GetWorksetVisibility(activeWorksetId))

# view_collector = FilteredElementCollector(doc).OfClass(View)

# t = Transaction(doc, 'Change workset visibility in all views')
# t.Start()
# for v in view_collector:
# 	if (v.ViewType == ViewType.ThreeD) or (v.ViewType == ViewType.Section) or (v.ViewType == ViewType.FloorPlan) or (v.ViewType == ViewType.CeilingPlan):
# 		print(v.Name)
# 		v.SetWorksetVisibility(activeWorksetId, WorksetVisibility.UseGlobalSetting)
# t.Commit()
# # # # # TURN OFF A WORKSET IN ALL VIEWS








# # # # # CHANGE IN-PLACE FAMILY CATEGORY
# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# e = get_selected_elements(doc)[0]
# print(e)
# print(e.Symbol.Family.IsInPlace)
# family = e.Symbol.Family
# familyCat = family.FamilyCategory
# print(familyCat.Name)
# bic = BuiltInCategory.OST_Ceilings
# cat = Category.GetCategory(doc, bic)
# print(cat.Name)
# famDoc = doc.EditFamily(family)
# # family.FamilyCategory = cat
# # # # # CHANGE IN-PLACE FAMILY CATEGORY








# # if doc.ActiveView.GetType().ToString() == "Autodesk.Revit.DB.ViewSheet":
# # 	viewId_list = doc.ActiveView.GetAllPlacedViews()
# # 	for viewId in viewId_list:
# # 		print(viewId)
# # 		view = doc.GetElement(viewId)
# # 		print(view)
# # 		for i in view.GetExternalFileReference():
# # 			print(i)
# # else:
# # 	print("The active view is not a view sheet!")

# dwgSettingsFilter = ElementClassFilter(ExportDWGSettings)
# settings = FilteredElementCollector(doc)
# settings = settings.WherePasses(dwgSettingsFilter)

# xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
# # worksheet = xlApp.Worksheets[1]
# worksheet = xlApp.ActiveSheet

# i = 0
# for setting in settings:
# 	# if setting.Name == "Export NTW_TESTbis":
# 	if setting.Name == "Export CITY":
# 		dwgOptions = setting.GetDWGExportOptions()
# 		layerTable = dwgOptions.GetExportLayerTable()
# 		for layerItem in layerTable:
# 			cat = layerItem.Key
# 			layerInfo = layerItem.Value
# 			categoryType = layerInfo.CategoryType
# 			catName = cat.CategoryName
# 			subCatName = cat.SubCategoryName
# 			layerName = layerInfo.LayerName
# 			cutLayerName = layerInfo.CutLayerName
# 			layerModifiers = layerInfo.GetLayerModifiers()
# 			layerModifier = ""
# 			cutLayerModifiers = layerInfo.GetCutLayerModifiers()
# 			cutLayerModifier = ""

# 			if (".dwg" not in catName) and ((categoryType.ToString() == "Model") or (categoryType.ToString() == "Annotation")) :
# 				i = i + 1
# 				print(dir(layerItem))
# 				break
# 				print("--")

# # k = 0
# # l = 0
# # for cat in doc.Settings.Categories:
# # 	catName = cat.Name
# # 	categoryType = cat.CategoryType
# # 	if (".dwg" not in catName) and ((categoryType.ToString() == "Model") or (categoryType.ToString() == "Annotation")) :
# # 		print(cat.SubCategories.Size)
# # 		k = k + 1
# # 		for j in cat.SubCategories:
# # 			l = l + 1

# # print(k+l)



# # CHERCHER ZONE REMPLIES!

# # OST_FillPatterns

# # collector = FilteredElementCollector(doc)\
# # 		.OfClass(FillPatternElement)\
# # 		.WhereElementIsNotElementType()

# collector = FilteredElementCollector(doc).OfClass(FilledRegionType)

# # collector = FilteredElementCollector(doc)\
# # 		.OfClass(FillPatternElement)

# # collector = FilteredElementCollector(doc)\
# # 		.OfCategory(BuiltInCategory.OST_FillPatterns)



# for i in collector:
# 	print(i.Name)

# # CHERCHER ZONE REMPLIES!





# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# e = get_selected_elements(doc)

# insulationId = InsulationLiningBase.GetInsulationIds(doc, e.Id)[0]
# e1 = doc.GetElement(insulationId)
# print(e1.LookupParameter("Sous-projet").AsValueString())

# wsparam = e.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
# wsparamId = wsparam.Id


# provider = ParameterValueProvider(wsparamId)
# evaluator = FilterStringEquals()
# wsparamValue = "13_CVP_CVC_Aeraulique"
# rule = FilterStringRule(provider, evaluator, wsparamValue, False)
# wsfilter = ElementParameterFilter(rule)

# collector = FilteredElementCollector(doc, doc.ActiveView.Id).WherePasses(wsfilter).ToElementIds()

# uidoc.Selection.SetElementIds(collector)

# print(wsparam.StorageType)
# print(wsparam.AsInteger())
# print(wsparam.AsValueString())





# # # # # Read Ai NGF RESA
# collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 				.OfCategory(BuiltInCategory.OST_GenericModel)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# for e in collector:
# 	familyName = e.Symbol.FamilyName
# 	if "BPS_RESA" in familyName:
# 		bb = e.get_Geometry(options).GetBoundingBox()
# 		trans = bb.Transform
# 		zMin = trans.OfPoint(bb.Min).Z
# 		number = e.LookupParameter("BPS_Repere")
# 		lot = e.LookupParameter("BPS_Lot")
# 		PH = e.LookupParameter("BPS_En_Plancher")
# 		print(number.AsString() + str(",") + lot.AsString() + str(",") + str(PH.AsValueString()))
# # # # # Read Ai NGF RESA








# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)


# e = get_selected_elements(doc)[0]
# print(e.Host)
# print(e.HostFace)

# collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# for e in collector:
# 	print(e.IsMonitoringLinkElement())
# 	print(e.IsMonitoringLocalElement())
# 	if e.IsMonitoringLinkElement() is True:
# 		print(e.Id)
# 		print("Link")
# 	if e.IsMonitoringLocalElement() is True:
# 		print(e.Id)
# 		print("Local")




def get_selected_elements(doc):
    try:
        # # Revit 2016
        return [doc.GetElement(id)
                for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
    except:
        # # old method
        return list(__revit__.ActiveUIDocument.Selection.Elements)


def get_solid(element):
	solid_list = []
	for i in element.get_Geometry(options):
		if i.ToString() == "Autodesk.Revit.DB.Solid":
			solid_list.append(i)
		elif i.ToString() == "Autodesk.Revit.DB.GeometryInstance":
			for j in i.GetInstanceGeometry():
				if j.ToString() == "Autodesk.Revit.DB.Solid":
					solid_list.append(j)
	return solid_list

def get_intersection(el1, el2):
	bb1 = el1.get_Geometry(options).GetBoundingBox()
	bb2 = el2.get_Geometry(options).GetBoundingBox()

	trans1 = bb1.Transform
	trans2 = bb2.Transform

	min1 = trans1.OfPoint(bb1.Min)
	max1 = trans1.OfPoint(bb1.Max)
	min2 = trans2.OfPoint(bb2.Min)
	max2 = trans2.OfPoint(bb2.Max)

	outline1 = Outline(min1, max1)
	outline2 = Outline(min2, max2)

	solid1_list = get_solid(el1)
	solid2_list = get_solid(el2)

	for i in solid1_list:
		for j in solid2_list:
			try:
				inter = BooleanOperationsUtils.ExecuteBooleanOperation(i, j, BooleanOperationsType.Intersect)
				if inter.Volume != 0:
					interBb = inter.GetBoundingBox()
					interTrans = interBb.Transform
					interPoint = interTrans.OfPoint(interBb.Min)
					break
			except:
				"Oh god!"

	try:
		interPoint
		return interPoint
	except:
		return None


e1 = get_selected_elements(doc)[0]
print(get_solid(e1)[0].Volume)
e2 = get_selected_elements(doc)[1]

print(get_intersection(e1, e2))

# # # # # NOMME LES COUPES
# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)


# e = get_selected_elements(doc)[0]
# cat = e.Category.Name
# Name = e.Name

# if ("_" not in Name) and (cat != "Vues"):
# 	print("Please select a presentation view section")
# 	sys.exit()

# viewType = e.ViewType.ToString()
# namePrefixe = Name.split("_")[0] + "_"

# collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
# 				.OfCategory(BuiltInCategory.OST_Viewers)\
# 				.WhereElementIsNotElementType()\
# 				.ToElements()

# if viewType == "Section":
# 	for i in collector:
# 		print(i.Name)
# else:
# 	print("Please select a presentation view section")

# # num_list = []
# # j = 0
# # t = Transaction(doc, 'Number bubbles')
# # t.Start()
# # for i in collector:
# # 	level_i = i.LookupParameter("Niveau").AsValueString()
# # 	bType_i = i.LookupParameter("BPS - Emetteur").AsString()
# # 	name_i = i.Name

# # 	if (level_i == level) and (name_i == name) and (bType_i == bType):
# # 		j = j + 1
# # 		num_i = i.LookupParameter("BPS - Num"+str("\xe9")+"ro")
# # 		num_i.Set("Rq " + str(j))
# # t.Commit()
# # print(j)

# # # # # NOMME LES COUPES










# # # # # ZOOM SUR ELEMENT SELECTIONNE
# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# e = get_selected_elements(doc)[0]
# # # # # ZOOM SUR ELEMENT SELECTIONNE


















# # # # # Check connection

# def get_selected_elements(doc):
#     try:
#         # # Revit 2016
#         return [doc.GetElement(id)
#                 for id in __revit__.ActiveUIDocument.Selection.GetElementIds()]
#     except:
#         # # old method
#         return list(__revit__.ActiveUIDocument.Selection.Elements)

# def CheckConnection(e1, e2):
# 	if (e1.ToString() == "Autodesk.Revit.DB.FamilyInstance") or (e2.ToString() == "Autodesk.Revit.DB.FamilyInstance"):
# 	# 	try:
# 	# 		connector = e1
# 	# 		connectorList = e2.ConnectorManager.Connectors
# 	# 	except:
# 	# 		connector = e2
# 	# 		connectorList = e1.ConnectorManager.Connectors
# 	# 	print(dir(connector.MEPModel.ConnectorManager.Connectors))
# 	# 	print(connectorList)
# 	# 	for i in connectorList:
# 	# 		print(i)
# 	# 		if i.Id == connector.Id:
# 	# 			return True
# 	# 			break
# 	# 	else :
# 	# 		return False
# 	# else:
# 	# 	return False
# 		try:
# 			connectorList1 = e1.MEPModel.ConnectorManager.Connectors
# 			connectorList2 = e2.ConnectorManager.Connectors
# 		except:
# 			connectorList2 = e2.MEPModel.ConnectorManager.Connectors
# 			connectorList1 = e1.ConnectorManager.Connectors
# 		for c1 in connectorList1:
# 			print(e1)
# 			print(c1)
# 			if c1 in connectorList2:
# 				return True
# 				break
# 			else:
# 				return False
# 		for c2 in connectorList2:
# 			print(e2)
# 			print(c2)
# 			if c2 in connectorList1:
# 				return True
# 				break
# 			else:
# 				pass

# e1 = get_selected_elements(doc)[0]
# e2 = get_selected_elements(doc)[1]

# print(CheckConnection(e1, e2))

# # # # # Check connection








# el = get_selected_elements(doc)[0]
# print(el)
# print(el.Category.Name)
# print(el.GetType())

# collector = FilteredElementCollector(doc).OfClass(FamilySymbol).OfCategory(BuiltInCategory.OST_FilledRegion)

# for i in collector:
# 	print(i.FamilyName)



# # # TRY TO GET ELEMENT FROM LINK MODEL
# sel = uidoc.Selection.PickObjects(Selection.ObjectType.Element, "Please select an object")

# el = get_selected_elements(doc)[0]
# print(el)
# # # TRY TO GET ELEMENT FROM LINK MODEL




# # # TRY TO REFRESH BUBBLES IN VIEW LOOK THE BUILDING CODER
# t = Transaction(doc, 'tg')
# t.Start()
# doc.Regenerate()
# t.Commit()
# uidoc.RefreshActiveView()
# # # TRY TO REFRESH BUBBLES IN VIEW LOOK THE BUILDING CODER