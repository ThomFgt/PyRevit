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
from pyrevit import framework
from pyrevit.framework import Interop
from pyrevit.api import AdWindows
from pyrevit.framework import wpf, Forms, Controls, Media

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

class WPFWindow(framework.Windows.Window):
    def __init__(self, xaml_source, literal_string=False):
        # self.Parent = self
        wih = Interop.WindowInteropHelper(self)
        wih.Owner = AdWindows.ComponentManager.ApplicationWindow

        if not literal_string:
            if not op.exists(xaml_source):
                wpf.LoadComponent(self,
                                  os.path.join(EXEC_PARAMS.command_path,
                                               xaml_source)
                                  )
            else:
                wpf.LoadComponent(self, xaml_source)
        else:
            wpf.LoadComponent(self, framework.StringReader(xaml_source))

        #2c3e50
        self.Resources['pyRevitDarkColor'] = \
            Media.Color.FromArgb(0xFF, 0x2c, 0x3e, 0x50)

        #23303d
        self.Resources['pyRevitDarkerDarkColor'] = \
            Media.Color.FromArgb(0xFF, 0x23, 0x30, 0x3d)

        #ffffff
        self.Resources['pyRevitButtonColor'] = \
            Media.Color.FromArgb(0xFF, 0xff, 0xff, 0xff)

        #f39c12
        self.Resources['pyRevitAccentColor'] = \
            Media.Color.FromArgb(0xFF, 0xf3, 0x9c, 0x12)

        self.Resources['pyRevitDarkBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitDarkColor'])
        self.Resources['pyRevitAccentBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitAccentColor'])

        self.Resources['pyRevitDarkerDarkBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitDarkerDarkColor'])

        self.Resources['pyRevitButtonForgroundBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitButtonColor'])

    def show(self, modal=False):
        if modal:
            return self.ShowDialog()
        self.Show()

    def show_dialog(self):
        return self.ShowDialog()

    def set_image_source(self, element_name, image_file):
        wpfel = getattr(self, element_name)
        if not op.exists(image_file):
            # noinspection PyUnresolvedReferences
            wpfel.Source = \
                framework.Imaging.BitmapImage(
                    framework.Uri(os.path.join(EXEC_PARAMS.command_path,
                                               image_file))
                    )
        else:
            wpfel.Source = \
                framework.Imaging.BitmapImage(framework.Uri(image_file))

    @staticmethod
    def hide_element(*wpf_elements):
        for wpfel in wpf_elements:
            wpfel.Visibility = framework.Windows.Visibility.Collapsed

    @staticmethod
    def show_element(*wpf_elements):
        for wpfel in wpf_elements:
            wpfel.Visibility = framework.Windows.Visibility.Visible

    @staticmethod
    def toggle_element(*wpf_elements):
        for wpfel in wpf_elements:
            if wpfel.Visibility == framework.Windows.Visibility.Visible:
                self.hide_element(wpfel)
            elif wpfel.Visibility == framework.Windows.Visibility.Collapsed:
                self.show_element(wpfel)

class TemplateUserInputWindow(WPFWindow):
    layout = """
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            ShowInTaskbar="False" ResizeMode="NoResize"
            WindowStartupLocation="CenterScreen"
            HorizontalContentAlignment="Center">
    </Window>
    """
    def __init__(self, context, title, width, height, **kwargs):
        WPFWindow.__init__(self, self.layout, literal_string=True)
        self.Title = title
        self.Width = width
        self.Height = height

        self._context = context
        self.response = None
        self.PreviewKeyDown += self.handle_input_key

        self._setup(**kwargs)

    def _setup(self, **kwargs):
        pass

    def handle_input_key(self, sender, args):
        if args.Key == framework.Windows.Input.Key.Escape:
            self.Close()

    @classmethod
    def show(cls, context,
             title='User Input', width=300, height=400, **kwargs):
        dlg = cls(context, title, width, height, **kwargs)
        dlg.ShowDialog()
        return dlg.response

class SelectFromCheckBoxes(TemplateUserInputWindow):
    layout = """
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            ShowInTaskbar="False" ResizeMode="CanResize"
            WindowStartupLocation="CenterScreen"
            HorizontalContentAlignment="Center">
            <Window.Resources>
                <Style x:Key="ClearButton" TargetType="Button">
                    <Setter Property="Background" Value="White"/>
                </Style>
            </Window.Resources>
            <DockPanel Margin="10">
                <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
                    <TextBlock FontSize="14" Margin="0,3,10,0" \
                               DockPanel.Dock="Left">Filter:</TextBlock>
                    <StackPanel>
                        <TextBox x:Name="search_tb" Height="25px"
                                 TextChanged="search_txt_changed"/>
                        <Button Style="{StaticResource ClearButton}"
                                HorizontalAlignment="Right"
                                x:Name="clrsearch_b" Content="x"
                                Margin="0,-25,5,0" Padding="0,-4,0,0"
                                Click="clear_search"
                                Width="15px" Height="15px"/>
                    </StackPanel>
                </DockPanel>
                <StackPanel DockPanel.Dock="Bottom">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto" />
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                            <ColumnDefinition Width="*" />
                        </Grid.ColumnDefinitions>
                        <Button x:Name="checkall_b"
                                Grid.Column="0" Grid.Row="0"
                                Content="Check" Click="check_all"
                                Margin="0,10,3,0"/>
                        <Button x:Name="uncheckall_b"
                                Grid.Column="1" Grid.Row="0"
                                Content="Uncheck" Click="uncheck_all"
                                Margin="3,10,3,0"/>
                        <Button x:Name="toggleall_b"
                                Grid.Column="2" Grid.Row="0"
                                Content="Toggle" Click="toggle_all"
                                Margin="3,10,0,0"/>
                    </Grid>
                    <Button x:Name="select_b" Content=""
                            Click="button_select" Margin="0,10,0,0"/>
                </StackPanel>
                <ListView x:Name="list_lb">
                    <ListView.ItemTemplate>
                         <DataTemplate>
                           <StackPanel>
                             <CheckBox Content="{Binding name}"
                                       IsChecked="{Binding state}"/>
                           </StackPanel>
                         </DataTemplate>
                   </ListView.ItemTemplate>
                </ListView>
            </DockPanel>
    </Window>
    """
    def _setup(self, **kwargs):
        self.hide_element(self.clrsearch_b)
        self.clear_search(None, None)
        self.search_tb.Focus()

        button_name = kwargs.get('button_name', None)
        if button_name:
            self.select_b.Content = button_name

        self._list_options()

    def _list_options(self, checkbox_filter=None):
        if checkbox_filter:
            self.checkall_b.Content = 'Check'
            self.uncheckall_b.Content = 'Uncheck'
            self.toggleall_b.Content = 'Toggle'
            checkbox_filter = checkbox_filter.lower()
            self.list_lb.ItemsSource = \
                [checkbox for checkbox in self._context
                 if checkbox_filter in checkbox.name.lower()]
        else:
            self.checkall_b.Content = 'Check All'
            self.uncheckall_b.Content = 'Uncheck All'
            self.toggleall_b.Content = 'Toggle All'
            self.list_lb.ItemsSource = self._context

    def _set_states(self, state=True, flip=False):
        current_list = self.list_lb.ItemsSource
        for checkbox in current_list:
            if flip:
                checkbox.state = not checkbox.state
            else:
                checkbox.state = state

        # push list view to redraw
        self.list_lb.ItemsSource = None
        self.list_lb.ItemsSource = current_list

    def toggle_all(self, sender, args):
        self._set_states(flip=True)

    def check_all(self, sender, args):
        self._set_states(state=True)

    def uncheck_all(self, sender, args):
        self._set_states(state=False)

    def button_select(self, sender, args):
        self.response = self._context
        self.Close()

    def search_txt_changed(self, sender, args):
        if self.search_tb.Text == '':
            self.hide_element(self.clrsearch_b)
        else:
            self.show_element(self.clrsearch_b)

        self._list_options(checkbox_filter=self.search_tb.Text)

    def clear_search(self, sender, args):
        self.search_tb.Text = ' '
        self.search_tb.Clear()
        self.list_lb.ItemsSource = self._context

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
									title = "Select in :",
									multiselect = False,
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

	all_checkboxes = SelectFromCheckBoxes.show(options)

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
										button_name = "OK")

	chosenParam = param_dir[res3[0]]
	try:
		chosenParamId = ElementId(chosenParam)
	except:
		chosenParamId = chosenParam.Id
	categoryFilter = LogicalOrFilter(elementFilter)
	provider = ParameterValueProvider(chosenParamId)
	
	element_collector = collector\
		.WhereElementIsNotElementType()\
		.WherePasses(categoryFilter)
		
	e0 = element_collector.FirstElement()
	
	try:
		storageType = e0.get_Parameter(chosenParam).StorageType
	except:
		storageType = e0.get_Parameter(chosenParam.GuidValue).StorageType
		
	if res3 != None:
		if str(storageType) == "Double":
			ops = ['>=', '=', '<=']
			res4 = forms.CommandSwitchWindow.show(ops, message='Select the filter rule')
		else:
			ops = ['Equals', 'Contains']
			res4 = forms.CommandSwitchWindow.show(ops, message='Select the filter rule')
		
		value = TextInput('Value', default="Parameter value")
		
		if (str(storageType) == "Integer") or (str(storageType) == "String") or (str(storageType) == "ElementId"):
			if res4 == "Equals":
				evaluator = FilterStringEquals()
				ruleValue = value
				rule = FilterStringRule(provider, evaluator, ruleValue, False)
				elementParamFilter = ElementParameterFilter(rule)
			elif res4 == "Contains":
				evaluator = FilterStringContains()
				ruleValue = value
				rule = FilterStringRule(provider, evaluator, ruleValue, False)
				elementParamFilter = ElementParameterFilter(rule)
		# elif str(storageType) == "String":
			# evaluator = FilterStringEquals()
			# ruleValue = value
			# rule = FilterStringRule(provider, evaluator, ruleValue, False)
			# elementParamFilter = ElementParameterFilter(rule)
		# elif str(storageType) == "ElementId":
			# evaluator = FilterStringEquals()
			# ruleValue = value
			# rule = FilterStringRule(provider, evaluator, ruleValue, False)
			# elementParamFilter = ElementParameterFilter(rule)
		elif str(storageType) == "Double":
			if res4 == "=":
				evaluator = FilterNumericEquals()
				d = str(value)[:-1].find(".")
				# ruleValue = round(UnitUtils.ConvertFromInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
				# ruleValue = round(float(value), d)
				ruleValue = round(UnitUtils.ConvertToInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
				rule = FilterDoubleRule(provider, evaluator, ruleValue, 10**(-d))
				elementParamFilter = ElementParameterFilter(rule)
			elif res4 == ">=":
				evaluator = FilterNumericGreaterOrEqual()
				d = str(value)[:-1].find(".")
				# ruleValue = round(UnitUtils.ConvertFromInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
				# ruleValue = round(float(value), d)
				ruleValue = round(UnitUtils.ConvertToInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
				rule = FilterDoubleRule(provider, evaluator, ruleValue, 10**(-d))
				elementParamFilter = ElementParameterFilter(rule)
			elif res4 == "<=":
				evaluator = FilterNumericLessOrEqual()
				d = str(value)[:-1].find(".")
				# ruleValue = round(UnitUtils.ConvertFromInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
				# ruleValue = round(float(value), d)
				ruleValue = round(UnitUtils.ConvertToInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
				rule = FilterDoubleRule(provider, evaluator, ruleValue, 10**(-d))
				elementParamFilter = ElementParameterFilter(rule)
			
			
		element_collector = collector\
			.WhereElementIsNotElementType()\
			.WherePasses(categoryFilter)\
			.WherePasses(elementParamFilter)\
			.ToElementIds()

		uidoc.Selection.SetElementIds(element_collector)
	else:
		"bye"
else:
	"bye"

# import clr
# import math
# clr.AddReference('RevitAPI') 
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from System.Collections.Generic import List
# from pyrevit import forms
# from rpw.ui.forms import TextInput

# doc = __revit__.ActiveUIDocument.Document
# uidoc = __revit__.ActiveUIDocument

# class CheckBoxOption:
	# def __init__(self, name, default_state=False):
		# self.name = name
		# self.state = default_state

	# def __nonzero__(self):
		# return self.state

	# def __bool__(self):
		# return self.state

# view_list = ["La vue active", "Tout le document"]
# res1 = forms.SelectFromList.show(view_list,
									# multiselect = False,
									# name_attr = "Vue",
									# button_name = "OK")

# if res1 != None:
	# if res1[0] == "La vue active":
		# collector = FilteredElementCollector(doc, doc.ActiveView.Id)
	# else:
		# collector = FilteredElementCollector(doc)

	# categories = doc.Settings.Categories
	# bic_dir = {}
	# for bic in BuiltInCategory.GetValues(BuiltInCategory):
		# try:
			# cat = categories.get_Item(bic)
			# bicName = cat.Name
			# bic_dir[bicName] = bic
		# except:
			# "shit"

	# bic_list = bic_dir.keys()
	# options = [CheckBoxOption(j) for j in bic_list]

	# all_checkboxes = forms.SelectFromCheckBoxes.show(options)

	# res2 = []
	# for checkbox in all_checkboxes:
		# if checkbox.state == True:
			# res2.append(checkbox.name)

	# chosenCat_list = [bic_dir[x] for x in res2]

	# icatcollection = List[ElementId]()
	# elementFilter = List[ElementFilter](len(chosenCat_list))
	# for i in chosenCat_list:
		# # icatcollection.Add(i.Id)
		# icatcollection.Add(categories.get_Item(i).Id)
		# elementFilter.Add(ElementCategoryFilter(i))

	# paramFilter = ParameterFilterUtilities.GetFilterableParametersInCommon(doc,icatcollection)

	# bip_dir = {}
	# bipName_dir = {}
	# for bip in BuiltInParameter.GetValues(BuiltInParameter):
		# bipId = str(ElementId(bip).IntegerValue)
		# bip_dir[bipId] = bip
		# try:
			# bipName = LabelUtils.GetLabelFor(bip)
		# except:
			# bipName = bip.ToString()
		# bipName_dir[bipId] = bipName

	# param_dir = {}
	# for paramId in paramFilter:
		# if paramId.IntegerValue < 0:
			# paramName = bipName_dir[str(paramId.IntegerValue)]
			# param = bip_dir[str(paramId.IntegerValue)]
			# param_dir[paramName] = param
		# else:
			# paramName = doc.GetElement(paramId).Name
			# param = doc.GetElement(paramId)
			# param_dir[paramName] = param


	# res3 = forms.SelectFromList.show(param_dir.keys(),
										# multiselect = False,
										# name_attr = "Parametre",
										# button_name = "OK")

										
										
										
										
										
	# chosenParam = param_dir[res3[0]]
	# try:
		# chosenParamId = ElementId(chosenParam)
	# except:
		# chosenParamId = chosenParam.Id
	# categoryFilter = LogicalOrFilter(elementFilter)
	# provider = ParameterValueProvider(chosenParamId)
	
	# element_collector = collector\
		# .WhereElementIsNotElementType()\
		# .WherePasses(categoryFilter)
		
	# e0 = element_collector.FirstElement()
	
	# try:
		# storageType = e0.get_Parameter(chosenParam).StorageType
	# except:
		# storageType = e0.get_Parameter(chosenParam.GuidValue).StorageType
	
	# if res3 != None:
		# value = TextInput('Value', default="Parameter value")
		
		# if str(storageType) == "Integer":
			# evaluator = FilterStringEquals()
			# ruleValue = value
			# rule = FilterStringRule(provider, evaluator, ruleValue, False)
			# elementParamFilter = ElementParameterFilter(rule)
		# elif str(storageType) == "String":
			# evaluator = FilterStringEquals()
			# ruleValue = value
			# rule = FilterStringRule(provider, evaluator, ruleValue, False)
			# elementParamFilter = ElementParameterFilter(rule)
		# elif str(storageType) == "ElementId":
			# evaluator = FilterStringEquals()
			# ruleValue = value
			# rule = FilterStringRule(provider, evaluator, ruleValue, False)
			# elementParamFilter = ElementParameterFilter(rule)
		# elif str(storageType) == "Double":
			# evaluator = FilterNumericEquals()
			# d = str(value)[:-1].find(".")
			# # ruleValue = round(UnitUtils.ConvertFromInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
			# # ruleValue = round(float(value), d)
			# ruleValue = round(UnitUtils.ConvertToInternalUnits(float(value), DisplayUnitType.DUT_METERS), d)
			# rule = FilterDoubleRule(provider, evaluator, ruleValue, 10**(-d))
			# elementParamFilter = ElementParameterFilter(rule)
			
			
		# element_collector = collector\
			# .WhereElementIsNotElementType()\
			# .WherePasses(categoryFilter)\
			# .WherePasses(elementParamFilter)\
			# .ToElementIds()

		# uidoc.Selection.SetElementIds(element_collector)
	# else:
		# "bye"
# else:
	# "bye"