__title__ = 'Fill Ai_NGF paramters\nin active view'

__doc__ = 'This program fills Ai_NGF parameters in active view'

import clr
import System
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Drawing import *
from System.Windows.Forms import *
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()

iList = List[BuiltInCategory]() 
iList.Add(BuiltInCategory.OST_CableTray)
iList.Add(BuiltInCategory.OST_DuctCurves)
iList.Add(BuiltInCategory.OST_FlexDuctCurves)
iList.Add(BuiltInCategory.OST_PipeCurves)
iList.Add(BuiltInCategory.OST_PipeCurves)

catFilter = ElementMulticategoryFilter(iList)

collector = FilteredElementCollector(doc, doc.ActiveView.Id)\
				.WherePasses(catFilter)\
				.WhereElementIsNotElementType()\
				.ToElements()

t = Transaction(doc, 'Fill Ai_NGF parameters')
t.Start()

for e in collector:
	bb = e.get_Geometry(options).GetBoundingBox()
	trans = bb.Transform
	minPoint = trans.OfPoint(bb.Min)
	zMin = UnitUtils.ConvertFromInternalUnits(minPoint.Z, DisplayUnitType.DUT_METERS)
	param = e.LookupParameter("Arase inf"+"\xe9"+"rieure NGF")
	param.Set(minPoint.Z)

t.Commit()