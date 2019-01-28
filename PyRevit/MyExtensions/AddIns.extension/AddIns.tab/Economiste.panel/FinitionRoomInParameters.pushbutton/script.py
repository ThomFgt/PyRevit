"""Remplit les parametres de finition des pieces"""

__title__ = 'Set room\n finishing parameters'

__doc__ = 'Ce programme remplit les parametres de finition des pieces'

import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
options = __revit__.Application.Create.NewGeometryOptions()
SEBoptions = SpatialElementBoundaryOptions()
roomcalculator = SpatialElementGeometryCalculator(doc)
	
def GetElementLayers(element):
	elementType = doc.GetElement(element.GetTypeId())
	compoundStructure = elementType.GetCompoundStructure()

	j = 1
	layers_dict = {}
	for i in compoundStructure.GetLayers():
		try:
			layers_dict[str(i.Function) + "_" + str(j)] = str(doc.GetElement(i.MaterialId).Name)
			j = j + 1
		except:
			"Non structurel"
	
	return layers_dict

td_button = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Cancel | TaskDialogCommonButtons.No

res = TaskDialog.Show("Finition des pieces","Les pieces sont-elles definies par rapport a la couche de finition?",td_button)

if res == TaskDialogResult.Yes:
	
	t = Transaction(doc, 'Setting rooms finitions')
	t.Start()

	room_collector = FilteredElementCollector(doc)\
		  .OfCategory(BuiltInCategory.OST_Rooms)\
		  .WhereElementIsNotElementType()\
		  .ToElements()
		  
	for room in room_collector:
		
		try:
			roomelement = roomcalculator.CalculateSpatialElementGeometry(room)	
			roomsolid = roomelement.GetGeometry()
			roomnumber = room.Number
			roomname = room.LookupParameter("Nom").AsString()
			roomlevel = room.Level.Name
			roomGroupId = room.GroupId
			if str(roomGroupId) == "-1":
				"Room not in group"
			else:
				group = doc.GetElement(room.GroupId)
				group.UngroupMembers()
		
			bb = room.get_Geometry(options).GetBoundingBox()
			
			outline = Outline(bb.Min, bb.Max)

			filter1 = BoundingBoxIntersectsFilter(outline)
			filter2 = BoundingBoxIsInsideFilter(outline)
			filter = LogicalOrFilter(filter1, filter2)

			IntersectRoomFilter = ElementIntersectsSolidFilter(roomsolid)
			
			part_collector = FilteredElementCollector(doc)\
				  .WherePasses(filter)\
				  .OfCategory(BuiltInCategory.OST_Parts)\
				  .WhereElementIsNotElementType()\
				  .ToElements()
				  
			part_test_collector = FilteredElementCollector(doc)\
				  .WherePasses(IntersectRoomFilter)\
				  .OfCategory(BuiltInCategory.OST_Parts)\
				  .WhereElementIsNotElementType()\
				  .ToElementIds()
			
			ceiling_collector = FilteredElementCollector(doc)\
				  .WherePasses(filter)\
				  .OfCategory(BuiltInCategory.OST_Ceilings)\
				  .WhereElementIsNotElementType()\
				  .ToElements()

			ceiling_test_collector = FilteredElementCollector(doc)\
				  .WherePasses(IntersectRoomFilter)\
				  .OfCategory(BuiltInCategory.OST_Ceilings)\
				  .WhereElementIsNotElementType()\
				  .ToElementIds()
				  
			floor_collector = FilteredElementCollector(doc)\
				  .WherePasses(filter)\
				  .OfCategory(BuiltInCategory.OST_Floors)\
				  .WhereElementIsNotElementType()\
				  .ToElements()
				  
			floor_test_collector = FilteredElementCollector(doc)\
				  .WherePasses(IntersectRoomFilter)\
				  .OfCategory(BuiltInCategory.OST_Floors)\
				  .WhereElementIsNotElementType()\
				  .ToElementIds()
				  
			part_list = []
			for part in part_collector:
				try:
					if (part.Id in part_test_collector) and (part.LookupParameter("Construction").AsString() == "Finition"):
						part_list.append(part)
					else :
						for i in part.get_Geometry(options):
							if (i.GetType().ToString() == "Autodesk.Revit.DB.Solid") and (part.LookupParameter("Construction").AsString() == "Finition"):
								for j in i.Edges:
									spoint = j.AsCurve().GetEndPoint(0)
									epoint = j.AsCurve().GetEndPoint(1)
									if (room.IsPointInRoom(spoint) is True) or (room.IsPointInRoom(epoint) is True):
										part_list.append(part)
										break
				except:
					"Construction not a part parameter"

			fpart = ""
			fpart_list = []
			for i in part_list:
				try:
					if (i.OriginalCategoryId.IntegerValue == -2000011) and (i.LookupParameter("Mat"+"\xe9"+"riau").AsValueString() not in fpart_list) and (i.LookupParameter("Mat"+"\xe9"+"riau").AsValueString()!="<Par cat"+"\xe9"+"gorie>"):
						fpart_list.append(i.LookupParameter("Mat"+"\xe9"+"riau").AsValueString())
				except:
					"Mat not a part parameter"
			
			for x in sorted(fpart_list):
				fpart = fpart + x + " / "
			if len(fpart)>0:
				room.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).Set(fpart[:len(fpart)-3])
			else:
				room.get_Parameter(BuiltInParameter.ROOM_FINISH_WALL).Set("")			
					
			ceiling_list = []
			for ceiling in ceiling_collector:
				if ceiling.Id in ceiling_test_collector:
					ceiling_list.append(ceiling)
				else :
					for i in ceiling.get_Geometry(options):
						if i.GetType().ToString() == "Autodesk.Revit.DB.Solid":
							for j in i.Edges:
								spoint = j.AsCurve().GetEndPoint(0)
								epoint = j.AsCurve().GetEndPoint(1)
								if (room.IsPointInRoom(spoint) is True) or (room.IsPointInRoom(epoint) is True):
									ceiling_list.append(ceiling)
									break
			
			fceiling = ""
			fceiling_list = []
			for j in ceiling_list:
				try:
					for m in GetElementLayers(j):
						if ("Finish2" in str(m)) and (str(GetElementLayers(j)[m]) not in fceiling_list):
							fceiling_list.append(str(GetElementLayers(j)[m]))
				except:
					"Not ceiling element"
			for y in sorted(fceiling_list):
				fceiling = fceiling + y + " / "
			if len(fceiling)>0:
				room.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING).Set(fceiling[:len(fceiling)-3])
			else:
				room.get_Parameter(BuiltInParameter.ROOM_FINISH_CEILING).Set("")
					
			floor_list = []
			for floor in floor_collector:
				if floor.Id in floor_test_collector:
					floor_list.append(floor)
				else :
					for i in floor.get_Geometry(options):
						if i.GetType().ToString() == "Autodesk.Revit.DB.Solid":
							for j in i.Edges:
								spoint = j.AsCurve().GetEndPoint(0)
								epoint = j.AsCurve().GetEndPoint(1)
								if (room.IsPointInRoom(spoint) is True) or (room.IsPointInRoom(epoint) is True):
									floor_list.append(floor)
									break
			
			ffloor = ""
			ffloor_list = []
			for k in floor_list:
				try:
					for n in GetElementLayers(k):
						if ("Finish2" in str(n)) and (str(GetElementLayers(k)[n]) not in ffloor_list):
							ffloor_list.append(str(GetElementLayers(k)[n]))
				except:
					"Not floor element"
			for z in sorted(ffloor_list):
				ffloor = ffloor + z + " / "
			if len(ffloor)>0:
				room.get_Parameter(BuiltInParameter.ROOM_FINISH_FLOOR).Set(ffloor[:len(ffloor)-3])
			else:
				room.get_Parameter(BuiltInParameter.ROOM_FINISH_FLOOR).Set("")
				
			print(str(roomnumber) + " - " + str(roomlevel) + str(" - ") + str(roomname) + str(" : ") + " OK!\n")
			
		except:
			print(str(roomnumber) + " - " + str(roomlevel) + str(" - ") + str(roomname) + str(" : ") + "Room not closed or outside the building!\n")
	
	t.Commit()
	
	td_button2 = TaskDialogCommonButtons.Ok

	res2 = TaskDialog.Show("Finition des pieces","Remplissage termine",td_button2)
	
else:
	"A plus tard!"