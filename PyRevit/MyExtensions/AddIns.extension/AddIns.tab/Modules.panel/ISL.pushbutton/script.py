# import clr
from pyrevit.framework import List
from pyrevit import forms
from pyrevit import revit, DB
# import math
# clr.AddReference('RevitAPI') 
# clr.AddReference('RevitAPIUI') 
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *


__doc__ = 'Isolates specific elements in current view and '\
          'put the view in isolate element mode.'

element_cats = {'Pipes': DB.BuiltInCategory.OST_PipeCurves,
                'Ducts': DB.BuiltInCategory.OST_DuctCurves,
				'Conduit': DB.BuiltInCategory.OST_Conduit,
                'Cable trays': DB.BuiltInCategory.OST_CableTray,
				'Floors': DB.BuiltInCategory.OST_Floors,
				'Beams': DB.BuiltInCategory.OST_StructuralFraming,
				'Revision clouds' : DB.BuiltInCategory.OST_RevisionClouds,
				'Vues' : DB.BuiltInCategory.OST_Viewers,
				'Generic Models' : DB.BuiltInCategory.OST_GenericModel,
				'Parts' : DB.BuiltInCategory.OST_Parts,
				'Ceilings' : DB.BuiltInCategory.OST_Ceilings,
				'Doors' : DB.BuiltInCategory.OST_Doors,
				'Walls' : DB.BuiltInCategory.OST_Walls}


selected_switch = \
    forms.CommandSwitchWindow.show(
        sorted(element_cats.keys()),
        message='Temporarily isolate elements of type:'
        )


if selected_switch:
    curview = revit.activeview

    if selected_switch == 'Room Tags':
        roomtags = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .OfCategory(DB.BuiltInCategory.OST_RoomTags)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        rooms = DB.FilteredElementCollector(revit.doc, curview.Id)\
                  .OfCategory(DB.BuiltInCategory.OST_Rooms)\
                  .WhereElementIsNotElementType()\
                  .ToElementIds()

        allelements = []
        allelements.extend(rooms)
        allelements.extend(roomtags)
        element_to_isolate = List[DB.ElementId](allelements)

    elif selected_switch == 'Model Groups':
        elements = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        modelgroups = []
        expanded = []
        for elid in elements:
            el = revit.doc.GetElement(elid)
            if isinstance(el, DB.Group) and not el.ViewSpecific:
                modelgroups.append(elid)
                members = el.GetMemberIds()
                expanded.extend(list(members))

        expanded.extend(modelgroups)
        element_to_isolate = List[DB.ElementId](expanded)

    elif selected_switch == 'Painted Elements':
        set = []

        elements = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        for elId in elements:
            el = revit.doc.GetElement(elId)
            if len(list(el.GetMaterialIds(True))) > 0:
                set.append(elId)
            elif isinstance(el, DB.Wall) and el.IsStackedWall:
                memberWalls = el.GetStackedWallMemberIds()
                for mwid in memberWalls:
                    mw = revit.doc.GetElement(mwid)
                    if len(list(mw.GetMaterialIds(True))) > 0:
                        set.append(elId)
        element_to_isolate = List[DB.ElementId](set)

    elif selected_switch == 'Model Elements':
        elements = DB.FilteredElementCollector(revit.doc, curview.Id)\
                     .WhereElementIsNotElementType()\
                     .ToElementIds()

        element_to_isolate = []
        for elid in elements:
            el = revit.doc.GetElement(elid)
            if not el.ViewSpecific:  # and not isinstance(el, DB.Dimension):
                element_to_isolate.append(elid)

        element_to_isolate = List[DB.ElementId](element_to_isolate)

    else:
        element_to_isolate = \
            DB.FilteredElementCollector(revit.doc, curview.Id)\
              .OfCategory(element_cats[selected_switch]) \
              .WhereElementIsNotElementType()\
              .ToElementIds()

    # now that list of elements is ready, let's isolate them in the active view
    with revit.Transaction('Isolate {}'.format(selected_switch)):
        curview.IsolateElementsTemporary(element_to_isolate)
