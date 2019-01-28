"""Etiquetage de poutres CANOPY!"""

# -*- coding: utf-8 -*-
e_a = str("\xe9")
a_a = str("\xe0")

__title__ = 'Beam tagging\nCANOPY'

__doc__ = 'Ce programme remplit les arases inferieures mini et maxi des poutres '\
      '(parametres AI_Min et AI_Max) de la vue active du projet CANOPY'

# from pyrevit import revit, DB, UI
# from pyrevit import script
# from pyrevit import forms

import clr
import math
from pyrevit import forms
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import * 

doc = __revit__.ActiveUIDocument.Document

options = __revit__.Application.Create.NewGeometryOptions()

BP_collector = FilteredElementCollector(doc)\
          .OfCategory(BuiltInCategory.OST_ProjectBasePoint)\
          .WhereElementIsNotElementType()\
          .ToElements()

beam_collector = FilteredElementCollector(doc,doc.ActiveView.Id)\
          .OfCategory(BuiltInCategory.OST_StructuralFraming)\
          .WhereElementIsNotElementType()\
          .ToElements()
      
    
td_button = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Cancel | TaskDialogCommonButtons.No

res = TaskDialog.Show("Etiquetage de poutres","Voulez-vous lancer l'"+e_a+"tiquetage des poutres dans la vue active?",td_button)

if res == TaskDialogResult.Yes:

  for BP in BP_collector:
    zBP = BP.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()/3.2808399
    
  tg = TransactionGroup(doc, 'Tag all beams in active view')

  tg.Start()

  for beam in beam_collector:
    # if beam.Id not in TaggegBeamsList:
      try:
        print("Nom de la poutre : " + beam.Name)
        print("ID : " + str(beam.Id))
        print("Niveau de r"+e_a+"f"+e_a+"rence : " + str(beam.LookupParameter('Niveau de r'+e_a+'f'+e_a+'rence').AsValueString()))
        start = beam.Location.Curve.GetEndPoint(0)
        print(start)
        end = beam.Location.Curve.GetEndPoint(1)
        print(end)
        z_min = zBP + ((beam.get_Geometry(options).GetBoundingBox().Min).Z)/3.2808399
        print("AI_Inf : " + str(z_min) + "m")
        
        # Si la poutre est en pente
        if beam.LookupParameter('El'+e_a+'vation '+a_a+' la base').AsDouble() == 0:
          print("EN PENTE")
          t = Transaction(doc, 'Tag beam')
          t.Start()
          try:
            delta = abs(float(beam.LookupParameter("D"+e_a+"calage du niveau de d"+e_a+"part").AsValueString())-float(beam.LookupParameter("D"+e_a+"calage du niveau d'arriv"+e_a+"e").AsValueString()))\
                  +abs(float(beam.LookupParameter("Valeur de d"+e_a+"calage de l'extr"+e_a+"mit"+e_a+" Z").AsValueString())-float(beam.LookupParameter("Valeur de d"+e_a+"calage Z de d"+e_a+"but").AsValueString()))
          except:
            delta = abs(float(beam.LookupParameter("D"+e_a+"calage du niveau de d"+e_a+"part").AsValueString())-float(beam.LookupParameter("D"+e_a+"calage du niveau d'arriv"+e_a+"e").AsValueString()))
          z_max = z_min + delta
          print("AI_Max : " + str(z_max) + "m")
          beam.LookupParameter('AI_Min').Set(" ")
          beam.LookupParameter('AI_Min').Set(str(round(z_min,2)))
          beam.LookupParameter('AI_Max').Set(" ")
          beam.LookupParameter('AI_Max').Set(str(round(z_max,2)))
          cen=XYZ((start.X+end.X)/2,(start.Y+end.Y)/2,(z_min+z_max)/2)
          print(cen)
          print("\n")
          beam_tag = doc.Create.NewTag(doc.ActiveView,beam,False,TagMode.TM_ADDBY_CATEGORY,TagOrientation.Horizontal,cen)
          t.Commit()
        # Si la poutre est horizontale
        else :
          print("HORIZONTALE")
          t = Transaction(doc, 'Tag beam')
          t.Start()
          beam.LookupParameter('AI_Min').Set(" ")
          beam.LookupParameter('AI_Min').Set(str(round(z_min,2)))
          beam.LookupParameter('AI_Max').Set(" ")
          cen=XYZ((start.X+end.X)/2,(start.Y+end.Y)/2,z_min)
          print(cen)
          print("\n")
          beam_tag = doc.Create.NewTag(doc.ActiveView,beam,False,TagMode.TM_ADDBY_CATEGORY,TagOrientation.Horizontal,cen)
          t.Commit()
      
      except:
        print(" ")
        
  tg.Assimilate()
  
  td_button2 = TaskDialogCommonButtons.Ok

  res2 = TaskDialog.Show("Mise au propre","Veuillez checker la non superposition des "+e_a+"tiquettes pour un r"+e_a+"sultat plus lisible!",td_button2)
        
else:
  print("Une autre fois peut-etre...")
  


