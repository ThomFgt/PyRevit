using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Windows;
using Autodesk.Revit.DB.Architecture;

namespace infoappart
{
    class SurfaceSol
    {


        public static Double Solstructurelparniveau(Document doc, Level level)
        {
            ElementId levelid = level.Id;
            ElementLevelFilter parniveau = new ElementLevelFilter(levelid);
            Dictionary<Level, List<Floor>> Dicosolparniveau = new Dictionary<Level, List<Floor>>();

            FilteredElementCollector floors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().WherePasses(parniveau);
            double SDPniveau = 0.00;

            //int compte = floors.ToList().Count;
            //TaskDialog.Show("nombre de sols dans le niveau", string.Format("niveau :{0} , nombre de sol : {1}", level.Name.ToString(), compte));

            foreach (Floor floor in floors)
            {
                    
                foreach (CompoundStructureLayer layer in floor.FloorType.GetCompoundStructure().GetLayers())
                {
                    if (layer.Function == MaterialFunctionAssignment.Structure & floor.Name.Contains("Terrasse") == false)
                    {
                        double surfacesol = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                        surfacesol = UnitUtils.ConvertFromInternalUnits(surfacesol, DisplayUnitType.DUT_SQUARE_METERS);
                        SDPniveau += surfacesol;
                    }
                }

            }

            return SDPniveau;
        }


        public static double Surfacemurextparniveau(Document doc, Level level)
        {
            double surfacetot = 0.00;
            ElementId levelid = level.Id;
            ElementLevelFilter parniveau = new ElementLevelFilter(levelid);
            FilteredElementCollector walls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().WherePasses(parniveau);

            foreach (Wall wall in walls)
            {
                string fonction = wall.WallType.LookupParameter("Fonction").AsString();

                if (fonction == "exterieur")
                {
                    double longueur = wall.LookupParameter("Longueur").AsDouble();
                    longueur = UnitUtils.ConvertFromInternalUnits(longueur, DisplayUnitType.DUT_METERS);
                    double epais = wall.WallType.Width;
                    epais = UnitUtils.ConvertFromInternalUnits(epais, DisplayUnitType.DUT_METERS);
                    double surfacemurext = epais * longueur;
                    surfacetot += surfacemurext;
                  
                }
            }
            return surfacetot;
        }
    }
}
