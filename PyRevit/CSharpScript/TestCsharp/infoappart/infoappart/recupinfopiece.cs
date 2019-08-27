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
    
    public class recupinfopiece 
    {

        public static Dictionary<string, List<Room>> Getpieceparappart(Document doc)
        {

            //colllecteur de piece
            FilteredElementCollector doccollecteur = new FilteredElementCollector(doc);
            FilteredElementCollector listepiece = doccollecteur.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();
            Dictionary<string, List<Room>> DicoAppart = new Dictionary<string, List<Room>>();

            //On complete le dictionnaire appartement piece
            foreach (Room piece in listepiece)
            {
                string appart = piece.LookupParameter("Appartement").AsString();
                string cage = piece.LookupParameter("Cage").AsString();
                string cageappart = cage + appart;
                string nompiece = piece.Name.ToString();
                string typeappart = piece.LookupParameter("Type de logement").AsString();

                // Premier tri des pièces par rapport a leur nom et leur  string appartement

                if (appart != "" & cage != "" & appart != null & cage != null & appart != "000" &
                    nompiece.Contains("CAGE") == false & nompiece.Contains("CIRCULATION") == false & nompiece.Contains("LOCAL") == false & nompiece.Contains("00") == false)
                {
                    //si l'appartement est repertorie
                    if (DicoAppart.ContainsKey(cageappart))
                    {
                        DicoAppart[cageappart].Add(piece);
                    }
                    // SI LAPPARTEMENT CONTENANT LA PIECE NEST PAS REPERTORIE
                    else
                    {
                        List<Room> rooms = new List<Room>
                        {
                            piece
                        };
                        DicoAppart.Add(cageappart, rooms);
                    }
                }
            }
            return DicoAppart;
        }


        public static List<double> SHABs (Dictionary<string,List<Room>> dico)
        {
            List<double> shabs = new List<double>();

            foreach (string appart in dico.Keys)
            {
                double shabappart = 0.00;
                foreach (Room piece in dico[appart])
                {
                    double surfacepiece = piece.Area;
                    surfacepiece = UnitUtils.ConvertFromInternalUnits(surfacepiece, DisplayUnitType.DUT_SQUARE_METERS);
                    surfacepiece = Math.Round(surfacepiece,2);
                    shabappart += surfacepiece;
                }
                shabs.Add(shabappart);
            }
            return shabs;
        }

        public static List<ElementId> collectwindows(Document doc, Room room,List<ElementId> listfenid)
        {
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            SpatialElement piece = room as SpatialElement;
            BoundingBoxXYZ bbpiece = piece.get_BoundingBox(null);
            IList<IList<BoundarySegment>> contourpiece = piece.GetBoundarySegments(options);
            Outline outroom = new Outline(new XYZ(bbpiece.Min.X, bbpiece.Min.Y, bbpiece.Min.Z), new XYZ(bbpiece.Max.X, bbpiece.Max.Y, bbpiece.Max.Z));
            BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outroom);
            FilteredElementCollector windows = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WherePasses(bbfilter).WhereElementIsNotElementType();
            List<ElementId> fenetres = windows.ToElementIds().ToList();

            foreach(ElementId fenid in listfenid)
            {
                if(fenetres.Contains(fenid))
                {
                    fenetres.Remove(fenid);
                }                
            }
            return fenetres;
        }                 
    }
}