using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
namespace Indicateurs
{    
    class Typoappart
    {       
        public static (Dictionary<string, string>,Dictionary<string,double>, Dictionary<string, double>) infoappart(List<Document> docs)
        {
            //CREATION DE DICTIONNAIRE
            Dictionary<string, double> piececom = new Dictionary<string, double>();// NOM PIECE COMMUN ET SURFACE
            Dictionary<string, string> typos = new Dictionary<string, string>();// CAGE+APPART ET TYPOLOGIE 
            Dictionary<string,double> apartareas = new Dictionary<string, double>();// CAGE+APPART ET SURFACE

            //CHAQUE LIEN
            foreach (Document doc in docs)
            {          
                //COLLECTE DES PIECES
                IList<Element> rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements();
                foreach (Room room in rooms)
                {
                    try
                    {//ON RECUPERE LES PARAMETRES DE LA PIECE
                        string appart = room.LookupParameter("Appartement").AsString();
                        string cage = room.LookupParameter("Cage").AsString();
                        string typo = room.LookupParameter("Type de logement").AsString();
                        string cageappart = cage + appart;
                        double area = room.Area;
                        area = UnitUtils.ConvertFromInternalUnits(area, DisplayUnitType.DUT_SQUARE_METERS);

                        if (typo != "" & cageappart!="")
                        {
                            if (typos.ContainsKey(cageappart))
                            {
                                apartareas[cageappart] += area;                              
                            }
                            else
                            {
                                typos.Add(cageappart, typo);
                                apartareas.Add(cageappart, area);
                            }
                        }
                        else
                        {
                            string nom = room.Name;
                            double surface = room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble();
                            surface = UnitUtils.ConvertFromInternalUnits(surface, DisplayUnitType.DUT_SQUARE_METERS);
                            surface = Math.Round(surface, 2);
                            if (piececom.ContainsKey(nom))
                            {
                                piececom[nom] += surface;
                            }
                            else
                            {
                                piececom.Add(nom, surface);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            return (typos,apartareas, piececom);
        }
        public static Dictionary<string, List<double>> Dictypo(Dictionary<string, string> typappart, Dictionary<string,double> areaappart)
        {
            Dictionary<string, List<double>> dicstypos = new Dictionary<string,List<double>>();
            foreach (string appart in typappart.Keys)
            {
                try
                {
                    if (dicstypos.ContainsKey(typappart[appart]))
                    {
                        // on ajoute 1 au nombre de logement de cette typologie
                        dicstypos[typappart[appart]][0] += 1;
                        // on remet la moyenne de surface pour cette typologie de logement a jour
                        dicstypos[typappart[appart]][1] = (((dicstypos[typappart[appart]][0]-1) * dicstypos[typappart[appart]][1] + areaappart[appart]) /dicstypos[typappart[appart]][0]);
                    }
                    else
                    {

                        List<double> areatyp = new List<double> { 1, areaappart[appart] };
                        dicstypos.Add(typappart[appart], areatyp);
                    }
                }
                catch
                {
                }
            }
            return dicstypos;
        }
    }
}
