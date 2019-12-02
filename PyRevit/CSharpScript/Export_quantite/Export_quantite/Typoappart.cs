using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
namespace Export_quantite
{
    class Typoappart
    {
        public static (Dictionary<string, string>, Dictionary<string, double>) infoappart(List<Document> docs)
        {
            Dictionary<string, double> dicpieces = new Dictionary<string, double>();
            Dictionary<string, string> typos = new Dictionary<string, string>();
            List<string> cageapparts = new List<string>();
            foreach (Document doc in docs)
            {
                string titredoc = doc.Title;
                IList<Element> rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements();
                foreach (Room room in rooms)
                {
                    try
                    {
                        string appart = room.LookupParameter("Appartement").AsString();
                        string cage = room.LookupParameter("Cage").AsString();
                        string typo = room.LookupParameter("Type de logement").AsString();
                        string cageappart = cage + appart;
                        if (typo != "" & cageappart != "")
                        {
                            if (typos.ContainsKey(cageappart))
                            {
                                continue;
                            }
                            else
                            {
                                typos.Add(cageappart, typo);
                            }
                        }
                        else
                        {
                            string nom = room.Name;
                            double surface = room.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble();
                            surface = UnitUtils.ConvertFromInternalUnits(surface, DisplayUnitType.DUT_SQUARE_METERS);
                            surface = Math.Round(surface, 2);
                            if (dicpieces.ContainsKey(nom))
                            {
                                dicpieces[nom] += surface;
                            }
                            else
                            {
                                dicpieces.Add(nom, surface);
                            }
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
            return (typos, dicpieces);
        }
        public static Dictionary<string, int> Dictypo(Dictionary<string, string> dicappart)
        {
            Dictionary<string, int> dicstypos = new Dictionary<string, int>();
            foreach (string appart in dicappart.Keys)
            {
                try
                {
                    if (dicstypos.ContainsKey(dicappart[appart]))
                    {
                        dicstypos[dicappart[appart]] += 1;
                    }
                    else
                    {
                        dicstypos.Add(dicappart[appart], 1);
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
