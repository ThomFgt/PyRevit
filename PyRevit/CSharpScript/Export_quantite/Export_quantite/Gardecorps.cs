using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;

namespace Export_quantite
{
    public class Gardecorps
    {
        public static (Dictionary<string, double>, Dictionary<int, string>) metregardecorps(List<Document> docs)
        {
            Dictionary<string, double> typegardecorps = new Dictionary<string, double>();
            Dictionary<int, string> erreursgc = new Dictionary<int, string>();
            int erreurgardecorps = 0;

            foreach (Document doc in docs)
            {
                ElementMulticategoryFilter multigc = new ElementMulticategoryFilter(new BuiltInCategory[] { BuiltInCategory.OST_Railings, BuiltInCategory.OST_StairsRailing });
                IList<Element> gardecorps = new FilteredElementCollector(doc).WherePasses(multigc).WhereElementIsNotElementType().ToList();
            
                foreach (Element rail in gardecorps)
                {
                    string typegarde = rail.Name;
                    int gardeid = rail.GetTypeId().IntegerValue;
                    try
                    {
                        double longueur = rail.LookupParameter("Longueur").AsDouble();
                        longueur = UnitUtils.ConvertFromInternalUnits(longueur, DisplayUnitType.DUT_METERS);
                        longueur = Math.Round(longueur, 2);
                        if (typegardecorps.ContainsKey(typegarde))
                        {
                            typegardecorps[typegarde] += longueur;
                        }
                        else
                        {
                            typegardecorps.Add(typegarde, longueur);
                        }
                    }
                    catch
                    {
                        if (typegarde == "")
                        {
                            typegarde = "AUCUN NOM DE TYPE";
                        }
                        erreurgardecorps += 1;
                        if (erreursgc.ContainsKey(gardeid) == false)
                        {
                            erreursgc.Add(gardeid, typegarde);
                        }
                    }
                }
            }
            return (typegardecorps, erreursgc);
        }
    }
}
