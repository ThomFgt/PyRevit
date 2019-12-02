using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace Indicateurs
{
    public class Gardecorps
    {
        public static (Dictionary<string, double>,Dictionary<int,string>) metregardecorps(List<Document> docs)
        {
            Dictionary<string, double> typegardecorps = new Dictionary<string, double>();
            Dictionary<int,string> erreursgc = new Dictionary<int, string>();
            int erreurgardecorps = 0;

            foreach (Document doc in docs)
            {
                ElementMulticategoryFilter multigc = new ElementMulticategoryFilter(new BuiltInCategory[] { BuiltInCategory.OST_Railings, BuiltInCategory.OST_StairsRailing });
                IList<Element> gardecorps = new FilteredElementCollector(doc).WherePasses(multigc).WhereElementIsNotElementType().ToList();
                //TaskDialog.Show("Nombre garde corps", string.Format("Lien Revit: {0};" +Environment.NewLine+ "Nombre de type de garde corps: {1}",titredoc,gardecorps.Count));               
                foreach (Element rail in gardecorps)
                {
                    string typegarde = rail.Name;
                    int gardeid = rail.GetTypeId().IntegerValue;
                    double longueur = 0;
                    try
                    {
                        longueur = rail.LookupParameter("Longueur").AsDouble();// getparameter (longueur) pour les garde corps n'existe pas...) ne fonctionne donc pas en anglais


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
