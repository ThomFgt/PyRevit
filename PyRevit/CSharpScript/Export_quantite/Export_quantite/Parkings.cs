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
    class Parking
    {
        public static (Dictionary<string, int>, int) parkinglots(List<Document> docs)
        {
            Dictionary<string, int> listparking = new Dictionary<string, int>();
            int sommeplace = 0;
            foreach (Document doc in docs)
            {
                IList<Element> parkings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Parking).WhereElementIsNotElementType().ToElements();
                sommeplace += parkings.Count;
                foreach (Element park in parkings)
                {
                    string typeplace = park.Name;
                    if (listparking.ContainsKey(typeplace))
                    {
                        listparking[typeplace] += 1;
                    }
                    else
                    {
                        listparking.Add(typeplace, 1);
                    }
                }
            }
            return (listparking, sommeplace);
        }
    }
}