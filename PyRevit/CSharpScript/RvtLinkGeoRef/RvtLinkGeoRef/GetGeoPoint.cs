using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;


namespace RvtLinkGeoRef
{
    class GetGeoPoint
    {
        public static ()
            {
                            foreach (Document doc in docs)
                {
                    FilteredElementCollector listebasepoint = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint);
        FilteredElementCollector listetopopoint = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint);
                    foreach(BasePoint point in listebasepoint)
                    {
                        point.Position.X;
                    }
}
    }
}
