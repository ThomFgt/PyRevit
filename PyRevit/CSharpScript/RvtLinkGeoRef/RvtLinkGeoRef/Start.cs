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
    [TransactionAttribute(TransactionMode.Manual)]
    public class Start : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {


            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document docprincipal = uidoc.Document;
            Options option = new Options();
            option.DetailLevel = ViewDetailLevel.Fine;

            FilteredElementCollector rvtlinks = new FilteredElementCollector(docprincipal, docprincipal.ActiveView.Id).OfCategory(BuiltInCategory.OST_RvtLinks);
            List<Document> docs = new List<Document>();
            docs.Add(docprincipal);
            try
            {
                foreach (RevitLinkInstance rvtlink in rvtlinks)
                {
                    if (docs.Contains(rvtlink.GetLinkDocument()) == false)
                    {
                        docs.Add(rvtlink.GetLinkDocument());
                    }
                }

            }
            catch
            {
                return Result.Failed;
            }
            return Result.Succeeded;
        }
    }
}
