using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using SYAToolbox;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Export_quantite
{
    [Transaction(TransactionMode.Manual)]
    public class Start : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ConnectOptionsData.AppsInfos info = ConnectOptionsData.GetSavedToken(
                    new SYAApplication("Export_quantite", null, null, -1));

            SYAApplication SYAApp = new SYAApplication
            (
                info.appName,
                info.clientId,
                info.server,
                info.port
            );

            Authentify auth = new Authentify(SYAApp, false);
            if (auth.Authentificate())
            {
                return InitPlugin(commandData);
            }
            else
            {
                return Result.Failed;
            }
        }
        public Result InitPlugin(ExternalCommandData commandData)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            Excel.Application xlapp = new Excel.Application();
            Excel.Workbook workbook = xlapp.Workbooks.Add();
            Excel.Worksheet pagedegarde = workbook.Sheets[1];
            pagedegarde.Name = "Page de garde";
            FilteredElementCollector rvtlinks = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_RvtLinks);
            List<Document> docs = new List<Document>();
            docs.Add(doc);
            foreach (RevitLinkInstance rvtlink in rvtlinks)
            {
                docs.Add(rvtlink.GetLinkDocument());
            }
            TaskDialog introduction = new TaskDialog("B&P - Plug-in Indicateurs");
            introduction.MainInstruction = "De quel type est le batîment de votre maquette ?";
            introduction.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Maquette de Logement ?");
            introduction.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Maquette de Bureaux ?");
            TaskDialogResult tintroresult = introduction.Show();
            if (tintroresult == TaskDialogResult.CommandLink1)
            {
                //APPARTEMENT
                (Dictionary<string, string> listappart, Dictionary<string, double> autrepieces) = Typoappart.infoappart(docs);
                Dictionary<string, int> typologie = Typoappart.Dictypo(listappart);
                Exportexcel.Exporttypoappart(xlapp, workbook, listappart, typologie, autrepieces);

            }
            else if (tintroresult == TaskDialogResult.CommandLink2)
            {
                //QUOTEPART
                Dictionary<string, List<double>> dicolotsurface = Lotbureau.BureauLot(docs).Item1;
                Dictionary<ElementId, List<double>> dicolevelsurface = Lotbureau.BureauLot(docs).Item2;
                Dictionary<ElementId, List<string>> dicolotlevel = Lotbureau.BureauLot(docs).Item3;
                Exportexcel.Exportbureaulot(doc, xlapp, workbook, dicolotsurface, dicolevelsurface, dicolotlevel);

            }
            return Result.Succeeded;
        }

    }
}
