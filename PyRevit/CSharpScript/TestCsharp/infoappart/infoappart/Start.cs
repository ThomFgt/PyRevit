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
using infoappart;
using Excel= Microsoft.Office.Interop.Excel;

namespace infoappart
{
   [Transaction(TransactionMode.ReadOnly)]
    class Start : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            { 
                // nbr fenetre appart
                List<int> listnbrfen = new List<int>();

                // typo appart
                List<string> listtypo = new List<string>();

                //List fenetre compté
                List<ElementId> fenetresids = new List<ElementId>();

                //Nombre de piece appart
                //List<int> nbrpieces = new List<int>();

                // list hsfp
                List<double> Lhsfp = new List<double>();

                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;
                FilteredElementCollector doccollecteur = new FilteredElementCollector(doc);
                ElementCategoryFilter filtreplafond = new ElementCategoryFilter(BuiltInCategory.OST_Ceilings);
                ElementCategoryFilter filtresol = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                Dictionary<string, List<Room>> DicoAppart = recupinfopiece.Getpieceparappart(doc);

            
                //surface habitable
                List<double> listshabs = recupinfopiece.SHABs(DicoAppart);

                //Collecte des donnee des pieces des appart
                foreach (string appart in DicoAppart.Keys)
                {
                    int nbrfenappart = 0;
                    int nbrpiece = 0;
                    //list hsfp
                    List<double> listhsfps = new List<double>();

                    //on recup la typo
                    listtypo.Add(DicoAppart[appart][0].LookupParameter("Type de logement").AsString());

                    foreach (Room piece in DicoAppart[appart])
                    {
                        //nbrpiece += 1;

                        //on recup les fenetres
                        List<ElementId> windows = recupinfopiece.collectwindows(doc, piece,fenetresids);
                        nbrfenappart += windows.Count;
                        fenetresids.Concat(windows);

                        //on recup les hsfp
                        double hsfp = Hsfp.Gethsfp(doc, piece, filtreplafond);
                        if (hsfp ==0.00 | hsfp>=4.00)
                        {
                            hsfp = Hsfp.Gethsfs(doc, piece, filtresol);
                        }
                        if (hsfp <1.80 | hsfp>4)
                        {
                            hsfp = 0.00;
                        }

                        listhsfps.Add(hsfp);
                    }
                    //nbrpieces.Add(nbrpiece);
                    listnbrfen.Add(nbrfenappart);
                    double haut = Hsfp.hsfpmoy(listhsfps);
                    Lhsfp.Add(haut);
                }
                Excel.Application xlapp = new Excel.Application();
                Exportexcel export = new Exportexcel(xlapp, DicoAppart, listtypo, listshabs, Lhsfp, listnbrfen);
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;

                return Result.Failed;
            }

        }
    }
}
