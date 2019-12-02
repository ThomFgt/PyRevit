using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;
using Indicateurs;
using forms = System.Windows.Forms;

namespace Indicateurs
{
    [Transaction(TransactionMode.ReadOnly)]
    public class Indicateurscommand : IExternalCommand
    {        
      
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //RUN EXCEL
            Excel.Application xlapp = new Excel.Application();
            object misvalue = System.Reflection.Missing.Value;
            Excel.Workbook workbook = xlapp.Workbooks.Add();
                       
            //UIDOC - DOCS- LIENS
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Options option = new Options();
            option.DetailLevel = ViewDetailLevel.Fine;

            // CHECK SURFACE CALCUL METHOD
            AreaVolumeSettings surfacefinition = AreaVolumeSettings.GetAreaVolumeSettings(doc);
            if (surfacefinition.GetSpatialElementBoundaryLocation(SpatialElementType.Room) != SpatialElementBoundaryLocation.Finish)
            {
                TaskDialog changeboundary = new TaskDialog("Attention à la définition de surface de vos pièces !");
                changeboundary.MainInstruction = "La surface des pièces n'est pas définies par rapport aux fini des murs. Vous pouvez modifier la méthode de calcul dans l'onglet 'Architecture-Pièces'.";
                changeboundary.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Continuer");
                changeboundary.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Annuler");
                if (changeboundary.Show() == TaskDialogResult.CommandLink2)
                {
                    return Result.Cancelled;
                }
            }

            //PRISE EN COMPTE DES LIENS REVIT
            FilteredElementCollector rvtlinks = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_RvtLinks);
            List<Document> docs = new List<Document>();
            docs.Add(doc);
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
            catch (Exception)
            { }


            // CHOIX ENTRE BUREAU ET LOGEMENT
            TaskDialog introduction = new TaskDialog("B&P - Plug-in Indicateurs");
            introduction.MainInstruction = "De quel type est le batîment de votre maquette ?";
            introduction.MainContent = "Attention aux doublons des maquettes en lien avant d'utiliser ce plugin." + string.Format("{0} document(s) récupéré(s) dans la maquette.",docs.Count);
            introduction.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Maquette de Logement");
            introduction.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Maquette de Bureaux");
            introduction.CommonButtons = TaskDialogCommonButtons.Cancel;
            TaskDialogResult tintroresult = introduction.Show();

            //SI LOGEMENT
            if(tintroresult == TaskDialogResult.CommandLink1)
            {
                //TYPOAPPART
                (Dictionary<string, string> listappart,Dictionary<string,double>areaappart, Dictionary<string, double> autrepieces) = Typoappart.infoappart(docs);
                Dictionary<string,List<double>> typologie = Typoappart.Dictypo(listappart,areaappart);
                Exportexcel.Exporttypoappart(xlapp, workbook, listappart,areaappart, typologie, autrepieces);                
            }
            //SI BUREAU
            else if (tintroresult==TaskDialogResult.CommandLink2)
            {
                //QUOTEPART
                (Dictionary<string, List<double>> dicolotsurface, Dictionary<ElementId, List<double>> dicolevelsurface, Dictionary < ElementId, Dictionary<string,double>> dicolotlevel )= quotepart.BureauLot(docs);
                Exportexcel.Exportbureaulot(docs, xlapp, workbook, dicolotsurface, dicolevelsurface, dicolotlevel);                
            }
            //ANNULATION
            else if (tintroresult == TaskDialogResult.Cancel)
            {
                return Result.Cancelled;
            }

            //PARKINGS
            (Dictionary<string, int> parkings, int sommeplace) = Parking.parkinglots(docs);
            Exportexcel.exportparking(xlapp, workbook, parkings, sommeplace);


            //GARDE CORPS
            (Dictionary<string, double> gardecorps, Dictionary<int, string> erreursgc) = Gardecorps.metregardecorps(docs);
            Exportexcel.exportgardecorps(xlapp, workbook, gardecorps, erreursgc);

            //FACADE VIDE POUR PLEIN
            TaskDialog factask = new TaskDialog("Calcul de façade");
            factask.MainInstruction = "Calculer la façade vide pour plein ?";
            factask.MainContent = "Attention ! Modélisation adéquate exigée !";
            factask.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            if (factask.Show()==TaskDialogResult.Yes)
            {
                //FACADE VIDE POUR PLEIN
                (double surfacefacade,int comptemur,double surfacepanneaux) = Facadevidepourplein.facadeopaque(docs);
                (Dictionary<ElementId, double> dicfenvitre, List<List<string>> nomtypefen) = Facadevidepourplein.surfacefenetrevitre(docs, option);
                double surfacefenetres = 0;
                foreach (ElementId wintypeid in dicfenvitre.Keys)
                {
                    surfacefenetres += dicfenvitre[wintypeid];
                }
                double facadevidepourplein = (surfacefenetres + surfacepanneaux) / surfacefacade;
                facadevidepourplein = Math.Round(facadevidepourplein, 2);
                Exportexcel.exportfacadevidepourplein(xlapp, workbook, surfacefacade, dicfenvitre, nomtypefen, surfacefenetres, surfacepanneaux, facadevidepourplein,comptemur);
            }


            //CHOIX SAUVEGARDE
            string nommaquette = System.IO.Path.GetFileNameWithoutExtension(doc.Title);
            TaskDialog demandesauvegarde = new TaskDialog("Sauvegarder ?");
            demandesauvegarde.MainInstruction = "Sauvegarder et ouvrir le fichier Excel ?";
            demandesauvegarde.MainContent = "Nom du fichier:" + "Export-Quantité-"+nommaquette;
            demandesauvegarde.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
            TaskDialogResult tresult5 = demandesauvegarde.Show();
            if(tresult5 == TaskDialogResult.Yes)
            {
                System.Windows.Forms.FolderBrowserDialog liens = new System.Windows.Forms.FolderBrowserDialog();
                liens.Description = "INDIQUEZ LE DOSSIER DE SAUVEGARDE DES DONNEES";
                if (liens.ShowDialog()==System.Windows.Forms.DialogResult.Cancel )
                {
                    return Result.Cancelled;
                }
                string path = liens.SelectedPath+"/ "+ "Export-Quantité-"+nommaquette;
                Exportexcel.OpenExcel(xlapp, workbook, path);
                return Result.Succeeded;
            }
            else
            {
                return Result.Failed;
            }
            
        }
    }
}
