using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Windows;

namespace infoappart
{
    [TransactionAttribute(TransactionMode.ReadOnly)]
    public class recupinfopiece : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get Document
            Document doc = uidoc.Document;

            //Dictionnaire : numero appart; liste : nbr piece, shab, surface vitre
            Dictionary<string, List<double>> DicoAppart = new Dictionary<string, List<double>>();
            Dictionary<string, double> DicoCommun = new Dictionary<string, double>();
            //colllecteur de piece
            FilteredElementCollector doccollecteur = new FilteredElementCollector(doc);
            IList<Element> listepiece = doccollecteur.OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements();
            ElementCategoryFilter filtresol = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
            ElementCategoryFilter filtreplafond = new ElementCategoryFilter(BuiltInCategory.OST_Ceilings);
            List<string> typedelogement = new List<string>();

            foreach (Element piece in listepiece)
            {
                            
                double surfacepiece = piece.LookupParameter("Surface").AsDouble();
                surfacepiece = UnitUtils.ConvertFromInternalUnits(surfacepiece, DisplayUnitType.DUT_SQUARE_METERS);
                surfacepiece = Math.Round(surfacepiece, 2);
                string appart = piece.LookupParameter("Appartement").AsString();
                string cage = piece.LookupParameter("Cage").AsString();
                string cageappart = cage + appart;
                double hsfp = 0.00;
                double hsfs = 0.00;
                string nompiece = piece.Name.ToString();
                string typeappart = piece.LookupParameter("Type de logement").AsString();




                // liste de lensemble des donnees de la piece
                List<double> donneappart = new List<double>();




                // ON CALCULE LA HAUTEUR SOUS PLAFOND OU DALLE
                try
                {
                    LocationPoint locpiece = piece.Location as LocationPoint;
                    XYZ ptpiece = locpiece.Point;
                    ptpiece = new XYZ(ptpiece.X, ptpiece.Y, ptpiece.Z + 0.1);
                    XYZ vecteurpiece = new XYZ(0, 0, 1);
                    ReferenceIntersector refi = new ReferenceIntersector(filtreplafond, FindReferenceTarget.Face, (View3D)doc.ActiveView);
                    ReferenceWithContext refc = refi.FindNearest(ptpiece, vecteurpiece);
                    Reference reference = refc.GetReference();
                    XYZ intpoint = reference.GlobalPoint;
                    hsfp = ptpiece.DistanceTo(intpoint);
                    hsfp = UnitUtils.ConvertFromInternalUnits(hsfp, DisplayUnitType.DUT_METERS);
                    hsfp = Math.Round(hsfp, 2);
                    //MessageBox.Show(string.Format("hauteur sous dalle : {0}", hsfp));
                }
                catch (Exception)
                {
                    //MessageBox.Show("pas de plafond au dessus de la piece","erreur");
                    hsfp = 0.00;

                }

                // on verifie si sol au dessus de piece
                try
                {
                    LocationPoint locpiece = piece.Location as LocationPoint;
                    XYZ ptpiece = locpiece.Point;
                    ptpiece = new XYZ(ptpiece.X, ptpiece.Y, ptpiece.Z + 0.1);
                    XYZ vecteurpiece = new XYZ(0, 0, 1);
                    ReferenceIntersector refi = new ReferenceIntersector(filtresol, FindReferenceTarget.Face, (View3D)doc.ActiveView);
                    ReferenceWithContext refc = refi.FindNearest(ptpiece, vecteurpiece);
                    Reference reference = refc.GetReference();
                    XYZ intpoint = reference.GlobalPoint;
                    hsfs = ptpiece.DistanceTo(intpoint);
                    hsfs = UnitUtils.ConvertFromInternalUnits(hsfs, DisplayUnitType.DUT_METERS);
                    hsfs = Math.Round(hsfs, 2);
                    //MessageBox.Show(string.Format("hauteur sous dalle : {0}", hsfs));
                }
                catch (Exception)
                {
                    //MessageBox.Show("pas de sol au dessus de la piece","erreur");
                    hsfs = 0.00;
                }

                // SI LAPPARTEMENT CONTENANT LA PIECE EST REPERTORIE    
                if (appart != "" && cage != "" && appart != null && cage != null)
                {

                    //si l'appartement est repertorie
                    if (DicoAppart.ContainsKey(cageappart))
                    {
                        if (nompiece.Contains("BALCON") | nompiece.Contains("TERRASSE"))
                        {
                            DicoAppart[cageappart][4] += surfacepiece;
                        }
                        else
                        {
                            DicoAppart[cageappart][0] += 1.00;
                            DicoAppart[cageappart][1] += surfacepiece;

                            //si le faud plafond calcule est plus faible que le precedent
                            if (hsfp <= DicoAppart[cageappart][2] && hsfp > 1.00)
                            {
                                DicoAppart[cageappart][2] = hsfp;
                            }
                            if (hsfs <= DicoAppart[cageappart][2] && hsfs > 1.00)
                            {
                                DicoAppart[cageappart][2] = hsfs;
                            }
                        }
                    }

                    // SI LAPPARTEMENT CONTENANT LA PIECE NEST PAS REPERTORIE
                    else
                    {
                        if (nompiece.Contains("BALCON") | nompiece.Contains("TERRASSE"))
                        {
                            donneappart.Add(0.00);
                            donneappart.Add(0.00);
                            donneappart.Add(0.00);
                            donneappart.Add(0.00);
                            donneappart.Add(surfacepiece);
                        }
                        else
                        {
                            donneappart.Add(1.00);
                            donneappart.Add(surfacepiece);

                            if (hsfp > 1.00)
                            {
                                donneappart.Add(hsfp);
                            }
                            else
                            {
                                donneappart.Add(0.00);
                            }
                            if (hsfs > 1.00)
                            {
                                donneappart.Add(hsfs);
                            }
                            else
                            {
                                donneappart.Add(0.00);
                            }
                            donneappart.Add(0.00);
                            typedelogement.Add(typeappart);
                        }
                        DicoAppart.Add(cageappart, donneappart);

                    }

                }
                else
                {
                    if (DicoCommun.ContainsKey(nompiece))
                    {
                        DicoCommun[nompiece] += surfacepiece;
                    }
                    else
                    {
                        DicoCommun.Add(nompiece, surfacepiece);
                    }
                }

            }
            try
            {
                exportexcel export = new exportexcel(DicoAppart,DicoCommun,typedelogement);
                return Result.Succeeded;
            }
            catch
            {
                TaskDialog.Show("dommage", "pas d'excel enregistre");
                return Result.Failed;
            }

            // MESSAGE BOX DES APPARTEMENTS
            //foreach (string a in DicoAppart.Keys)
            //{
            //    _ = MessageBox.Show(string.Format("Nombre de pièce : {0}" + Environment.NewLine +
            //    "surface habitable : {1}" + Environment.NewLine +
            //    "hauteur sous plafond {2}" + Environment.NewLine +
            //    "Hauteur sous dalle {3}" + Environment.NewLine +
            //    "Surface du balcon {4}",
            //    DicoAppart[a][0],
            //    DicoAppart[a][1],
            //    DicoAppart[a][2],
            //    DicoAppart[a][3],
            //    DicoAppart[a][4]),
            //    string.Format("numero d'appartement : {0}", a));

            //}


        }
    }
}