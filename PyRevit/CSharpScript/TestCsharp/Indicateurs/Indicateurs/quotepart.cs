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

namespace Indicateurs
{
    class quotepart
    {
        public static (Dictionary<string, List<double>>, Dictionary<ElementId, List<double>>, Dictionary<ElementId, Dictionary<string,double>>) BureauLot(List<Document> docs)
        {
            // creation des dicos
            Dictionary<string, List<double>> dicolotsurface = new Dictionary<string, List<double>>();// LOT ET SURFACE BUREAU ET REUNION
            Dictionary<ElementId, List<double>> dicolevelsurface = new Dictionary<ElementId, List<double>>();// ETAGE ET SURFACE BUREAU ET REUNION
            Dictionary<ElementId, Dictionary<string,double>> dicolotlevel = new Dictionary<ElementId, Dictionary<string,double>>(); // ETAGE ET DICO LOT ET SURFACE

            //Liste pour compte 1 seule fois chaque piece
            List<int> pieceids = new List<int>();

            //On liste les erreurs
            int nombreerreur = 0;
            List<string> messageerreur = new List<string>();

            // Regex Bureau et Reunion
            Regex rexbur = new Regex(@"[bB][uU][rR][eE][aA][uU]", RegexOptions.Singleline);
            Regex rexreu = new Regex(@"[rR][é|eE][uU][nN][iI][oO][nN]", RegexOptions.Singleline);
            //On passe tous les documents
            foreach (Document doc in docs)
            {
                FilteredElementCollector pieces = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();
                foreach (Element piece in pieces)
                {
                    try
                    { 
                        //on check si la piece a deja ete compté
                        if (pieceids.Contains(piece.Id.IntegerValue))
                        {
                            continue;
                        }
                        // on check que le parametre "Service" n'est pas nul
                        if (piece.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT)==null)
                        {
                            nombreerreur += 1;
                            continue;
                        }

                        string lotpiece = piece.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString();
                        //on verifie si le parametre "Service" n'est pas vide
                        if (lotpiece == "" )
                        {
                            nombreerreur += 1;
                            continue;                      
                        }

                        ElementId levelpieceid = piece.LevelId;
                        Element levelpiece = doc.GetElement(levelpieceid);
                        double surfacepiece = piece.get_Parameter(BuiltInParameter.ROOM_AREA).AsDouble();
                        surfacepiece = UnitUtils.ConvertFromInternalUnits(surfacepiece, DisplayUnitType.DUT_SQUARE_METERS);
                        surfacepiece = Math.Round(surfacepiece, 2);
                        string nompiece = piece.Name;

                        //ANALYSE
                        //NOUVEAU LOT DE PIECE POUR SURFACE LOT ET BUREAU            
                        if (dicolotsurface.ContainsKey(lotpiece) == false)
                        {
                            
                            List<double> listetriplezero = new List<double> { 0.00, 0.00, 0.00 };
                            dicolotsurface.Add(lotpiece, listetriplezero);
                            dicolotsurface[lotpiece][0] += surfacepiece;
                            //SI LA PIECE EST UN BUREAU
                            if (rexbur.IsMatch(nompiece))
                            {
                                dicolotsurface[lotpiece][1] += surfacepiece;
                            }
                            //SI LA PIECE EST UNE SALLE DE REUNION
                            else if (rexreu.IsMatch(nompiece))
                            {
                                dicolotsurface[lotpiece][2] += surfacepiece;
                            }
                        }

                        //LOT DEJA VU
                        else
                        {
                            //
                            dicolotsurface[lotpiece][0] += surfacepiece;
                            if (rexbur.IsMatch(nompiece))
                            {
                                dicolotsurface[lotpiece][1] += surfacepiece;
                            }
                            else if (rexreu.IsMatch(nompiece))
                            {
                                dicolotsurface[lotpiece][2] += surfacepiece;
                            }
                        }
                        // SI NOUVEL ETAGE DICO SURFACE ETAGE ET BUREAU REU
                        if (dicolevelsurface.ContainsKey(levelpieceid) == false)
                        {
                            List<double> listedoublezero = new List<double> { 0.00, 0.00 };
                            dicolevelsurface.Add(levelpieceid, listedoublezero);
                            if (rexbur.IsMatch(nompiece))
                            {
                                dicolevelsurface[levelpieceid][0] += surfacepiece;
                            }
                            else if (rexreu.IsMatch(nompiece))
                            {
                                dicolevelsurface[levelpieceid][1] += surfacepiece;
                            }
                        }
                        else
                        {
                            if (rexbur.IsMatch(nompiece))
                            {
                                dicolevelsurface[levelpieceid][0] += surfacepiece;
                            }
                            else if (rexreu.IsMatch(nompiece))
                            {
                                dicolevelsurface[levelpieceid][1] += surfacepiece;
                            }
                        }
                        // SURFACE ETAGE LOT
                        if (dicolotlevel.ContainsKey(levelpieceid) == false)
                        {
                            Dictionary<string, double> nvletage = new Dictionary<string, double>();
                            nvletage.Add(lotpiece, surfacepiece);
                            dicolotlevel.Add(levelpieceid, nvletage);
                        }
                        else
                        {
                            if (dicolotlevel[levelpieceid].ContainsKey(lotpiece))
                            {
                                dicolotlevel[levelpieceid][lotpiece] += surfacepiece;
                            }
                            else
                            {
                                dicolotlevel[levelpieceid].Add(lotpiece, surfacepiece);
                            }
                        }
                    }
                    catch(Exception e)
                    {
                        string meser = e.Message;
                        if (messageerreur.Contains(meser)!=true)
                        {
                            messageerreur.Add(meser);
                        }
                        nombreerreur += 1;
                    }
                }
            }
            
            if (nombreerreur!=0)
            {
                foreach (string erreur in messageerreur)
                {
                    TaskDialog.Show(string.Format("{0} erreur(s) rencontrée(s) lors de l'export des pièces", nombreerreur), "Message d'erreur :" + Environment.NewLine + string.Format("{0}", erreur));
                }
               
            }
            return (dicolotsurface, dicolevelsurface, dicolotlevel);
        }
    }
}
