using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows;
using revitapi = Autodesk.Revit.DB;
using revitui = Autodesk.Revit.UI;
using revitattri = Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;

namespace Indicateurs
{
    public class Exportexcel
    {
        
        //EXPORT APPARTEMENT
        public static void Exporttypoappart(Excel.Application xlapp,Excel.Workbook workbook, Dictionary<string, string> listappart,Dictionary<string,double> areaappart, Dictionary<string,List<double>>listtypo,Dictionary<string,double> autrepieces)
        {
            object misvalue = System.Reflection.Missing.Value;
            Excel.Worksheet worksheet1 = workbook.Sheets[1];
            worksheet1.Columns.ColumnWidth = 25;
            worksheet1.Name = "Logements";
            if (listappart.Count != 0)
            {
                Excel.XlBordersIndex lgauche = Excel.XlBordersIndex.xlEdgeLeft;
                Excel.XlBordersIndex ldroite = Excel.XlBordersIndex.xlEdgeRight;
                Excel.XlBordersIndex lhaut = Excel.XlBordersIndex.xlEdgeTop;
                Excel.XlBordersIndex lbas = Excel.XlBordersIndex.xlEdgeBottom;
                Excel.XlLineStyle lcontinue = Excel.XlLineStyle.xlContinuous;
                Excel.XlBorderWeight llarge = Excel.XlBorderWeight.xlMedium;
                Excel.XlBorderWeight lfin = Excel.XlBorderWeight.xlThin;
                Excel.XlVAlign centre = Excel.XlVAlign.xlVAlignCenter;
                Excel.XlVAlign justi = Excel.XlVAlign.xlVAlignJustify;
                Excel.XlRgbColor rouge = Excel.XlRgbColor.rgbRed;
                Excel.XlRgbColor gold = Excel.XlRgbColor.rgbGold;
                Excel.XlRgbColor vert = Excel.XlRgbColor.rgbLightGreen;
                Excel.XlRgbColor gris = Excel.XlRgbColor.rgbLightGrey;

                Excel.Range rangetitre1 = worksheet1.Range["A1"];
                rangetitre1.Value = "INDICATEUR LOGEMENT";
                rangetitre1.Font.Bold = true;
                rangetitre1.Font.Color = rouge;
                rangetitre1.Font.Size = 14;
                rangetitre1.VerticalAlignment = justi;
                rangetitre1.VerticalAlignment = centre;

                Excel.Range rangetitre2 = worksheet1.Range["A3"];
                rangetitre2.Value = "Nombre de logement total :";
                worksheet1.Range["A3", "B3"].BorderAround2(lcontinue, llarge);
                worksheet1.Range["A3", "B3"].Interior.Color = vert;

                Excel.Range nombrelogement = worksheet1.Range["B3"];
                nombrelogement.Value = listappart.Count;

                Excel.Range rangetypo = worksheet1.Range["A4"];
                rangetypo.Value = "Typologie";
                rangetypo.BorderAround2(lcontinue, llarge);
                rangetypo.Interior.Color = gold;
                rangetypo.Font.Bold = true;

                Excel.Range rangenbrappart = worksheet1.Range["A5"];
                rangenbrappart.Value = "nombre de type de logement";
                rangenbrappart.BorderAround2(lcontinue, llarge);
                rangenbrappart.Interior.Color = gold;
                rangenbrappart.Font.Bold = true;

                Excel.Range rangemoy = worksheet1.Range["A6"];
                rangemoy.Value = "Surface moyenne du type";
                rangemoy.BorderAround2(lcontinue, llarge);
                rangemoy.Interior.Color = gold;
                rangemoy.Font.Bold = true;

                Excel.Range rangetitre3 = worksheet1.Range["H8", "I8"];
                rangetitre3.BorderAround2(lcontinue, llarge);
                rangetitre3.Value = "Nombre de pièces communes";
                worksheet1.Cells[8, 9] = autrepieces.Count;
                rangetitre3.Interior.Color = vert;
                

                Excel.Range rangetitre4 = worksheet1.Range["H9"];
                rangetitre4.BorderAround2(lcontinue, llarge);
                rangetitre4.Value = "Pièces communes";
                rangetitre4.Font.Bold = true;
                rangetitre4.VerticalAlignment = centre;
                rangetitre4.Interior.Color = gold;

                Excel.Range rangetitreC = worksheet1.Range["I9"];
                rangetitreC.BorderAround2(lcontinue, llarge);
                rangetitreC.Value = "Surface (m²)";
                rangetitreC.Font.Bold = true;
                rangetitreC.VerticalAlignment = centre;
                rangetitreC.Interior.Color = gold;

                Excel.Range rangetitre5 = worksheet1.Range["A9"];
                rangetitre5.BorderAround2(lcontinue, llarge);
                rangetitre5.Value = "Appartements";
                rangetitre5.Font.Bold = true;
                rangetitre5.VerticalAlignment = centre;
                rangetitre5.Interior.Color = gold;

                Excel.Range rangetitre6 = worksheet1.Range["B9"];
                rangetitre6.BorderAround2(lcontinue, llarge);
                rangetitre6.Value = "Typologies";
                rangetitre6.Font.Bold = true;
                rangetitre6.VerticalAlignment = centre;
                rangetitre6.Interior.Color = gold;

                Excel.Range rangetitre7 = worksheet1.Range["C9"];
                rangetitre7.BorderAround2(lcontinue, llarge);
                rangetitre7.Value = "Surfaces (m²)";
                rangetitre7.Font.Bold = true;
                rangetitre7.VerticalAlignment = centre;
                rangetitre7.Interior.Color = gold;

                int i = 9;
                foreach (string appart in listappart.Keys)
                {
                    i += 1;
                    Excel.Range nomappartcell = worksheet1.Cells[i, 1];
                    nomappartcell.Value = appart;
                    Excel.Range typapartcell = worksheet1.Cells[i, 2];
                    typapartcell.Value = listappart[appart];
                    double surfaceappart = areaappart[appart];
                    surfaceappart = Math.Round(surfaceappart, 2);
                    Excel.Range surfappartcell = worksheet1.Cells[i, 3];
                    surfappartcell.Value = surfaceappart;
                    if (i%2 ==0)
                    {
                        nomappartcell.Interior.Color = gris;
                        typapartcell.Interior.Color = gris;
                        surfappartcell.Interior.Color = gris;
                    }           
                }
                string pBG = "A" + i.ToString();
                string pBD = "C" + i.ToString();
                Excel.Range rangeA = worksheet1.Range["A10", pBG];
                Excel.Range rangeB = worksheet1.Range["C10", pBD];
                Excel.Range rangefin = worksheet1.Range[pBG, pBD];
                rangeA.Borders[lgauche].LineStyle = lcontinue;
                rangeA.Borders[lgauche].Weight = llarge;
                rangeB.Borders[ldroite].LineStyle = lcontinue;
                rangeB.Borders[ldroite].Weight = llarge;
                rangefin.Borders[ldroite].LineStyle = lcontinue;
                rangefin.Borders[lbas].Weight = llarge;

                int j = 1;
                foreach (string typo in listtypo.Keys)
                {
                    j += 1;
                    Excel.Range typocell = worksheet1.Cells[4, j];
                    typocell.Value = typo;
                    typocell.Interior.Color = gris;
                    typocell.BorderAround2(lcontinue, lfin);
                    typocell.Borders[ldroite].Weight = llarge;
                    typocell.Borders[lgauche].Weight = llarge;
                    Excel.Range nbrtypocell = worksheet1.Cells[5, j];
                    nbrtypocell.Value = listtypo[typo][0];
                    nbrtypocell.BorderAround2(lcontinue, lfin);
                    nbrtypocell.Borders[ldroite].Weight = llarge;
                    nbrtypocell.Borders[lgauche].Weight = llarge;
                    double moyennetypo = listtypo[typo][1];
                    moyennetypo = Math.Round(moyennetypo, 2);
                    Excel.Range moytypocell = worksheet1.Cells[6, j];
                    moytypocell.Value = moyennetypo;
                    moytypocell.Interior.Color = gris;
                    moytypocell.BorderAround2(lcontinue, lfin);
                    moytypocell.Borders[ldroite].Weight = llarge;
                    moytypocell.Borders[lgauche].Weight = llarge;

                }
                Excel.Range cell = worksheet1.Cells[4, j];
                Excel.Range cell2 = worksheet1.Cells[6, j];
                Excel.Range rangetypohaut = worksheet1.Range["B4", cell];
                Excel.Range rangetypobas = worksheet1.Range["B6", cell2];
                Excel.Range rangetypD = worksheet1.Range[cell, cell2];
                rangetypohaut.Borders[lhaut].LineStyle = lcontinue;
                rangetypohaut.Borders[lhaut].Weight = llarge;
                rangetypobas.Borders[lbas].LineStyle = lcontinue;
                rangetypobas.Borders[lbas].Weight = llarge;
                rangetypD.Borders[ldroite].LineStyle = lcontinue;
                rangetypD.Borders[ldroite].Weight = llarge;

                int k = 9;
                foreach (string piece in autrepieces.Keys)
                {
                    k += 1;
                    Excel.Range piececell = worksheet1.Cells[k, 8];
                    piececell.Value = piece;
                    Excel.Range autrecell = worksheet1.Cells[k, 9];
                    autrecell.Value = autrepieces[piece];
                    if (k%2==0)
                    {
                        piececell.Interior.Color = gris;
                        autrecell.Interior.Color = gris;
                    }
                }               
                string communBG = "H" + k.ToString();
                string communBD = "I" + k.ToString();
                Excel.Range rangecomg = worksheet1.Range["H10", communBG];
                Excel.Range rangecomd = worksheet1.Range["I10", communBD];
                Excel.Range rangecomb = worksheet1.Range[communBG, communBD];
                rangecomg.Borders[lgauche].LineStyle = lcontinue;
                rangecomg.Borders[lgauche].Weight = llarge;
                rangecomd.Borders[ldroite].LineStyle = lcontinue;
                rangecomd.Borders[ldroite].Weight = llarge;
                rangecomb.Borders[lbas].LineStyle = lcontinue;
                rangecomb.Borders[lbas].Weight = llarge;

            } 
            else
            {
                revitui.TaskDialog.Show("Aucun appartement", "Aucun appartement trouvé. L'export des appartements vers Excel ne sera pas réalisé");
            }
            
        }
        // EXPORT PARKINGS
        public static void exportparking(Excel.Application xlapp, Excel.Workbook workbook, Dictionary<string, int> parkings, int sommeplace)
        {
            object misvalue = System.Reflection.Missing.Value;
            Excel.Worksheet worksheet2 = workbook.Sheets.Add(misvalue, workbook.Sheets.Item[1]);
            worksheet2.Columns.ColumnWidth = 25;
            worksheet2.Name = "Parking";
            Excel.XlBordersIndex lgauche = Excel.XlBordersIndex.xlEdgeLeft;
            Excel.XlBordersIndex ldroite = Excel.XlBordersIndex.xlEdgeRight;
            Excel.XlBordersIndex lhaut = Excel.XlBordersIndex.xlEdgeTop;
            Excel.XlBordersIndex lbas = Excel.XlBordersIndex.xlEdgeBottom;
            Excel.XlLineStyle lcontinue = Excel.XlLineStyle.xlContinuous;
            Excel.XlBorderWeight llarge = Excel.XlBorderWeight.xlMedium;
            Excel.XlBorderWeight lfin = Excel.XlBorderWeight.xlThin;
            Excel.XlVAlign centre = Excel.XlVAlign.xlVAlignCenter;
            Excel.XlRgbColor rouge = Excel.XlRgbColor.rgbRed;
            Excel.XlRgbColor sofyacolor = Excel.XlRgbColor.rgbGold;
            Excel.XlRgbColor vert = Excel.XlRgbColor.rgbLightGreen;
            Excel.XlRgbColor gris = Excel.XlRgbColor.rgbLightGrey;
            if (parkings.Count != 0)
            {
                Excel.Range rangetitre1 = worksheet2.Range["A1", "C1"];
                rangetitre1.Merge();
                rangetitre1.Font.Bold = true;
                rangetitre1.Font.Size = 14;
                rangetitre1.Font.Color = rouge;
                rangetitre1.Value = "PARKINGS";
                rangetitre1.VerticalAlignment = centre;

                worksheet2.Range["A4"].Value = "Type de place de parking";
                worksheet2.Range["A4"].BorderAround2(lcontinue, llarge);
                worksheet2.Range["A4"].Interior.Color = sofyacolor;

                worksheet2.Range["B4"].Value = "Nombre de place du type";
                worksheet2.Range["B4"].BorderAround2(lcontinue, llarge);
                worksheet2.Range["B4"].Interior.Color = sofyacolor;

                worksheet2.Range["C4"].Value = "Pourcentage (%)";
                worksheet2.Range["C4"].BorderAround2(lcontinue, llarge);
                worksheet2.Range["C4"].Interior.Color = sofyacolor;
                try
                {
                    int i = 4;
                    foreach (string park in parkings.Keys)
                    {
                        i += 1;
                        Excel.Range nompark = worksheet2.Cells[i, 1];
                        nompark.Value = park;
                        nompark.BorderAround2(lcontinue, lfin);
                        Excel.Range nbrparking = worksheet2.Cells[i, 2];
                        nbrparking.Value = parkings[park];
                        nbrparking.BorderAround2(lcontinue,lfin);
                        double parc = parkings[park];
                        double somme = sommeplace;
                        double pourcent = (parc / somme) * 100;
                        pourcent = Math.Round(pourcent, 2);
                        Excel.Range ratiocell = worksheet2.Cells[i, 3];
                        ratiocell.Value = pourcent;
                        ratiocell.BorderAround2(lcontinue, lfin);
                        if (i % 2 == (0))
                        {
                            nompark.Interior.Color = gris;
                            nbrparking.Interior.Color = gris;
                            ratiocell.Interior.Color = gris;
                        }
                    }
                    string parkBG = "A" + i.ToString();
                    string parkBD = "C" + i.ToString();                  
                    Excel.Range rangeparkD = worksheet2.Range["C4", parkBD];
                    Excel.Range rangeparkB = worksheet2.Range[parkBG, parkBD];
                    rangeparkD.Borders[ldroite].LineStyle = lcontinue;
                    rangeparkD.Borders[ldroite].Weight = llarge;
                    rangeparkB.Borders[lbas].LineStyle = lcontinue;
                    rangeparkB.Borders[lbas].Weight = llarge;
                    //Total place
                    worksheet2.Range["A3"].Value = "Total :";
                    worksheet2.Range["B3"].Value = sommeplace;
                    worksheet2.Range["A3", "B3"].BorderAround2(lcontinue, llarge);
                    worksheet2.Range["A3", "B3"].Interior.Color = vert;
                }

                catch
                {
                    revitui.TaskDialog.Show("Erreur", "Erreur lors de l'export Excel des parkings");
                }
            }
            else
            {
                revitui.TaskDialog.Show("Aucune place de parking modélisée", "L'export des places de parkings vers Excel ne sera pas réalisé.");
                Excel.Range rangetitre1 = worksheet2.Range["A1", "C1"];
                rangetitre1.Merge();
                rangetitre1.Font.Bold = true;
                rangetitre1.Font.Size = 12;
                rangetitre1.Font.Color = rouge;
                rangetitre1.Value = "Aucune places de parking modélisée";
                rangetitre1.VerticalAlignment = centre;
            }
        }

        //  EXPORT GARDE CORPS
        public static void exportgardecorps(Excel.Application xlapp, Excel.Workbook workbook, Dictionary<string, double> gardecorps, Dictionary<int, string> erreursgc)
        {
            object misvalue = System.Reflection.Missing.Value;
            Excel.Worksheet worksheet3 = workbook.Sheets.Add(misvalue, workbook.Sheets.Item[2]);
            worksheet3.Columns.ColumnWidth = 25;
            worksheet3.Name = "Garde-Corps";
            Excel.XlBordersIndex lgauche = Excel.XlBordersIndex.xlEdgeLeft;
            Excel.XlBordersIndex ldroite = Excel.XlBordersIndex.xlEdgeRight;
            Excel.XlBordersIndex lhaut = Excel.XlBordersIndex.xlEdgeTop;
            Excel.XlBordersIndex lbas = Excel.XlBordersIndex.xlEdgeBottom;
            Excel.XlLineStyle lcontinue = Excel.XlLineStyle.xlContinuous;
            Excel.XlBorderWeight llarge = Excel.XlBorderWeight.xlMedium;
            Excel.XlBorderWeight lfin = Excel.XlBorderWeight.xlThin;
            Excel.XlVAlign centre = Excel.XlVAlign.xlVAlignCenter;
            Excel.XlRgbColor rouge = Excel.XlRgbColor.rgbRed;
            Excel.XlRgbColor gold = Excel.XlRgbColor.rgbGold;
            Excel.XlRgbColor vert = Excel.XlRgbColor.rgbLightGreen;
            Excel.XlRgbColor gris = Excel.XlRgbColor.rgbLightGrey;

            Excel.Range rangetitre1 = worksheet3.Range["A1", "C1"];
            rangetitre1.Merge();
            rangetitre1.Value = "GARDE-CORPS";
            rangetitre1.Font.Bold = true;
            rangetitre1.Font.Color = rouge;
            rangetitre1.Font.Size = 14;
            rangetitre1.VerticalAlignment = centre;

            if (gardecorps.Count != 0)
            {

                worksheet3.Range["A3"].Value = "Nombre de type de garde corps";
                worksheet3.Range["A3", "A4"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A3", "A4"].Interior.Color = vert;

                worksheet3.Range["A6"].Value = "Liste des type de garde corps";
                worksheet3.Range["A6", "B6"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A6", "B6"].Interior.Color = gold;

                worksheet3.Range["B6"].Value = "métré (en m)";

                worksheet3.Range["B3"].Value = "métré total (en m)";
                worksheet3.Range["B3", "B4"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["B3", "B4"].Interior.Color = vert;


                int i = 6;
                double sommemetre = 0;
                worksheet3.Cells[4, 1] = gardecorps.Count;
                foreach (string garde in gardecorps.Keys)
                {
                    i += 1;
                    Excel.Range nomgarde = worksheet3.Cells[i, 1];
                    nomgarde.Value = garde;
                    nomgarde.BorderAround2(lcontinue,lfin);
                    Excel.Range metregarde = worksheet3.Cells[i, 2];
                    metregarde.Value = gardecorps[garde];
                    metregarde.BorderAround2(lcontinue, lfin);
                    sommemetre += gardecorps[garde];
                    if (i % 2 == 0)
                    {
                        nomgarde.Interior.Color = gris;
                        metregarde.Interior.Color = gris;
                    }
                }
                worksheet3.Range["B4"].Value = sommemetre;
                string gardeBD = "B" + i.ToString();
                string gardeBG = "A" + i.ToString();
                Excel.Range rangegardeD = worksheet3.Range["B7", gardeBD];
                Excel.Range rangegardeB = worksheet3.Range[gardeBG, gardeBD];
                rangegardeB.Borders[lbas].LineStyle = lcontinue;
                rangegardeB.Borders[lbas].Weight = llarge;
                rangegardeD.Borders[ldroite].LineStyle = lcontinue;
                rangegardeD.Borders[ldroite].Weight = llarge;
            }
            else if (gardecorps.Count==0)
            {
                revitui.TaskDialog.Show("Erreur", "Aucun garde-corps correctement modélisé");
            }
            if (erreursgc.Count != 0 )
            {
                worksheet3.Range["D5"].Value = "ERREURS DE MODELISATION";
                worksheet3.Range["D5"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["D5"].Interior.Color = gold;
                worksheet3.Range["D6"].Value = "ID du type de garde-corps";
                worksheet3.Range["D6"].Interior.Color = gold;
                worksheet3.Range["E6"].Value = "Type du garde corps";
                worksheet3.Range["E6"].Interior.Color = gold;
                worksheet3.Range["D6", "E6"].BorderAround2(lcontinue, llarge);

                int j = 6;
                foreach (int gardeid in erreursgc.Keys)
                {
                    j += 1;
                    Excel.Range cellid = worksheet3.Cells[j, 4];
                    cellid.Value = gardeid;
                    cellid.BorderAround2(lcontinue, lfin);
                    Excel.Range cellerr = worksheet3.Cells[j, 5];
                    cellerr.Value = erreursgc[gardeid];
                    cellerr.BorderAround2(lcontinue, lfin);
                    if (j%2 == 0)
                    {
                        cellid.Interior.Color = gris;
                        cellerr.Interior.Color = gris;
                    }
                }
                string erreurB = "D" + j.ToString();
                string erreurBD = "E" + j.ToString();
                Excel.Range lgarderr = worksheet3.Range["D7", erreurB];
                Excel.Range rangeD = worksheet3.Range["E7", erreurBD];
                rangeD.Borders[ldroite].LineStyle = lcontinue;
                rangeD.Borders[ldroite].Weight = llarge;
                lgarderr.Borders[lgauche].LineStyle = lcontinue;
                lgarderr.Borders[lgauche].Weight = llarge;
                worksheet3.Range[erreurB, erreurBD].Borders[lbas].Weight = llarge;
            }
            else if (gardecorps.Count ==0 & erreursgc.Count==0)
            {
                revitui.TaskDialog.Show("Aucun garde-corps modélisé", "L'export des garde-corps vers Excel ne sera pas réalisé.");
            }
        }

        //EXPORT FACADE VIDE POUR PLEIN
        public static void exportfacadevidepourplein(Excel.Application xlapp, Excel.Workbook workbook, double surfacemur, Dictionary<ElementId, double> dicfenvitre, List<List<string>> nomtypefen, double surfacefenetre, double surfacerideau, double ratio,int comptemur)
        {
            object misvalue = System.Reflection.Missing.Value;
            Excel.Worksheet worksheet4 = workbook.Sheets.Add(misvalue, workbook.Sheets.Item[3]);
            worksheet4.Columns.ColumnWidth = 30;
            worksheet4.Name = "Façade vide pour plein";
            Excel.XlBordersIndex lgauche = Excel.XlBordersIndex.xlEdgeLeft;
            Excel.XlBordersIndex ldroite = Excel.XlBordersIndex.xlEdgeRight;
            Excel.XlBordersIndex lhaut = Excel.XlBordersIndex.xlEdgeTop;
            Excel.XlBordersIndex lbas = Excel.XlBordersIndex.xlEdgeBottom;
            Excel.XlLineStyle lcontinue = Excel.XlLineStyle.xlContinuous;
            Excel.XlBorderWeight llarge = Excel.XlBorderWeight.xlMedium;
            Excel.XlRgbColor gold = Excel.XlRgbColor.rgbGold;
            Excel.XlRgbColor vert = Excel.XlRgbColor.rgbLightGreen;
            Excel.XlRgbColor gris = Excel.XlRgbColor.rgbLightGrey;

            Excel.Range rangetitre1 = worksheet4.Range["A1"];
            rangetitre1.Value = "FACADE VIDE POUR PLEIN";
            rangetitre1.Font.Size = 14;
            rangetitre1.Font.Color = Excel.XlRgbColor.rgbRed;
            rangetitre1.Font.Bold = true;

            Excel.Range rangetitre2 = worksheet4.Range["D1"];
            rangetitre2.Value = "Liste des fenêtres";
            rangetitre2.Font.Bold = true;
            rangetitre2.Interior.Color = gold;
            rangetitre2.BorderAround2(lcontinue, llarge);


            worksheet4.Range["A3"].Value = "Surface de façade ";
            worksheet4.Range["A3"].Font.Bold = true;
            worksheet4.Range["A3", "B3"].BorderAround2(lcontinue, llarge);
            worksheet4.Range["A3", "B3"].Interior.Color = vert;

            worksheet4.Range["A5"].Value = "Surface de façade vitrée ";
            worksheet4.Range["A5"].Font.Bold = true;
            worksheet4.Range["A5", "B5"].BorderAround2(lcontinue, llarge);
            worksheet4.Range["A5", "B5"].Interior.Color = vert;

            worksheet4.Range["A7"].Value = "Ratio";
            worksheet4.Range["A7"].Font.Bold = true;
            worksheet4.Range["A7", "B7"].BorderAround2(lcontinue, llarge);
            worksheet4.Range["A7", "B7"].Interior.Color = vert;

            worksheet4.Range["A10"].Value = "Nombre de murs rideau trouvés ";
            worksheet4.Range["A10"].Font.Bold = true;
            worksheet4.Range["A10", "B10"].BorderAround2(lcontinue, llarge);
            worksheet4.Range["A10", "B10"].Interior.Color = vert;
            worksheet4.Range["B10"].Value = comptemur;

            worksheet4.Range["D2"].Value = "Famille";
            worksheet4.Range["D2"].Font.Bold = true;
            worksheet4.Range["D2"].Interior.Color = gold;
            worksheet4.Range["E2"].Value = "Type";
            worksheet4.Range["E2"].Font.Bold = true;
            worksheet4.Range["E2"].Interior.Color = gold;
            worksheet4.Range["F2"].Value = "Surface";
            worksheet4.Range["F2"].Font.Bold = true;
            worksheet4.Range["F2"].Interior.Color = gold;
            worksheet4.Range["D2", "F2"].BorderAround2(lcontinue, llarge);
            try
            {
                worksheet4.Cells[3, 2] = surfacemur;
                worksheet4.Cells[5, 2] = surfacerideau+surfacefenetre;
                worksheet4.Cells[7, 2] = ratio;

                int y = 2;
                foreach (ElementId wintypeid in dicfenvitre.Keys)
                {
                    y += 1;
                    Excel.Range famcell = worksheet4.Cells[y, 4];
                    famcell.Value= nomtypefen[y - 3][0];
                    Excel.Range typecell = worksheet4.Cells[y, 5];
                    typecell.Value = nomtypefen[y - 3][1];
                    Excel.Range surfcell = worksheet4.Cells[y, 6];
                    surfcell.Value = dicfenvitre[wintypeid];
                    if(y%2==0)
                    {
                        famcell.Interior.Color = gris;
                        typecell.Interior.Color = gris;
                        surfcell.Interior.Color = gris;
                    }
                }
                string fenBG = "D" + y.ToString();
                string fenBD = "F" + y.ToString();
                Excel.Range rangefenG = worksheet4.Range["D3", fenBG];
                Excel.Range rangefenD = worksheet4.Range["F3", fenBD];
                Excel.Range rangefenB = worksheet4.Range[fenBG, fenBD];
                rangefenG.Borders[lgauche].LineStyle = lcontinue;
                rangefenG.Borders[lgauche].Weight = llarge;
                rangefenD.Borders[ldroite].Weight = llarge;
                rangefenD.Borders[ldroite].LineStyle = lcontinue;
                rangefenB.Borders[lbas].LineStyle = lcontinue;
                rangefenB.Borders[lbas].Weight = llarge;
            }
            catch (Exception)
            {

            }
        }

        public static void Exportbureaulot(List<Document> docs, Excel.Application xlapp,Excel.Workbook workbook,Dictionary<string,List<double>>dicoBlot,Dictionary<ElementId,List<double>>dicoBetage,Dictionary<ElementId,Dictionary<string,double>>dicolotlevel)
        {
            object misvalue = System.Reflection.Missing.Value;
            Excel.Worksheet worksheet3 = workbook.Sheets[1];
            worksheet3.Columns.ColumnWidth = 20;
            worksheet3.Name = "Bureaux";
            if (dicoBlot.Count != 0)
            {
                Excel.XlBordersIndex lgauche = Excel.XlBordersIndex.xlEdgeLeft;
                Excel.XlBordersIndex ldroite = Excel.XlBordersIndex.xlEdgeRight;
                Excel.XlBordersIndex lhaut = Excel.XlBordersIndex.xlEdgeTop;
                Excel.XlBordersIndex lbas = Excel.XlBordersIndex.xlEdgeBottom;
                Excel.XlLineStyle lcontinue = Excel.XlLineStyle.xlContinuous;
                Excel.XlBorderWeight llarge = Excel.XlBorderWeight.xlMedium;
                Excel.XlBorderWeight lfin = Excel.XlBorderWeight.xlThin;
                Excel.XlRgbColor gold = Excel.XlRgbColor.rgbGold;
                Excel.XlRgbColor vert = Excel.XlRgbColor.rgbLightGreen;
                Excel.XlRgbColor gris = Excel.XlRgbColor.rgbLightGrey;

                Excel.Range rangetitre1 = worksheet3.Range["A1"];
                rangetitre1.Value = "BUREAUX";
                rangetitre1.Font.Bold = true;
                rangetitre1.Font.Size = 14;
                rangetitre1.Font.Color = Excel.XlRgbColor.rgbRed;

                worksheet3.Range["A3"].Value = "Surface totale de bureau";
                worksheet3.Range["A3", "B3"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A3", "B3"].Interior.Color = vert;

                worksheet3.Range["A5"].Value = "Surface totale de salle de réunion";
                worksheet3.Range["A5", "B5"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A5", "B5"].Interior.Color = vert;

                worksheet3.Range["A7"].Value = "Ratio REU/BUR";
                worksheet3.Range["A7", "B7"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A7", "B7"].Interior.Color = vert;

                worksheet3.Range["D3"].Value = "Nombre total de lot";
                worksheet3.Range["D3", "D4"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["D3", "D4"].Interior.Color = vert;
                worksheet3.Range["D6"].Value = "Surface moyenne d'un lot";
                worksheet3.Range["D6", "D7"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["D6", "D7"].Interior.Color = vert;

                //SURFACE PAR ETAGE
                worksheet3.Range["A13"].Value = "SURFACE PAR ETAGE";
                worksheet3.Range["A13"].Font.Bold = true;
                worksheet3.Range["A13"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A13"].Interior.Color = gold;
                worksheet3.Range["A14"].Value = "ETAGE";
                worksheet3.Range["A14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["B14"].Value = "BUREAUX";
                worksheet3.Range["B14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["C14"].Value = "REUNIONS";
                worksheet3.Range["C14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["D14"].Value = "RATIO";
                worksheet3.Range["D14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["A14", "D14"].Interior.Color = gold;

                int i = 14;
                int nombrelot = dicoBlot.Keys.Count;
                worksheet3.Cells[4, 4] = nombrelot;
                double sommebureaux = 0;
                double sommereu = 0;
                foreach (ElementId etageid in dicoBetage.Keys)
                {
                    i += 1;
                    string nometage = "";
                    //on recupere le nom de l'etage en fonction de son document
                    foreach (Document doc in docs)
                    {
                        try
                        {
                            nometage = doc.GetElement(etageid).Name;
                            break;
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    Excel.Range etagecell = worksheet3.Cells[i, 1];
                    etagecell.Value = nometage;
                    etagecell.BorderAround2(lcontinue, lfin);
                    Excel.Range burcell = worksheet3.Cells[i, 2];
                    burcell.Value = dicoBetage[etageid][0];
                    burcell.BorderAround2(lcontinue, lfin);
                    Excel.Range reucell = worksheet3.Cells[i, 3];
                    reucell.Value = dicoBetage[etageid][1];
                    reucell.BorderAround2(lcontinue, lfin);
                    sommebureaux += dicoBetage[etageid][0];
                    sommereu += dicoBetage[etageid][1];
                    double ratiot = dicoBetage[etageid][1] / dicoBetage[etageid][0];
                    ratiot = Math.Round(ratiot, 2);
                    Excel.Range ratiocell = worksheet3.Cells[i, 4];
                    ratiocell.Value = ratiot;
                    ratiocell.BorderAround2(lcontinue, lfin);
                    string setageMG = "A" + i.ToString();
                    string setageMD = "D" + i.ToString();
                    Excel.Range rangeetageM = worksheet3.Range[setageMG, setageMD];
                    rangeetageM.Borders[lbas].LineStyle = lcontinue;
                    rangeetageM.Borders[lbas].Weight = llarge;
                    if ( i%2 ==0)
                    {
                        etagecell.Interior.Color = gris;
                        burcell.Interior.Color = gris;
                        reucell.Interior.Color = gris;
                        ratiocell.Interior.Color = gris;
                    }
                }
                //TOTAUX
                worksheet3.Range["B3"].Value = sommebureaux;
                worksheet3.Range["B5"].Value = sommereu;
                double ratiooo = (sommereu / sommebureaux);
                ratiooo = Math.Round(ratiooo, 2);
                worksheet3.Range["B7"].Value = ratiooo;
                string setageMGB = "B" + i.ToString();
                string setageBD = "D" + i.ToString();
                Excel.Range rangeetageMGV = worksheet3.Range["B15", setageMGB];
                Excel.Range rangeetageD = worksheet3.Range["D15", setageBD];
                rangeetageD.Borders[ldroite].LineStyle = lcontinue;
                rangeetageD.Borders[ldroite].Weight = llarge;
                rangeetageD.Borders[lgauche].LineStyle = lcontinue;
                rangeetageMGV.Borders[lgauche].LineStyle = lcontinue;
                rangeetageMGV.Borders[lgauche].LineStyle = lcontinue;

                //SURFACE PAR LOT
                Excel.Range rangetitre2 = worksheet3.Range["F13"];
                rangetitre2.Value = "SURFACE PAR LOT";
                rangetitre2.BorderAround2(lcontinue, llarge);
                rangetitre2.Interior.Color = gold;
                rangetitre2.Font.Bold = true;
                worksheet3.Range["F14"].Value = "NOM LOT";
                worksheet3.Range["F14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["G14"].Value = "SURFACE LOT";
                worksheet3.Range["G14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["H14"].Value = "BUREAUX";
                worksheet3.Range["H14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["I14"].Value = "REUNIONS";
                worksheet3.Range["I14"].BorderAround2(lcontinue, llarge);
                worksheet3.Range["F14", "I14"].Interior.Color = gold;

                int j = 14;
                double sommetotlot = 0;
                foreach (string lot in dicoBlot.Keys)
                {
                    j += 1;
                    Excel.Range nomlot = worksheet3.Cells[j, 6];
                    nomlot.Value = lot;
                    nomlot.BorderAround2(lcontinue, lfin);                    
                    Excel.Range surlot = worksheet3.Cells[j, 7];
                    surlot.Value = dicoBlot[lot][0];
                    surlot.BorderAround2(lcontinue, lfin);
                    Excel.Range surbur = worksheet3.Cells[j, 8];
                    surbur.Value = dicoBlot[lot][1];
                    surbur.BorderAround2(lcontinue, lfin);
                    Excel.Range surreu = worksheet3.Cells[j, 9];
                    surreu.Value = dicoBlot[lot][2];
                    surreu.BorderAround2(lcontinue, lfin);
                    sommetotlot += dicoBlot[lot][0];
                    if (j%2==0)
                    {
                        nomlot.Interior.Color = gris;
                        surlot.Interior.Color = gris;
                        surbur.Interior.Color = gris;
                        surreu.Interior.Color = gris;
                    }
                }
                double moyennelot = (sommetotlot / nombrelot);
                moyennelot = Math.Round(moyennelot, 2);
                worksheet3.Range["D7"].Value = moyennelot;
                string lotBG = "F" + j.ToString();
                string lotH = "H" + j.ToString();
                string lotI = "I" + j.ToString();
                Excel.Range rangelotG = worksheet3.Range["F15", lotBG];
                Excel.Range rangelotH = worksheet3.Range["H15", lotH];
                Excel.Range rangelotI = worksheet3.Range["I15", lotI];
                rangelotI.Borders[ldroite].LineStyle = lcontinue;
                rangelotI.Borders[ldroite].Weight = llarge;
                rangelotH.Borders[ldroite].LineStyle = lcontinue;
                rangelotH.Borders[lgauche].LineStyle = lcontinue;
                rangelotG.Borders[lgauche].LineStyle = lcontinue;
                rangelotG.Borders[lgauche].Weight = llarge;
                rangelotG.Borders[ldroite].LineStyle = lcontinue;

                //NOMBRE DE LOT PAR ETAGE
                Excel.Range rangetitre3 = worksheet3.Range["K14","L14"];
                rangetitre3.Merge();
                rangetitre3.Value = "SURFACE DE LOT PAR ETAGE (m²)";
                rangetitre3.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                rangetitre3.BorderAround2(lcontinue, llarge);
                rangetitre3.Font.Bold = true;
                rangetitre3.Interior.Color = gold;

                int maxL = 0;
                int k = 13;
                foreach (ElementId etageid in dicolotlevel.Keys)
                {
                    k += 2;
                    string nometage = "";
                    //on recupere le nom de l'etage en fonction de son document
                    foreach (Document doc in docs)
                    {
                        try
                        {
                            nometage = doc.GetElement(etageid).Name;
                            break;
                        }
                        catch
                        {
                            continue;
                        }
                    }
                    Excel.Range etagecell = worksheet3.Cells[k, 11];
                    etagecell.Value = nometage;
                    Excel.Range nbrlotcell = worksheet3.Cells[k + 1, 11];
                    nbrlotcell.Value = dicolotlevel[etageid].Count.ToString() + " lots";
                    Excel.Range countcell = worksheet3.Range[worksheet3.Cells[k, 11], worksheet3.Cells[k+1, 11]];
                    countcell.BorderAround2(lcontinue, lfin);

                    int L = 11;
                    foreach (string nomlot in dicolotlevel[etageid].Keys)
                    {
                        L += 1;
                        Excel.Range lotcell = worksheet3.Cells[k, L];
                        lotcell.Value = nomlot;
                        Excel.Range surlotcell = worksheet3.Cells[k + 1, L];
                        surlotcell.Value = dicolotlevel[etageid][nomlot];
                        worksheet3.Range[lotcell, surlotcell].BorderAround2(lcontinue, lfin);
                    }
                    if (maxL < L)
                    {
                        maxL = L;
                    }
                }
                int x = 14;
                while (x < k)
                {
                    x += 1;
                    Excel.Range cellHG = worksheet3.Cells[x, 11];
                    Excel.Range lotetatD = worksheet3.Cells[x, maxL];
                    Excel.Range rangelotetatMH = worksheet3.Range[cellHG, lotetatD];
                    rangelotetatMH.Interior.Color = Excel.XlRgbColor.rgbLightGrey;
                    x += 1;
                    cellHG = worksheet3.Cells[x, 11];
                    lotetatD = worksheet3.Cells[x,maxL ];
                    rangelotetatMH = worksheet3.Range[cellHG, lotetatD];
                    rangelotetatMH.Borders[lbas].LineStyle = lcontinue;
                    rangelotetatMH.Borders[lbas].Weight = llarge;
                }
                int y = 10;
                while (y<maxL)
                {
                    y += 1;
                    Excel.Range cellhaute = worksheet3.Cells[15, y];
                    Excel.Range cellbasse = worksheet3.Cells[k+1, y];
                    Excel.Range rangehautbas = worksheet3.Range[cellhaute, cellbasse];
                    rangehautbas.Borders[ldroite].LineStyle = lcontinue;
                }
                string lotetaBG = "K" + (k+1).ToString();
                Excel.Range lotetatHD = worksheet3.Cells[15, maxL];
                Excel.Range lotetatBD = worksheet3.Cells[(k+1), maxL];
                Excel.Range rangeetatD = worksheet3.Range[lotetatHD, lotetatBD];
                Excel.Range rangelotetaG = worksheet3.Range["K15", lotetaBG];
                Excel.Range rangelotetatH = worksheet3.Range["K15", lotetatHD];
                Excel.Range rangelotetatB = worksheet3.Range[lotetaBG, lotetatBD];
                rangelotetatB.Borders[lbas].LineStyle = lcontinue;
                rangelotetatB.Borders[lbas].Weight = llarge;
                rangelotetatH.Borders[lhaut].LineStyle = lcontinue;
                rangelotetatH.Borders[lhaut].Weight = llarge;
                rangelotetaG.Borders[lgauche].LineStyle = lcontinue;
                rangelotetaG.Borders[lgauche].Weight = llarge;
                rangeetatD.Borders[ldroite].LineStyle = lcontinue;
                rangeetatD.Borders[ldroite].Weight = llarge;
            }
            else
            {
                revitui.TaskDialog.Show("Erreur","Aucune pièce 'bureau' trouvée. L'export des lots et bureaux vers Excel ne sera pas réalisé.");
            }
            
        }
        

        public static void OpenExcel(Excel.Application xlapp, Excel.Workbook workbook, string path)
        {
            try
            {
                revitui.TaskDialog.Show("Export Excel", "Sauvegarde des données");
                object misvalue = System.Reflection.Missing.Value;
                workbook.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook, misvalue, misvalue, misvalue, misvalue, Excel.XlSaveAsAccessMode.xlExclusive, misvalue, misvalue, misvalue, misvalue);
                xlapp.Visible = true;

            }
            catch
            {
                revitui.TaskDialog.Show("Sauvegarde annulée", "La sauvegarde a été annulé");
            }
        }
    }
}