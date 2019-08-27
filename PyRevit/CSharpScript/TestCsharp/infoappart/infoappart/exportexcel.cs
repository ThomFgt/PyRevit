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

namespace infoappart
{
    class Exportexcel
    {

        
        public Exportexcel (Excel.Application xlapp, Dictionary<string, List<Room>> DicAppart,List<string> typo, List<double> shabs,List<double> hsfp,List<int> nbrfen)

        {
            
            string path = @"C:\Users\JulesSablé\Desktop\Helloworld3";
            object misvalue = System.Reflection.Missing.Value;          
            Excel.Workbook workbook = xlapp.Workbooks.Add(misvalue);
            Excel.Worksheet worksheet1 = workbook.Sheets.Item[1];
            worksheet1.Name = "Information Appartement";

            try
            {
                if (xlapp == null)
                {
                    revitui.TaskDialog.Show("erreur", "excel pas installe");
                }

                //Excel.Worksheet worksheet2 = workbook.Sheets.Add(misvalue, worksheet1, misvalue, misvalue);
                //worksheet2.Name = "Pièces Communes ou non attribuées";
                int i = 2;
                foreach (string appart in DicAppart.Keys)
                {
                    worksheet1.Cells[1, 1] = "numero appart";
                    worksheet1.Cells[i, 1] = appart;
                    worksheet1.Cells[1, 2] = "nombre de piece";
                    worksheet1.Cells[i, 2] = DicAppart[appart].Count;
                    worksheet1.Cells[1, 3] = "Type de logement";
                    worksheet1.Cells[i, 3] = typo[i - 2];
                    worksheet1.Cells[1, 4] = "surface habitable";
                    worksheet1.Cells[i, 4] = shabs[i - 2];
                    worksheet1.Cells[1, 5] = "hauteur sous plafond";
                    worksheet1.Cells[i, 5] = hsfp[i - 2];
                    worksheet1.Cells[1, 6] = "nombre de fenetre";
                    worksheet1.Cells[i, 6] = nbrfen[i - 2];
                    i += 1;
                }
                
                revitui.TaskDialog.Show("exportexcel", "att sauvegarde");
                workbook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misvalue, misvalue, misvalue, misvalue, Excel.XlSaveAsAccessMode.xlExclusive, misvalue, misvalue, misvalue, misvalue, misvalue);
                workbook.Close(true, misvalue, misvalue);
                xlapp.Quit();
                Marshal.ReleaseComObject(worksheet1);
                //Marshal.ReleaseComObject(worksheet2);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(xlapp);
            }
            catch 
            {
                revitui.TaskDialog.Show("erreur", "gros probleme");
            }

        }

    }
}