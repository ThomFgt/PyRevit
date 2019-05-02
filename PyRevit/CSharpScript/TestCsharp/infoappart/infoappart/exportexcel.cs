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

namespace infoappart
{
    class exportexcel
    {
        public string path = @"C:\Users\JulesSablé\Desktop\RECUPINFOPIECE";
        public Dictionary<string, List<double>> dico1;
        public Dictionary<string, double> dico2;
        public List<string> typelog;

        public exportexcel(Dictionary<string, List<double>> dic1, Dictionary<string,double> dic2, List<string> typelogement )
        {
            this.dico1 = dic1;
            this.dico2 = dic2;
            this.typelog = typelogement;
            Excel.Application xlapp = new Excel.Application();

            try
            {
                if (xlapp == null)
                {
                    revitui.TaskDialog.Show("erreur", "excel pas installe");
                }

                object misvalue = System.Reflection.Missing.Value;
                xlapp.DisplayAlerts = false;
                Excel.Workbook workbook = xlapp.Workbooks.Add(misvalue);
                Excel.Worksheet worksheet1 = workbook.Sheets.Item[1];
                worksheet1.Name = "Information Appartement";
                Excel.Worksheet worksheet2 = workbook.Sheets.Add(misvalue, worksheet1, misvalue, misvalue);
                worksheet2.Name = "Pièces non attribuées";
                //xlapp.Sheets.add(Type:"template.xlsx");
                revitui.TaskDialog.Show("exportexcel","app excel ouverte et feuillles creee");
                int ligne1 = 2;
                foreach (string key in dico1.Keys)
                {
                    worksheet1.Cells[ligne1, 1] = key;
                    worksheet1.Cells[ligne1, 3] = dico1[key][0];
                    worksheet1.Cells[ligne1, 4] = dico1[key][1];
                    worksheet1.Cells[ligne1, 5] = dico1[key][2];
                    worksheet1.Cells[ligne1, 6] = dico1[key][3];
                    worksheet1.Cells[ligne1, 7] = dico1[key][4];
                    ligne1++;
                }
                int ligne3 = 2;
                foreach (string typlog in typelog)
                {
                    worksheet1.Cells[ligne3, 2] = typlog;
                    ligne3++;
                }
                int ligne2 = 2;
                foreach (string key in dico2.Keys)
                {
                    worksheet2.Cells[ligne2, 1] = key;
                    worksheet2.Cells[ligne2, 2] = dico2[key];
                    ligne2++;
                }
                revitui.TaskDialog.Show("exportexcel", "att sauvegarde");
                workbook.SaveAs(path, Excel.XlFileFormat.xlWorkbookNormal, misvalue, misvalue, misvalue, misvalue, Excel.XlSaveAsAccessMode.xlExclusive, misvalue, misvalue, misvalue, misvalue, misvalue);
                workbook.Close(true, misvalue, misvalue);
                xlapp.Quit();
                Marshal.ReleaseComObject(worksheet1);
                Marshal.ReleaseComObject(worksheet2);
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