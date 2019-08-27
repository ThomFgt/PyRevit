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


namespace infoappart
{
    class Hsfp
    {
        public static double Gethsfp(Document doc,Room piece, ElementCategoryFilter filtreplafond)
        {
            double hsfp = 0.00;

            // ON CALCULE LA HAUTEUR SOUS PLAFOND OU DALLE
            try
            {
                LocationPoint locpiece = piece.Location as LocationPoint;
                XYZ ptpiece = locpiece.Point;
                ptpiece = new XYZ(ptpiece.X, ptpiece.Y, ptpiece.Z + 0.05);
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
            return hsfp;        
        }

        public static double Gethsfs (Document doc,Room piece, ElementCategoryFilter filtresol)
        {
            double hsfs = 0.00;
            // on verifie si sol au dessus de piece
            try
            {
                LocationPoint locpiece = piece.Location as LocationPoint;
                XYZ ptpiece = locpiece.Point;
                ptpiece = new XYZ(ptpiece.X, ptpiece.Y, ptpiece.Z + 0.05);
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
            return hsfs;
        }

        public static double hsfpmoy (List<double> listhsfp)
        {
            listhsfp.RemoveAll(x => x == 0.00);
            if (listhsfp.Count==0)
            {
                double moyennehsfp = 0.00;
                return moyennehsfp;
            }
            else
            {
                 double moyennehsfp = listhsfp.Average();
                return moyennehsfp;
            }
            
            
        }
    }
}



   