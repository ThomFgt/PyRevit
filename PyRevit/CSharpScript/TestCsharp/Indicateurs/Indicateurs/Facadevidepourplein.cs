using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;


namespace Indicateurs
{
    class Facadevidepourplein
    {
        public static (double,int,double) facadeopaque (List<Document> docs)
        {
            double surfacemurext = 0.00;          
            double surfacepanneaux = 0;
            int comptemur = 0;
            try
            {
                foreach (Document doc in docs)
                {
                    IList<Element> murs = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToList(); ;
                    foreach (Element mur in murs)
                    {
                        double surfacedumur = mur.LookupParameter("Surface").AsDouble();
                        surfacedumur = UnitUtils.ConvertFromInternalUnits(surfacedumur, DisplayUnitType.DUT_SQUARE_METERS);
                        try
                        {
                            Wall wall = mur as Wall;
                            WallType murtype = wall.WallType;
                            
                            if (wall.WallType.Kind == WallKind.Curtain & murtype.Function == WallFunction.Exterior)
                            {
                                surfacemurext += surfacedumur;
                                comptemur += 1;
                                IList<ElementId> panelids = wall.CurtainGrid.GetPanelIds().ToList();
                                foreach (ElementId panelid in panelids)
                                {
                                    Element panel = doc.GetElement(panelid);
                                    List<ElementId> matids = panel.GetMaterialIds(false).ToList();
                                    foreach (ElementId matid in panelids)
                                    {
                                        Material panmat = doc.GetElement(panelid) as Material;
                                        if (panmat.Transparency > 30)
                                        {
                                            double surfacepanel = panel.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                                            surfacepanel = UnitUtils.ConvertFromInternalUnits(surfacepanel, DisplayUnitType.DUT_SQUARE_METERS);
                                            surfacepanneaux += surfacepanel;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch
            {

            }
            surfacemurext = Math.Round(surfacemurext, 2);
            surfacepanneaux = Math.Round(surfacepanneaux, 2);
            return (surfacemurext,comptemur,surfacepanneaux);
        }
            
        
        public static (Dictionary<ElementId,double>,List<List<string>>) surfacefenetrevitre(List<Document> docs, Options option)
        {
            Dictionary<ElementId, double> surfacevitretypefen = new Dictionary<ElementId, double>();
            List<ElementId> fennonvitre = new List<ElementId>();
            List<List<string>> nomtypefen = new List<List<string>>();
            foreach (Document doc in docs)
            {
                IList<Element> windows = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsNotElementType().ToElements();
                IList<ElementId> windowtypeids = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Windows).WhereElementIsElementType().ToElementIds().ToList();

                foreach (ElementId windowtypeid in windowtypeids)
                {
                    ElementType wintype = doc.GetElement(windowtypeid) as ElementType;
                    string famwinname = wintype.FamilyName;
                    string fentypename = wintype.Name;
                    List<ElementId> fenids = wintype.GetMaterialIds(false).ToList();
                    foreach (ElementId fenid in fenids)
                    {                       
                        Material femat = doc.GetElement(fenid) as Material;
                        if (femat.Transparency > 30)
                        {                           
                            List<string> famtypelist = new List<string> { famwinname, fentypename };
                            nomtypefen.Add(famtypelist);
                            surfacevitretypefen.Add(windowtypeid, 0);
                            break;
                        }
                        else
                        {
                            fennonvitre.Add(windowtypeid);
                        }
                    }
                }
                foreach(Element window in windows)
                {
                    if (surfacevitretypefen.ContainsKey(window.GetTypeId()))
                    {
                        double surfacewin = window.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                        surfacewin = UnitUtils.ConvertFromInternalUnits(surfacewin, DisplayUnitType.DUT_SQUARE_METERS);
                        surfacewin = Math.Round(surfacewin, 2);
                        surfacevitretypefen[window.GetTypeId()] += surfacewin;
                    }
                }             
            }
            return (surfacevitretypefen, nomtypefen);
        }

    }
}
