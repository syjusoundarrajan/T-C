using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    class FloorOpening : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\WerkStudent\ppt\Fr 160819\Gen_Open_Fin");

            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            double cF = 1 / 0.3048;
            //Get UI document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get document
            Document doc = uidoc.Document;

            string Level_1 = (String)xlRange.Cells[4, "A"].Value2;
            //Create levels
            Level Level1 = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType().Cast<Level>().First(x => x.Name == Level_1);
            Level Level2 = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType().Cast<Level>().First(x => x.Name == "Level 2");

            string ty1 = (String)xlRange.Cells[3, "A"].Value2;

            var FloorType = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors)
               .WhereElementIsElementType().Cast<FloorType>().First(x => x.Name == ty1);
            //var Floor = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FloorsCut)
            //   .WhereElementIsElementType().Cast<Floor>().First(x => x.Name == ty1);

            int points = Convert.ToInt32(xlRange.Cells[2, "C"].value2);

            List<XYZ> pointlist = new List<XYZ>();
            for(int i = 5; i <= points + 4; i++)
            {
                XYZ pts = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                pointlist.Add(pts);
            }
            List<Line> concurves = new List<Line>();
            for (int i = 0; i < pointlist.Count; i++)
            {
                if( i<pointlist.Count-1)
                {
                    Line Alllines = Line.CreateBound(pointlist.ElementAt(i), pointlist.ElementAt(i + 1));
                    concurves.Add(Alllines);
                }
                else
                {
                    Line Alllines = Line.CreateBound(pointlist.ElementAt(pointlist.Count - 1), pointlist.ElementAt(0));
                    concurves.Add(Alllines);
                }
            }
            List<Curve> licurve = new List<Curve>();
            for (int i = 0; i < concurves.Count; i++)
            {
                licurve.Add(concurves.ElementAt(i));
            }
            // Making a curve in a loop
            CurveLoop crvloop = CurveLoop.Create(licurve);
            CurveArray curArr = new CurveArray();
            foreach (Curve c in crvloop)
            {
                //Put the curves to curve array
                curArr.Append(c);
            }

            List<XYZ> pointlist0 = new List<XYZ>();
            for (int i = 5; i <= points + 4; i++)
            {
                XYZ pts0 = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                pointlist0.Add(pts0);
            }
            List<Line> concurves0 = new List<Line>();
            for (int i = 0; i < pointlist0.Count; i++)
            {
                if (i < pointlist0.Count - 1)
                {
                    Line Alllines0 = Line.CreateBound(pointlist0.ElementAt(i), pointlist0.ElementAt(i + 1));
                    concurves0.Add(Alllines0);
                }
                else
                {
                    Line Alllines0 = Line.CreateBound(pointlist0.ElementAt(pointlist0.Count - 1), pointlist0.ElementAt(0));
                    concurves0.Add(Alllines0);
                }
            }
            List<Curve> licurve0 = new List<Curve>();
            for (int i = 0; i < concurves0.Count; i++)
            {
                licurve0.Add(concurves0.ElementAt(i));
            }
            // Making a curve in a loop
            CurveLoop crvloop0 = CurveLoop.Create(licurve0);
            CurveArray curArr0 = new CurveArray();
            foreach (Curve c0 in crvloop0)
            {
                //Put the curves to curve array
                curArr0.Append(c0);
            }


            List<XYZ> opointlist = new List<XYZ>();
            for (int i = 3; i <= 6; i++)
            {
                XYZ opoints = new XYZ(xlRange.Cells[i, "H"].value2 * cF, xlRange.Cells[i, "I"].value2 * cF, xlRange.Cells[i, "J"].value2 * cF);
                opointlist.Add(opoints);
            }
            List<Line> opline = new List<Line>();
            for (int i = 0; i < opointlist.Count; i++)
            {
                if (i < opointlist.Count - 1)
                {

                    Line oline = Line.CreateBound(opointlist.ElementAt(i), opointlist.ElementAt(i + 1));
                    opline.Add(oline);
                }
                else
                {
                    Line oline = Line.CreateBound(opointlist.ElementAt(3), opointlist.ElementAt(0));
                    opline.Add(oline);
                }
            }
            List<Curve> ocurve = new List<Curve>();
            for (int i = 0; i < opline.Count; i++)
            {
                ocurve.Add(opline.ElementAt(i));
            }
            //CurveLoop crvvloopopen = CurveLoop.Create(ocurve);
            CurveArray curvopen = new CurveArray();
            foreach (Curve co in ocurve)
            {
                //Put the curves to curve array
                curvopen.Append(co);
            }

            Floor fl;
            Floor f0;
            
            try
            {
                using (Transaction trans = new Transaction(doc, "Bungalow Floors"))
                {
                    trans.Start();
                    fl = doc.Create.NewFloor(curArr, FloorType, Level1, false);
                    f0 = doc.Create.NewFloor(curArr0, FloorType, Level1, false);
                    trans.Commit();
                    
                    using (Transaction trans1 = new Transaction(doc, "Bung open"))
                    {
                       
                        trans1.Start();
                        //var Open = doc.Create.NewFloor(curArr1, FloorType, Level2, false);
                        doc.Create.NewOpening(fl, curvopen, true);
                        trans1.Commit();
                    }
                }
                

                return Result.Succeeded;
            }
            catch (Exception somethingwrong)
            {
                message = somethingwrong.Message;
                return Result.Failed;
            }
           

        }

    }
}