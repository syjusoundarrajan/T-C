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
    class Opening : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\WerkStudent\Updated\Floor100719");

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


            string ty1 = (String)xlRange.Cells[3, "A"].Value2;

            var FloorType = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors)
               .WhereElementIsElementType().Cast<FloorType>().First(x => x.Name == ty1);


            int points = Convert.ToInt32(xlRange.Cells[2, "C"].value2);

            //List<double> zLi = new List<double>();
            //for (int i = 5; i < points + 4; i++)
            //{
            //    double zPoints = xlRange.Cells[i, "D"].value2;
            //    zLi.Add(zPoints);
            //}

            List<XYZ> pointlist1 = new List<XYZ>();
            List<XYZ> pointlist3 = new List<XYZ>();


            for (int i = 5; i <= points + 4; i = i + 2)
            {
                if (xlRange.Cells[i, "D"].value2 == 0)
                {
                    XYZ pisb = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);

                    pointlist1.Add(pisb);

                }
                else
                {

                    XYZ pisb = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);

                    pointlist3.Add(pisb);
                }

            }

            List<XYZ> pointlist2 = new List<XYZ>();
            List<XYZ> pointlist4 = new List<XYZ>();

            for (int i = 6; i <= points + 4; i = i + 2)
            {
                if (xlRange.Cells[i, "D"].value2 == 0)
                {
                    XYZ pisf = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                    pointlist2.Add(pisf);
                }
                else
                {
                    XYZ pisf = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                    pointlist4.Add(pisf);
                }
            }



            List<Line> concurves = new List<Line>();

            for (int i = 0; i < pointlist1.Count; i++)
            {
                XYZ pisb = pointlist1.ElementAt(i);
                XYZ pisf = pointlist2.ElementAt(i);
                Line Alllines = Line.CreateBound(pisb, pisf);
                concurves.Add(Alllines);

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

            List<Line> concurves1 = new List<Line>();

            for (int i = 0; i < pointlist3.Count; i++)
            {
                XYZ pisb = pointlist3.ElementAt(i);
                XYZ pisf = pointlist4.ElementAt(i);
                Line Alllines1 = Line.CreateBound(pisb, pisf);
                concurves1.Add(Alllines1);

            }

            List<Curve> licurve1 = new List<Curve>();
            for (int i = 0; i < concurves1.Count; i++)
            {
                licurve1.Add(concurves1.ElementAt(i));
            }
            // Making a curve in a loop
            CurveLoop crvloop1 = CurveLoop.Create(licurve1);
            CurveArray curArr1 = new CurveArray();

            XYZ p1 = new XYZ(xlRange.Cells[3, "H"].value2 * cF, xlRange.Cells[3, "I"].value2 * cF, xlRange.Cells[3, "J"].value2 * cF);
            XYZ p2 = new XYZ(xlRange.Cells[4, "H"].value2 * cF, xlRange.Cells[4, "I"].value2 * cF, xlRange.Cells[4, "J"].value2 * cF);
            XYZ p3 = new XYZ(xlRange.Cells[5, "H"].value2 * cF, xlRange.Cells[5, "I"].value2 * cF, xlRange.Cells[5, "J"].value2 * cF);
            XYZ p4 = new XYZ(xlRange.Cells[6, "H"].value2 * cF, xlRange.Cells[6, "I"].value2 * cF, xlRange.Cells[6, "J"].value2 * cF);

            Line l1 = Line.CreateBound(p1, p2);
            Line l2 = Line.CreateBound(p2, p3);
            Line l3 = Line.CreateBound(p3, p4);
            Line l4 = Line.CreateBound(p4, p1);

            List<Curve> open = new List<Curve>();
            open.Add(l1);
            open.Add(l2);
            open.Add(l3);
            open.Add(l4);
            CurveArray curopen = new CurveArray();
            foreach (Curve co in open)
            {
                //Put the curves to curve array
                curopen.Append(co);
            }

            foreach (Curve c1 in crvloop1)
            {
                //Put the curves to curve array
                curArr1.Append(c1);
            }
            Floor FLev1;
            Floor FLev2;
            try
            {
                using (Transaction trans = new Transaction(doc, "Bungalow Floors"))
                {
                    trans.Start();


                    FLev1 = doc.Create.NewFloor(curArr, FloorType, Level1, false);
                    FLev2 = doc.Create.NewFloor(curArr1, FloorType, Level1, false);
                    
                    trans.Commit();
                }
                using (Transaction trans1 = new Transaction(doc, "Bungalow Floors"))
                {
                    trans1.Start();


                    doc.Create.NewOpening(FLev2, curopen, false);


                    trans1.Commit();
                }
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


