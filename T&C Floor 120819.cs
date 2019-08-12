using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using Microsoft.Office.Interop.Excel;
using Line = Autodesk.Revit.DB.Line;    
//using Floor = Autodesk.Revit.DB.Floor;

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    class WithNoOfPoints : IExternalCommand
    {


        public  Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\WerkStudent\Updated\General_WithOpening2");

            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            double cF = 1 / 0.3048;
            //Get UI document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get document
            Document doc = uidoc.Document;

            string level = (String)xlRange.Cells[4, "A"].Value2;
            //Create levels
            Level levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType().Cast<Level>().First(x => x.Name == level);

            string ty1 = (String)xlRange.Cells[3, "A"].Value2;

            var FloorType = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors)
               .WhereElementIsElementType().Cast<FloorType>().First(x => x.Name == ty1);

                      
            int points = Convert.ToInt32(xlRange.Cells[2, "B"].value2);
            

            List<XYZ> pointlist1 = new List<XYZ>();

            for (int i = 5; i <= points + 4; i = i + 2)
            {

                XYZ pisb = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);

                pointlist1.Add(pisb);

            }
            List<XYZ> pointlist2 = new List<XYZ>();
            for (int i = 6; i <= points + 4; i = i + 2)
            {

                XYZ pisf = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                pointlist2.Add(pisf);
            }

            List<Line> concurves = new List<Line>();

            for (int i = 0; i <= (points - 2) / 2; i++)
            {
                XYZ pisb = pointlist1.ElementAt(i);
                XYZ pisf = pointlist2.ElementAt(i);
                Line Alllines = Line.CreateBound(pisb, pisf);
                concurves.Add(Alllines);

            }
            //Line close = Line.CreateBound(pointlist2.ElementAt(points -6), pointlist1.ElementAt(0));
           // concurves.Add(close);

            List<Curve> licurve = new List<Curve>();
            for(int i = 0; i < concurves.Count; i++)
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

            try
            {
                using (Transaction trans = new Transaction(doc, "Bungalow Floors"))
                {
                    trans.Start();
                    doc.Create.NewFloor(curArr, FloorType, levels, false);     
                    trans.Commit();
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


