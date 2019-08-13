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
    class FloorCutOut : IExternalCommand
    {


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\WerkStudent\Updated\General_WithOpening");

            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[3];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //Get UI document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get document
            Document doc = uidoc.Document;

            string level = (String)xlRange.Cells[3, "A"].Value2;
            //Create levels
            Level levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType().Cast<Level>().First(x => x.Name == level);

            string ty1 = (String)xlRange.Cells[2, "A"].Value2;

            var FloorType = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors)
               .WhereElementIsElementType().Cast<FloorType>().First(x => x.Name == ty1);

            //string ty2 = (String)xlRange.Cells[2, "B"].Value2;
            ////var opening =sourceFloor.Document.Create.NewOpening(destFloor,openingCurveArray,true);
            //var Opening = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_FloorOpening)
            //    .WhereElementIsElementType().Cast<Opening>();




            double p01x = xlRange.Cells[4, "B"].value2;
            double p01y = xlRange.Cells[4, "C"].value2;
            double p01z = xlRange.Cells[4, "D"].value2;

            double p02x = xlRange.Cells[5, "B"].value2;
            double p02y = xlRange.Cells[5, "C"].value2;
            double p02z = xlRange.Cells[5, "D"].value2;

            double p03x = xlRange.Cells[6, "B"].value2;
            double p03y = xlRange.Cells[6, "C"].value2;
            double p03z = xlRange.Cells[6, "D"].value2;

            double p04x = xlRange.Cells[7, "B"].value2;
            double p04y = xlRange.Cells[7, "C"].value2;
            double p04z = xlRange.Cells[7, "D"].value2;

            double cF = 1 / 0.3048;
            XYZ p01 = new XYZ(p01x * cF, p01y * cF, p01z * cF);
            XYZ p02 = new XYZ(p02x * cF, p02y * cF, p02z * cF);
            XYZ p03 = new XYZ(p03x * cF, p03y * cF, p03z * cF);
            XYZ p04 = new XYZ(p04x * cF, p04y * cF, p04z * cF);


            double p11x = xlRange.Cells[9, "B"].value2;
            double p11y = xlRange.Cells[9, "C"].value2;
            double p11z = xlRange.Cells[9, "D"].value2;

            double p12x = xlRange.Cells[10, "B"].value2;
            double p12y = xlRange.Cells[10, "C"].value2;
            double p12z = xlRange.Cells[10, "D"].value2;

            double p13x = xlRange.Cells[11, "B"].value2;
            double p13y = xlRange.Cells[11, "C"].value2;
            double p13z = xlRange.Cells[11, "D"].value2;

            double p14x = xlRange.Cells[12, "B"].value2;
            double p14y = xlRange.Cells[12, "C"].value2;
            double p14z = xlRange.Cells[12, "D"].value2;


            XYZ p11 = new XYZ(p11x * cF, p11y * cF, p11z * cF);
            XYZ p12 = new XYZ(p12x * cF, p12y * cF, p12z * cF);
            XYZ p13 = new XYZ(p13x * cF, p13y * cF, p13z * cF);
            XYZ p14 = new XYZ(p14x * cF, p14y * cF, p14z * cF);

            double pf1x = xlRange.Cells[14, "B"].value2;
            double pf1y = xlRange.Cells[14, "C"].value2;
            double pf1z = xlRange.Cells[14, "D"].value2;

            double pf2x = xlRange.Cells[15, "B"].value2;
            double pf2y = xlRange.Cells[15, "C"].value2;
            double pf2z = xlRange.Cells[15, "D"].value2;

            double pf3x = xlRange.Cells[16, "B"].value2;
            double pf3y = xlRange.Cells[16, "C"].value2;
            double pf3z = xlRange.Cells[16, "D"].value2;

            double pf4x = xlRange.Cells[17, "B"].value2;
            double pf4y = xlRange.Cells[17, "C"].value2;
            double pf4z = xlRange.Cells[17, "D"].value2;

            
            XYZ pf1 = new XYZ(pf1x * cF, pf1y * cF, pf1z * cF);
            XYZ pf2 = new XYZ(pf2x * cF, pf2y * cF, pf2z * cF);
            XYZ pf3 = new XYZ(pf3x * cF, pf3y * cF, pf3z * cF);
            XYZ pf4 = new XYZ(pf4x * cF, pf4y * cF, pf4z * cF);

            Line l01 = Line.CreateBound(p01, p02);
            Line l02 = Line.CreateBound(p02, p03);
            Line l03 = Line.CreateBound(p03, p04);
            Line l04 = Line.CreateBound(p04, p01);

            Line l11 = Line.CreateBound(p11, p12);
            Line l12 = Line.CreateBound(p12, p13);
            Line l13 = Line.CreateBound(p13, p14);
            Line l14 = Line.CreateBound(p14, p11);

            Line lf1 = Line.CreateBound(pf1, pf2);
            Line lf2 = Line.CreateBound(pf2, pf3);
            Line lf3 = Line.CreateBound(pf3, pf4);
            Line lf4 = Line.CreateBound(pf4, pf1);

            List<Curve> curves = new List<Curve>();

            curves.Add(l01);
            curves.Add(l02);
            curves.Add(l03);
            curves.Add(l04);

            List<Curve> curves1 = new List<Curve>();

            curves1.Add(l11);
            curves1.Add(l12);
            curves1.Add(l13);
            curves1.Add(l14);

            List<Curve> curvesf = new List<Curve>();

            curvesf.Add(lf1);
            curvesf.Add(lf2);
            curvesf.Add(lf3);
            curvesf.Add(lf4);

            // Making a curve in a loop
            CurveLoop crvloop = CurveLoop.Create(curves);
            CurveLoop crvloop1 = CurveLoop.Create(curves1);
            CurveLoop crvloopf = CurveLoop.Create(curvesf);

            CurveLoop offcr = CurveLoop.CreateViaOffset(crvloop, 0.1 * cF, new XYZ(0, 0, 1));
            CurveLoop offcr1 = CurveLoop.CreateViaOffset(crvloop1, 0.1 * cF, new XYZ(0, 0, 1));
            CurveLoop offcrf = CurveLoop.CreateViaOffset(crvloopf, 0.1 * cF, new XYZ(0, 0, 1));

            // Creating a curve array object required for method
            CurveArray curArr = new CurveArray();
            CurveArray curArr1 = new CurveArray();
            CurveArray openingCurveArray = new CurveArray();


            foreach (Curve c in offcr)
            {
                //Put the curves to curve array

                curArr.Append(c);
            }
            foreach (Curve c in offcr1)
            {
                //Put the curves to curve array

                curArr1.Append(c);
            }
            foreach (Curve c in offcrf)
            {
                //Put the curves to curve array

                openingCurveArray.Append(c);
            }

            try
            {
                using (Transaction trans = new Transaction(doc, "Bungalow/Flair110/Flair152RE"))
                {
                    trans.Start();


                  doc.Create.NewFloor(curArr, FloorType, levels, false);
                  //  doc.Create.NewFloor(curArr1, FloorType, levels, false);

                    //////Floor flaw =  doc.Create.NewFloor(curArr,false);
                    //Floor fl = doc.Create.NewFloor(curArr1, FloorType, levels, false);

                    //doc.Create.NewOpening(fl, curArr1, false);


                    //XYZ openingEdges1 = pf1;
                    //XYZ openingEdges2 = pf2;
                    //XYZ openingEdges3 = pf3;
                    //XYZ openingEdges4 = pf4;

                    //var openingEdges = {openingEdges1, openingEdges2, openingEdges3, openingEdges4};

                    //var openingCurveArray = openingEdges;
                    var opening = doc.Create.NewFloor(curArr, FloorType, levels, false);

                    doc.Create.NewOpening(opening, openingCurveArray, true);


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


