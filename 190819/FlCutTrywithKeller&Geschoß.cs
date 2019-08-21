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
    class FloorCutOut : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\WerkStudent\Mail\Gen_Open_Fin_kel");

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
            string keller = (String)xlRange.Cells[2, "F"].value2;
            int geschoss = Convert.ToInt32(xlRange.Cells[3, "F"].value2);

            List<XYZ> pointlistkf = new List<XYZ>();
            List<XYZ> pointlistgf = new List<XYZ>();
            List<XYZ> pointlistufs = new List<XYZ>();

            for (int i = 5; i <= points + 4; i++)
            {
                if (keller == "Ja")
                {
                    XYZ ptskf = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, -3.06 * cF);
                    pointlistkf.Add(ptskf);
                    XYZ ptsgf = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                    pointlistgf.Add(ptsgf);
                    for (int j = 0; j < geschoss; j++)
                    {
                        XYZ ptsufs = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, (xlRange.Cells[i, "D"].value2+(j * 3.06)) * cF);
                        pointlistufs.Add(ptsufs);
                    }

                }
                else
                {
                    XYZ ptsgf = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                    pointlistgf.Add(ptsgf);
                    for (int j = 0; j < geschoss; j++)
                    {
                        XYZ ptsufs = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, (xlRange.Cells[i, "D"].value2+(j *3.06))* cF);
                        pointlistufs.Add(ptsufs);
                    }
                }
            }

            List<Line> concurveskf = new List<Line>();
            List<Line> concurvesgf = new List<Line>();
            

            for (int i = 0; i < pointlistkf.Count; i++)
            {

                if (i < pointlistkf.Count - 1)
                {
                    Line Alllines = Line.CreateBound(pointlistkf.ElementAt(i), pointlistkf.ElementAt(i + 1));
                    concurveskf.Add(Alllines);
                }
                else
                {
                    Line Alllines = Line.CreateBound(pointlistkf.ElementAt(pointlistkf.Count - 1), pointlistkf.ElementAt(0));
                    concurveskf.Add(Alllines);
                }
            }

            for (int i = 0; i < pointlistgf.Count; i++)
            {

                if (i < pointlistgf.Count - 1)
                {
                    Line Alllines = Line.CreateBound(pointlistgf.ElementAt(i), pointlistgf.ElementAt(i + 1));
                    concurvesgf.Add(Alllines);
                }
                else
                {
                    Line Alllines = Line.CreateBound(pointlistgf.ElementAt(pointlistgf.Count - 1), pointlistgf.ElementAt(0));
                    concurvesgf.Add(Alllines);
                }
            }

            List<Line> concurvesuf = new List<Line>();
            List<Line> concurvesuf1 = new List<Line>();
            List<Line> concurvesuf2 = new List<Line>();

            switch (geschoss)
            {
            case 1:
                 break;
            case 2:
            for (int i = 0; i < 4; i++)
            {
                if (i < 3)
                {
                    Line Alllines = Line.CreateBound(pointlistufs.ElementAt(i), pointlistufs.ElementAt(i + 1));
                    concurvesuf.Add(Alllines); 
                }
                else
                {
                    Line Alllines = Line.CreateBound(pointlistufs.ElementAt(pointlistufs.Count - 1), pointlistufs.ElementAt(0));
                    concurvesuf.Add(Alllines);
                }
            }
                break;
            case 3:
            for (int i = 0; i < 4; i++)
            {
                if (i < 3)
                {
                     Line Alllines = Line.CreateBound(pointlistufs.ElementAt(i), pointlistufs.ElementAt(i + 1));
                     Line Alllines1 = Line.CreateBound(pointlistufs.ElementAt(i + 4), pointlistufs.ElementAt(i + 5));
                     concurvesuf.Add(Alllines);
                     concurvesuf1.Add(Alllines1);
                }
                else
                {
                     Line Alllines = Line.CreateBound(pointlistufs.ElementAt(i), pointlistufs.ElementAt(0));
                     Line Alllines1 = Line.CreateBound(pointlistufs.ElementAt(i + 4), pointlistufs.ElementAt(i+1));
                     concurvesuf.Add(Alllines);
                     concurvesuf1.Add(Alllines1);
                }
            }
                break;
            case 4:
            for (int i = 0; i < 4; i++)
            {
                if (i < 3)
                {
                     Line Alllines = Line.CreateBound(pointlistufs.ElementAt(i), pointlistufs.ElementAt(i + 1));
                     Line Alllines1 = Line.CreateBound(pointlistufs.ElementAt(i + 4), pointlistufs.ElementAt(i + 5));
                     Line Alllines2 = Line.CreateBound(pointlistufs.ElementAt(i + 8), pointlistufs.ElementAt(i + 9));
                     concurvesuf.Add(Alllines);
                     concurvesuf1.Add(Alllines1);
                     concurvesuf2.Add(Alllines2);
                }
                else
                {
                     Line Alllines = Line.CreateBound(pointlistufs.ElementAt(i), pointlistufs.ElementAt(0));
                     Line Alllines1 = Line.CreateBound(pointlistufs.ElementAt(i + 4), pointlistufs.ElementAt(i + 1));
                     Line Alllines2 = Line.CreateBound(pointlistufs.ElementAt(i + 8), pointlistufs.ElementAt(i + 5));
                     concurvesuf.Add(Alllines);
                     concurvesuf1.Add(Alllines1);
                     concurvesuf2.Add(Alllines2);
                }
            }
                break;
            }

            List<Curve> licurvekf = new List<Curve>();
            List<Curve> licurvegf = new List<Curve>();
            
            for (int i = 0; i < concurveskf.Count; i++)
            {
                licurvekf.Add(concurveskf.ElementAt(i));
            }
            for (int i = 0; i < concurvesgf.Count; i++)
            {
                licurvegf.Add(concurvesgf.ElementAt(i));
            }

            List<Curve> licurveuf = new List<Curve>();
            for (int i = 0; i < concurvesuf.Count; i++)
            {
                licurveuf.Add(concurvesuf.ElementAt(i));
            }
            List<Curve> licurveuf1 = new List<Curve>();
            for (int i = 0; i < concurvesuf1.Count; i++)
            {
                licurveuf1.Add(concurvesuf1.ElementAt(i));
            }
            List<Curve> licurveuf2 = new List<Curve>();
            for (int i = 0; i < concurvesuf2.Count; i++)
            {
                licurveuf2.Add(concurvesuf2.ElementAt(i));
            }
            // Making a curve in a loop
            List<CurveLoop> loops = new List<CurveLoop>();
            CurveLoop crvloop0 = CurveLoop.Create(licurvekf);
            loops.Add(crvloop0);
            CurveLoop crvloop = CurveLoop.Create(licurvegf);
            loops.Add(crvloop);
            CurveLoop crvloop1 = CurveLoop.Create(licurveuf);
            loops.Add(crvloop1);
            CurveLoop crvloop2 = CurveLoop.Create(licurveuf1);
            loops.Add(crvloop2);
            CurveLoop crvloop3 = CurveLoop.Create(licurveuf2);
            loops.Add(crvloop3);

            List<CurveArray> curArri = new List<CurveArray>();
           
            for (int i = 0; i < loops.Count; i++)
            {
                CurveArray curArr = new CurveArray();
                foreach (Curve c in loops[i])
                {
                    
                    //Put the curves to curve array
                    curArr.Append(c);
                }
                curArri.Add(curArr);
            }
            

            //
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
            CurveLoop crvvloopopen = CurveLoop.Create(ocurve);
            CurveArray curvopen = new CurveArray();
            foreach (Curve co in crvvloopopen)
            {
                //Put the curves to curve array
                curvopen.Append(co);
            }
            //
            try
            {
                Floor FL;
                List<Floor> stories = new List<Floor>();
                
                if (keller == "Ja")
                {

                    for (int i = 0; i <= geschoss; i++)
                    {

                        using (Transaction trans = new Transaction(doc, "Bungalow Floors"))
                        {
                            trans.Start();
                            FL = doc.Create.NewFloor(curArri[i], FloorType, Level1, false);
                            stories.Add(FL);
                            trans.Commit();
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < geschoss; i++)
                    {

                        using (Transaction trans = new Transaction(doc, "Bungalow Floors"))
                        {
                            trans.Start();
                            FL = doc.Create.NewFloor(curArri[i], FloorType, Level1, false);
                            stories.Add(FL);
                            trans.Commit();
                        }
                    }
                }
                for (int i = 1; i < stories.Count; i++)
                {
                    using (Transaction trans1 = new Transaction(doc, "Bung open"))
                    {
                        trans1.Start();
                        //var Open = doc.Create.NewFloor(curArr1, FloorType, Level2, false);
                        doc.Create.NewOpening(stories[i], curvopen, true);
                        trans1.Commit();
                    }
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
