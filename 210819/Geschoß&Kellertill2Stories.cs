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
        if(keller == "Ja")
            {
               if(geschoss == 1)
                {
                    List<XYZ> pointskg = new List<XYZ>();
                    List<XYZ> pointseg = new List<XYZ>();
                    List<XYZ> pointsug = new List<XYZ>();
                    List<XYZ> pointsop = new List<XYZ>();

                    for(int i = 5;i<=points+4;i++)
                    {
                        XYZ ptskg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, -3.06 * cF);
                        pointskg.Add(ptskg);
                        XYZ ptseg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                        pointseg.Add(ptseg);
                        XYZ ptsug = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                        pointsug.Add(ptsug);
                    }
                    for (int i = 3; i <= 6; i++)
                    {
                        XYZ ptsop = new XYZ(xlRange.Cells[i, "H"].value2 * cF, xlRange.Cells[i, "I"].value2 * cF, xlRange.Cells[i, "J"].value2 * cF);
                        pointsop.Add(ptsop);
                    }

                    List<Line> conlineskg = new List<Line>();
                    List<Line> conlineseg = new List<Line>();
                    List<Line> conlinesug = new List<Line>();
                    List<Line> conlinesop = new List<Line>();
                    for(int i = 0; i < pointskg.Count; i++)
                    {
                        if (i < pointskg.Count - 1)
                        {
                            Line linekg = Line.CreateBound(pointskg.ElementAt(i), pointskg.ElementAt(i + 1));
                            conlineskg.Add(linekg);
                        }
                        else
                        {
                            Line linekg = Line.CreateBound(pointskg.ElementAt(pointskg.Count - 1), pointskg.ElementAt(0));
                            conlineskg.Add(linekg);
                        }
                    }
                    for (int i = 0; i < pointseg.Count; i++)
                    {
                        if (i < pointseg.Count - 1)
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(i), pointseg.ElementAt(i + 1));
                            conlineseg.Add(lineeg);
                        }
                        else
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(pointseg.Count - 1), pointseg.ElementAt(0));
                            conlineseg.Add(lineeg);
                        }
                    }
                    for (int i = 0; i < pointsug.Count; i++)
                    {
                        if (i < pointsug.Count - 1)
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(i), pointsug.ElementAt(i + 1));
                            conlinesug.Add(lineug);
                        }
                        else
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(pointsug.Count - 1), pointsug.ElementAt(0));
                            conlinesug.Add(lineug);
                        }
                    }
                    for (int i = 0; i < pointsop.Count; i++)
                    {
                        if (i < pointsop.Count - 1)
                        {

                            Line lineop = Line.CreateBound(pointsop.ElementAt(i), pointsop.ElementAt(i + 1));
                            conlinesop.Add(lineop);
                        }
                        else
                        {
                            Line lineop = Line.CreateBound(pointsop.ElementAt(3), pointsop.ElementAt(0));
                            conlinesop.Add(lineop);
                        }
                    }

                    List<Curve> licurvekg = new List<Curve>();
                    List<Curve> licurveeg = new List<Curve>();
                    List<Curve> licurveug = new List<Curve>();
                    List<Curve> licurveop = new List<Curve>();
                    for(int i = 0; i < conlineskg.Count; i++)
                    {
                        licurvekg.Add(conlineskg.ElementAt(i));
                    }
                    for(int i=0; i< conlineseg.Count; i++)
                    {
                        licurveeg.Add(conlineseg.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesug.Count; i++)
                    {
                        licurveug.Add(conlinesug.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesop.Count; i++)
                    {
                        licurveop.Add(conlinesop.ElementAt(i));
                    }
                    CurveArray curArrkg = new CurveArray();
                    foreach(Curve ckg in licurvekg)
                    {
                        curArrkg.Append(ckg);
                    }
                    CurveArray curArreg = new CurveArray();
                    foreach(Curve ceg in licurveeg)
                    {
                        curArreg.Append(ceg);
                    }
                    CurveArray curArrug = new CurveArray();
                    foreach(Curve cug in licurveug)
                    {
                        curArrug.Append(cug);
                    }
                    CurveArray curArrop = new CurveArray();
                    foreach(Curve cop in licurveop)
                    {
                        curArrop.Append(cop);
                    }
                    Floor Fkg, Feg, Fug;
                    using (Transaction transfl = new Transaction(doc, "Bungalow Floors"))
                    {
                        transfl.Start();
                        Fkg = doc.Create.NewFloor(curArrkg, FloorType, Level1, false);
                        Feg = doc.Create.NewFloor(curArreg, FloorType, Level1, false);
                        Fug = doc.Create.NewFloor(curArrug, FloorType, Level1, false);
                        transfl.Commit();
                    }
                    using (Transaction transop = new Transaction(doc, "Bungalow Floors"))
                    {
                        transop.Start();
                        doc.Create.NewOpening(Feg, curArrop, false);
                        doc.Create.NewOpening(Fug, curArrop, false);
                        transop.Commit();
                    }
                }
                else if(geschoss == 2)
                {
                    List<XYZ> pointskg = new List<XYZ>();
                    List<XYZ> pointseg = new List<XYZ>();
                    List<XYZ> pointsug = new List<XYZ>();
                    List<XYZ> pointsdg = new List<XYZ>();
                    List<XYZ> pointsop = new List<XYZ>();

                    for (int i = 5; i <= points + 4; i++)
                    {
                        XYZ ptskg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, -3.06 * cF);
                        pointskg.Add(ptskg);
                        XYZ ptseg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                        pointseg.Add(ptseg);
                        XYZ ptsug = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                        pointsug.Add(ptsug);
                        XYZ ptsdg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 *2* cF);
                        pointsdg.Add(ptsdg);
                    }
                    for (int i = 3; i <= 6; i++)
                    {
                        XYZ ptsop = new XYZ(xlRange.Cells[i, "H"].value2 * cF, xlRange.Cells[i, "I"].value2 * cF, xlRange.Cells[i, "J"].value2 * cF);
                        pointsop.Add(ptsop);
                    }

                    List<Line> conlineskg = new List<Line>();
                    List<Line> conlineseg = new List<Line>();
                    List<Line> conlinesug = new List<Line>();
                    List<Line> conlinesdg = new List<Line>();
                    List<Line> conlinesop = new List<Line>();
                    for (int i = 0; i < pointskg.Count; i++)
                    {
                        if (i < pointskg.Count - 1)
                        {
                            Line linekg = Line.CreateBound(pointskg.ElementAt(i), pointskg.ElementAt(i + 1));
                            conlineskg.Add(linekg);
                        }
                        else
                        {
                            Line linekg = Line.CreateBound(pointskg.ElementAt(pointskg.Count - 1), pointskg.ElementAt(0));
                            conlineskg.Add(linekg);
                        }
                    }
                    for (int i = 0; i < pointseg.Count; i++)
                    {
                        if (i < pointseg.Count - 1)
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(i), pointseg.ElementAt(i + 1));
                            conlineseg.Add(lineeg);
                        }
                        else
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(pointseg.Count - 1), pointseg.ElementAt(0));
                            conlineseg.Add(lineeg);
                        }
                    }
                    for (int i = 0; i < pointsug.Count; i++)
                    {
                        if (i < pointsug.Count - 1)
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(i), pointsug.ElementAt(i + 1));
                            conlinesug.Add(lineug);
                        }
                        else
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(pointsug.Count - 1), pointsug.ElementAt(0));
                            conlinesug.Add(lineug);
                        }
                    }
                    for (int i = 0; i < pointsdg.Count; i++)
                    {
                        if (i < pointsdg.Count - 1)
                        {
                            Line linedg = Line.CreateBound(pointsdg.ElementAt(i), pointsdg.ElementAt(i + 1));
                            conlinesdg.Add(linedg);
                        }
                        else
                        {
                            Line linedg = Line.CreateBound(pointsdg.ElementAt(pointsdg.Count - 1), pointsdg.ElementAt(0));
                            conlinesdg.Add(linedg);
                        }
                    }
                    for (int i = 0; i < pointsop.Count; i++)
                    {
                        if (i < pointsop.Count - 1)
                        {

                            Line lineop = Line.CreateBound(pointsop.ElementAt(i), pointsop.ElementAt(i + 1));
                            conlinesop.Add(lineop);
                        }
                        else
                        {
                            Line lineop = Line.CreateBound(pointsop.ElementAt(3), pointsop.ElementAt(0));
                            conlinesop.Add(lineop);
                        }
                    }

                    List<Curve> licurvekg = new List<Curve>();
                    List<Curve> licurveeg = new List<Curve>();
                    List<Curve> licurveug = new List<Curve>();
                    List<Curve> licurvedg = new List<Curve>();
                    List<Curve> licurveop = new List<Curve>();
                    for (int i = 0; i < conlineskg.Count; i++)
                    {
                        licurvekg.Add(conlineskg.ElementAt(i));
                    }
                    for (int i = 0; i < conlineseg.Count; i++)
                    {
                        licurveeg.Add(conlineseg.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesug.Count; i++)
                    {
                        licurveug.Add(conlinesug.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesdg.Count; i++)
                    {
                        licurvedg.Add(conlinesdg.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesop.Count; i++)
                    {
                        licurveop.Add(conlinesop.ElementAt(i));
                    }
                    CurveArray curArrkg = new CurveArray();
                    foreach (Curve ckg in licurvekg)
                    {
                        curArrkg.Append(ckg);
                    }
                    CurveArray curArreg = new CurveArray();
                    foreach (Curve ceg in licurveeg)
                    {
                        curArreg.Append(ceg);
                    }
                    CurveArray curArrug = new CurveArray();
                    foreach (Curve cug in licurveug)
                    {
                        curArrug.Append(cug);
                    }
                    CurveArray curArrdg = new CurveArray();
                    foreach (Curve cdg in licurvedg)
                    {
                        curArrdg.Append(cdg);
                    }
                    CurveArray curArrop = new CurveArray();
                    foreach (Curve cop in licurveop)
                    {
                        curArrop.Append(cop);
                    }
                    Floor Fkg, Feg, Fug, Fdg;
                    using (Transaction transfl = new Transaction(doc, "Bungalow Floors"))
                    {
                        transfl.Start();
                        Fkg = doc.Create.NewFloor(curArrkg, FloorType, Level1, false);
                        Feg = doc.Create.NewFloor(curArreg, FloorType, Level1, false);
                        Fug = doc.Create.NewFloor(curArrug, FloorType, Level1, false);
                        Fdg = doc.Create.NewFloor(curArrdg, FloorType, Level1, false);
                        transfl.Commit();
                    }
                    using (Transaction transop = new Transaction(doc, "Bungalow Floors"))
                    {
                        transop.Start();
                        doc.Create.NewOpening(Feg, curArrop, false);
                        doc.Create.NewOpening(Fug, curArrop, false);
                        doc.Create.NewOpening(Fdg, curArrop, false);
                        transop.Commit();
                    }
                }

            }
            else if (keller == "Nein")
            {
                if (geschoss == 1)
                {
                    
                    List<XYZ> pointseg = new List<XYZ>();
                    List<XYZ> pointsug = new List<XYZ>();
                    List<XYZ> pointsop = new List<XYZ>();

                    for (int i = 5; i <= points + 4; i++)
                    {
                        
                        XYZ ptseg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                        pointseg.Add(ptseg);
                        XYZ ptsug = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                        pointsug.Add(ptsug);
                    }
                    for (int i = 3; i <= 6; i++)
                    {
                        XYZ ptsop = new XYZ(xlRange.Cells[i, "H"].value2 * cF, xlRange.Cells[i, "I"].value2 * cF, xlRange.Cells[i, "J"].value2 * cF);
                        pointsop.Add(ptsop);
                    }

                    
                    List<Line> conlineseg = new List<Line>();
                    List<Line> conlinesug = new List<Line>();
                    List<Line> conlinesop = new List<Line>();
                    
                    for (int i = 0; i < pointseg.Count; i++)
                    {
                        if (i < pointseg.Count - 1)
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(i), pointseg.ElementAt(i + 1));
                            conlineseg.Add(lineeg);
                        }
                        else
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(pointseg.Count - 1), pointseg.ElementAt(0));
                            conlineseg.Add(lineeg);
                        }
                    }
                    for (int i = 0; i < pointsug.Count; i++)
                    {
                        if (i < pointsug.Count - 1)
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(i), pointsug.ElementAt(i + 1));
                            conlinesug.Add(lineug);
                        }
                        else
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(pointsug.Count - 1), pointsug.ElementAt(0));
                            conlinesug.Add(lineug);
                        }
                    }
                    for (int i = 0; i < pointsop.Count; i++)
                    {
                        if (i < pointsop.Count - 1)
                        {

                            Line lineop = Line.CreateBound(pointsop.ElementAt(i), pointsop.ElementAt(i + 1));
                            conlinesop.Add(lineop);
                        }
                        else
                        {
                            Line lineop = Line.CreateBound(pointsop.ElementAt(3), pointsop.ElementAt(0));
                            conlinesop.Add(lineop);
                        }
                    }

                    
                    List<Curve> licurveeg = new List<Curve>();
                    List<Curve> licurveug = new List<Curve>();
                    List<Curve> licurveop = new List<Curve>();
                    
                    for (int i = 0; i < conlineseg.Count; i++)
                    {
                        licurveeg.Add(conlineseg.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesug.Count; i++)
                    {
                        licurveug.Add(conlinesug.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesop.Count; i++)
                    {
                        licurveop.Add(conlinesop.ElementAt(i));
                    }
                    
                    CurveArray curArreg = new CurveArray();
                    foreach (Curve ceg in licurveeg)
                    {
                        curArreg.Append(ceg);
                    }
                    CurveArray curArrug = new CurveArray();
                    foreach (Curve cug in licurveug)
                    {
                        curArrug.Append(cug);
                    }
                    CurveArray curArrop = new CurveArray();
                    foreach (Curve cop in licurveop)
                    {
                        curArrop.Append(cop);
                    }
                    Floor Feg, Fug;
                    using (Transaction transfl = new Transaction(doc, "Bungalow Floors"))
                    {
                        transfl.Start();
                        
                        Feg = doc.Create.NewFloor(curArreg, FloorType, Level1, false);
                        Fug = doc.Create.NewFloor(curArrug, FloorType, Level1, false);
                        transfl.Commit();
                    }
                    using (Transaction transop = new Transaction(doc, "Bungalow Floors"))
                    {
                        transop.Start();
                        
                        doc.Create.NewOpening(Fug, curArrop, false);
                        transop.Commit();
                    }
                }
                else if (geschoss == 2)
                {
                    
                    List<XYZ> pointseg = new List<XYZ>();
                    List<XYZ> pointsug = new List<XYZ>();
                    List<XYZ> pointsdg = new List<XYZ>();
                    List<XYZ> pointsop = new List<XYZ>();

                    for (int i = 5; i <= points + 4; i++)
                    {
                       
                        XYZ ptseg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, 0 * cF);
                        pointseg.Add(ptseg);
                        XYZ ptsug = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * cF);
                        pointsug.Add(ptsug);
                        XYZ ptsdg = new XYZ(xlRange.Cells[i, "B"].value2 * cF, xlRange.Cells[i, "C"].value2 * cF, xlRange.Cells[i, "D"].value2 * 2 * cF);
                        pointsdg.Add(ptsdg);
                    }
                    for (int i = 3; i <= 6; i++)
                    {
                        XYZ ptsop = new XYZ(xlRange.Cells[i, "H"].value2 * cF, xlRange.Cells[i, "I"].value2 * cF, xlRange.Cells[i, "J"].value2 * cF);
                        pointsop.Add(ptsop);
                    }

                    
                    List<Line> conlineseg = new List<Line>();
                    List<Line> conlinesug = new List<Line>();
                    List<Line> conlinesdg = new List<Line>();
                    List<Line> conlinesop = new List<Line>();
                    
                    for (int i = 0; i < pointseg.Count; i++)
                    {
                        if (i < pointseg.Count - 1)
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(i), pointseg.ElementAt(i + 1));
                            conlineseg.Add(lineeg);
                        }
                        else
                        {
                            Line lineeg = Line.CreateBound(pointseg.ElementAt(pointseg.Count - 1), pointseg.ElementAt(0));
                            conlineseg.Add(lineeg);
                        }
                    }
                    for (int i = 0; i < pointsug.Count; i++)
                    {
                        if (i < pointsug.Count - 1)
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(i), pointsug.ElementAt(i + 1));
                            conlinesug.Add(lineug);
                        }
                        else
                        {
                            Line lineug = Line.CreateBound(pointsug.ElementAt(pointsug.Count - 1), pointsug.ElementAt(0));
                            conlinesug.Add(lineug);
                        }
                    }
                    for (int i = 0; i < pointsdg.Count; i++)
                    {
                        if (i < pointsdg.Count - 1)
                        {
                            Line linedg = Line.CreateBound(pointsdg.ElementAt(i), pointsdg.ElementAt(i + 1));
                            conlinesdg.Add(linedg);
                        }
                        else
                        {
                            Line linedg = Line.CreateBound(pointsdg.ElementAt(pointsdg.Count - 1), pointsdg.ElementAt(0));
                            conlinesdg.Add(linedg);
                        }
                    }
                    for (int i = 0; i < pointsop.Count; i++)
                    {
                        if (i < pointsop.Count - 1)
                        {

                            Line lineop = Line.CreateBound(pointsop.ElementAt(i), pointsop.ElementAt(i + 1));
                            conlinesop.Add(lineop);
                        }
                        else
                        {
                            Line lineop = Line.CreateBound(pointsop.ElementAt(3), pointsop.ElementAt(0));
                            conlinesop.Add(lineop);
                        }
                    }

                    
                    List<Curve> licurveeg = new List<Curve>();
                    List<Curve> licurveug = new List<Curve>();
                    List<Curve> licurvedg = new List<Curve>();
                    List<Curve> licurveop = new List<Curve>();
                    
                    for (int i = 0; i < conlineseg.Count; i++)
                    {
                        licurveeg.Add(conlineseg.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesug.Count; i++)
                    {
                        licurveug.Add(conlinesug.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesdg.Count; i++)
                    {
                        licurvedg.Add(conlinesdg.ElementAt(i));
                    }
                    for (int i = 0; i < conlinesop.Count; i++)
                    {
                        licurveop.Add(conlinesop.ElementAt(i));
                    }
                    
                    CurveArray curArreg = new CurveArray();
                    foreach (Curve ceg in licurveeg)
                    {
                        curArreg.Append(ceg);
                    }
                    CurveArray curArrug = new CurveArray();
                    foreach (Curve cug in licurveug)
                    {
                        curArrug.Append(cug);
                    }
                    CurveArray curArrdg = new CurveArray();
                    foreach (Curve cdg in licurvedg)
                    {
                        curArrdg.Append(cdg);
                    }
                    CurveArray curArrop = new CurveArray();
                    foreach (Curve cop in licurveop)
                    {
                        curArrop.Append(cop);
                    }
                    Floor Feg, Fug, Fdg;
                    using (Transaction transfl = new Transaction(doc, "Bungalow Floors"))
                    {
                        transfl.Start();
                        
                        Feg = doc.Create.NewFloor(curArreg, FloorType, Level1, false);
                        Fug = doc.Create.NewFloor(curArrug, FloorType, Level1, false);
                        Fdg = doc.Create.NewFloor(curArrdg, FloorType, Level1, false);
                        transfl.Commit();
                    }
                    using (Transaction transop = new Transaction(doc, "Bungalow Floors"))
                    {
                        transop.Start();
                        
                        doc.Create.NewOpening(Fug, curArrop, false);
                        doc.Create.NewOpening(Fdg, curArrop, false);
                        transop.Commit();
                    }
                }

            }


            return Result.Succeeded;


        }

    }
}