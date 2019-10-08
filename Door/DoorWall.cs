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

namespace MyRevitCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    class DoorsWall : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            // get uI document
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //Get dokument
            Document doc = uidoc.Document;
            // Create levels
            Level level = new FilteredElementCollector(doc)
                           .OfCategory(BuiltInCategory.OST_Levels)
                           .WhereElementIsNotElementType()
                           .Cast<Level>()
                           .First(x => x.Name == "Level 1");

            var wallExterior = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsElementType().Cast<WallType>().First(x => x.Name == "AW 01 a & b(24cm)");
            //var doorType = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsElementType().Cast<FamilyInstance>().First(x => x.Name == "Innentür");

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_Doors);

            FamilySymbol symbol = collector.First(x => x.Name == "Innentür") as FamilySymbol;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\WerkStudent\Mail\Gen_Open_Fin_keller");
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // Input for client
            double cF = 1 / 0.3048;
            double height = 2.755;
            double z = 0.0;
            

            //Exterior Walls
            XYZ p1 = new XYZ(xlRange.Cells[5, "B"].value2 * cF, xlRange.Cells[5, "C"].value2 * cF, z);
            XYZ p2 = new XYZ(xlRange.Cells[6, "B"].value2 * cF, xlRange.Cells[6, "C"].value2 * cF, z);
            XYZ p3 = new XYZ(xlRange.Cells[7, "B"].value2 * cF, xlRange.Cells[7, "C"].value2 * cF, z);
            XYZ p4 = new XYZ(xlRange.Cells[8, "B"].value2 * cF, xlRange.Cells[8, "C"].value2 * cF, z);
            XYZ pd = new XYZ(xlRange.Cells[9, "H"].value2 * cF, xlRange.Cells[9, "I"].value2 * cF, z);
            
            //Exterior curves
            Line l1 = Line.CreateBound(p1, p2);
            Line l2 = Line.CreateBound(p2, p3);
            Line l3 = Line.CreateBound(p3, p4);
            Line l4 = Line.CreateBound(p4, p1);

            List<Curve> curExterior = new List<Curve>();
            curExterior.Add(l1);
            curExterior.Add(l2);
            curExterior.Add(l3);
            curExterior.Add(l4);


            Wall wb;
            try
            {
                using (Transaction trans = new Transaction(doc, "Bungalow"))
                {
                    trans.Start();

                    List<Wall> wl = new List<Wall>();
                    foreach (Curve cExt in curExterior)
                    {

                        wb = Wall.Create(doc, cExt, wallExterior.Id, level.Id, height * cF, 0, false, false);
                        wl.Add(wb);
                    }

                    int i = Convert.ToInt32(xlRange.Cells[8, "G"].value2);
                    doc.Create.NewFamilyInstance(pd, symbol, wl.ElementAt(i-1), level, StructuralType.NonStructural);
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
