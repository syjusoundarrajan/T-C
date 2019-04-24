using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using Line = Autodesk.Revit.DB.Line;



namespace MyRevitPluginTasks
{
    class Floor : IExternalCommand
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
            double cF = 1 / 0.3028;
            XYZ p1 = new XYZ(0 * cF, 0 * cF, 0 * cF);
            XYZ p2 = new XYZ(13.135 * cF, 0 * cF, 0 * cF);
            XYZ p3 = new XYZ(13.135 * cF, 9.135 * cF, 0 * cF);
            XYZ p4 = new XYZ(0 * cF, 9.135 * cF, 0 * cF);
            Line l1 = Line.CreateBound(p1, p2);
            Line l2 = Line.CreateBound(p2, p3);
            Line l3 = Line.CreateBound(p3, p4);
            Line l4 = Line.CreateBound(p4, p1);

            List<Curve> curves = new List<Curve>();
            curves.Add(l1);
            curves.Add(l2);
            curves.Add(l3);
            curves.Add(l4);

            // // making a curve in a loop
            CurveLoop crvloop = CurveLoop.Create(curves);
            ////double off = UnitUtils.ConvertFromInternalUnits(120, DisplayUnitType.DUT_MILLIMETERS);
            // // giving the offset
            CurveLoop offcr = CurveLoop.CreateViaOffset(crvloop, 0.1 * cF, new XYZ(0, 0, 1));

            //// creating a curve array object required for method
            CurveArray curArr = new CurveArray();

            foreach (Curve c in offcr)
            {
                //// To put the curves to Currve array
                // //Append adds data to a StringBuilder

                curArr.Append(c);

            }

            try
            {
                using (Transaction trans = new Transaction(doc, "Bungalow"))
                {
                    trans.Start();

                    // for foundation

                    doc.Create.NewFloor(curArr, false);


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
    



    

