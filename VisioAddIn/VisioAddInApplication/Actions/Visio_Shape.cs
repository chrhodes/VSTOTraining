using System.Reflection;

using Microsoft.Office.Interop.Visio;
using System;

namespace VisioAddInApplication.Actions
{
    internal class Visio_Shape
    {
        public static void HandleShapeAdded(Shape shape)
        {
            var isPageNameShape = shape.CellExists["User.IsPageName", 0];    // 0 is Local and Inherited, 1 is Local only 

            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}() isPageName:{isPageNameShape}");

            if (0 != isPageNameShape)
            {
                Application app = Common.VisioApplication;
                Page page = app.ActivePage;
                shape.Text = page.NameU;
            }
        }

        public static void LinkShapeToPage(Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            string pageLevel = array[0];
            string separator = "";

            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}() PageLevel:{pageLevel}");

            // Current shape contains text for new page name.

            Shape activeShape = app.ActivePage.Shapes[shape];
            Common.WriteToDebugWindow($"  Shape(Name:{activeShape.Name}  Text:{activeShape.Text}");

            // Update the current shape's hyperlink to point to the page represented by the text

            if (pageLevel.Length > 0)
            {
                separator = "-";
            }

            // shape.Text comes in as OBJ if use fields and Shape Data.   Use shape.Characters instead. 

            string pageName = $"{pageLevel}{separator}{activeShape.Characters.TextAsString.Replace("\n", " ")}";
            //string pageName = string.Format("{0}{1}{2}", pageLevel, separator, activeShape.Text.Replace("\n", " "));

            Hyperlink newHyperLink = activeShape.AddHyperlink();
            newHyperLink.SubAddress = pageName;
        }

        public static void UpdatePageNameShape(Shape shape, string pageName)
        {
            var isPageName = shape.CellExistsU["User.IsPageName", 0];    // 0 is Local and Inherited, 1 is Local only 

            Common.WriteToDebugWindow(string.Format("{0}({1}  isPageName:{2})",
                MethodBase.GetCurrentMethod().Name, shape.Name, isPageName));

            if (isPageName != 0)
            {
                Cell cell = shape.CellsU["User.IsPageName"];

                Common.WriteToDebugWindow(string.Format("    Shape({0}).Cell(Section:{1} RowName:{2} Name:{3} Value:{4})",
                    shape.Name, cell.Section, cell.RowName, cell.Name, cell.ResultIU));

                if (cell.ResultIU > 0)
                {
                    shape.Text = pageName;
                }
            }
        }
    }
}
