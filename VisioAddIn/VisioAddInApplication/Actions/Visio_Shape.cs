using System.Reflection;

using Microsoft.Office.Interop.Visio;

namespace VisioAddInApplication.Actions
{
    internal class Visio_Shape
    {
        public static void LinkShapeToPage(Application app, string doc, string page, string shape, string shapeu, string[] array)
        {
            Common.WriteToDebugWindow("LinkShapeToPage");
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
