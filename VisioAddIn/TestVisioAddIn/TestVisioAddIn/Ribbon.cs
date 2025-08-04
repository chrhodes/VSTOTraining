using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Tools.Ribbon;

using Visio = Microsoft.Office.Interop.Visio;

namespace TestVisioAddIn
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void AddShapeToNewPage()
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            Visio.Page currentPage = app.ActivePage;

            Visio.Page newPage = doc.Pages.Add();

            Visio.Document stencil = app.Documents.OpenEx("Basic_U.vssx", (short)Visio.VisOpenSaveArgs.visOpenDocked);

            Visio.Shape stencilSquare = currentPage.Drop(stencil.Masters["Square"], 1, 5);
            Visio.Shape stencilCircle = currentPage.Drop(stencil.Masters["Circle"], 3, 5);
            Visio.Shape stencilTriangle = currentPage.Drop(stencil.Masters["Triangle"], 5, 5);

            stencilSquare.Text = "Square";
            stencilCircle.Text = "Circle";
            stencilTriangle.Text = "Triangle";

            newPage.NameU = "My New Page";

            Visio.Shape shape1 = currentPage.DrawRectangle(1, 1, 2, 1.5);
            Visio.Shape shape2 = currentPage.DrawRectangle(1, 3, 2, 3.5);

            Visio.Shape shape3 = newPage.DrawRectangle(1, 1, 2, 1.5);
            Visio.Shape shape4 = newPage.DrawRectangle(1, 3, 2, 3.5);

            shape1.Text = currentPage.Name;
            shape2.Text = newPage.Name;

            shape3.Text = shape3.Name;
            shape4.Text = shape4.Name;
        }

        private void AddFooter()
        {
            Visio.Application app = Globals.ThisAddIn.Application;

            Visio.Document doc = app.ActiveDocument;

            doc.FooterLeft = "&f&e";
            doc.FooterCenter = "";
            doc.FooterRight = "&d &p-&P";

            var font = doc.HeaderFooterFont;

            font.Size = (decimal)8;

            doc.HeaderFooterFont = font;

            var size = doc.HeaderFooterFont.Size;

            doc.HeaderMargin[Visio.VisUnitCodes.visDrawingUnits] = 0.13;
            doc.FooterMargin[Visio.VisUnitCodes.visDrawingUnits] = 0.13;
        }

    }
}
