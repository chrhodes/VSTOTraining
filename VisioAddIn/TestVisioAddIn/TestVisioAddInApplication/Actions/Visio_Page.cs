using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using Microsoft.Office.Interop.Visio;

//using VNC;
//using VNC.Core;

using VNCVisioAddIn = VNC.Visio.VSTOAddIn;

namespace TestVisioAddInApplication.Actions
{
    internal class Visio_Page
    {
        #region Events

        public static void PageChanged(Page page)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            if (page.NameU != page.Name)
            {
                SyncPageNames(Common.VisioApplication, page);
                UpdatePageNameShapes(page);
            }
        }

        #endregion

        #region Action Handlers

        public static void CreatePageForShape(Application app, string doc, string page, string shape, string shapeu, string[] args)
        {
            string prefix = null;
            string delimiter = null;
            string backgroundPageName = null;

            if (args.Count() != 3)
            {
                Common.WriteToDebugWindow($"Incorrect Argument Count {args.Count()}, expected 3.  Check ShapeSheet");
            }
            else
            {
                prefix = args[0];
                delimiter = args[1];
                backgroundPageName = args[2];
            }

            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}() prefix:>{prefix}< delimiter:>{delimiter}< backgroundPageName:>{backgroundPageName}<");

            try
            {
                // Current shape contains text for new page name.
                Page activePage = app.ActivePage;
                Shape activeShape = app.ActivePage.Shapes[shape];
                Common.WriteToDebugWindow($"  Shape(Name:>{activeShape.Name}< Text:>{activeShape.Text}< Characters:>{activeShape.Characters.TextAsString}<");

                string shapePageName = "Error-PageNameNotProvided";

                if (activeShape.CellExistsU["Prop.PageName", 0] != 0)
                {
                    shapePageName = activeShape.CellsU["Prop.PageName"].ResultStrU[VisUnitCodes.visUnitsString];
                }
                else if (activeShape.Characters.TextAsString.Length > 0)
                {
                    //string newPageName = string.Format("{0}{1}{2}", prefix, delimiter, activeShape.Text);
                    // shape.Text comes in as OBJ if use fields and Shape Data.   Use shape.Characters instead.
                    shapePageName = activeShape.Characters.TextAsString;
                }

                string newPageName = $"{prefix}{delimiter}{shapePageName}";

                Page newPage = CreatePage(app, newPageName, backgroundPageName);

                // The old style linkable masters did not have Prop.Data for the HyperLink.  Check first before updating.
                // Should really retire all the old shapes and remove this code.

                if (activeShape.CellExistsU["Prop.HyperLink", 0] != 0)
                {
                    // FIX(crhodes)
                    // Decide if just want to give in and reference VNC.Core
                    //activeShape.CellsU["Prop.HyperLink"].FormulaU = newPageName.WrapInDblQuotes();
                }
                else
                {
                    Hyperlink currentHyperLink = activeShape.AddHyperlink();
                    currentHyperLink.SubAddress = newPageName;
                }

                // Check to see if there is a ReturnLink Property with values that can be used
                // to create a return link to the page that linked to us.

                if (activeShape.CellExistsU["Prop.ReturnLink", 0] != 0)
                {
                    //string returnLinkProp = activeShape.CellsU["Prop.ReturnLink"].FormulaU;   // This returns "<string>"  we want just <string>
                    string returnLinkProp = activeShape.CellsU["Prop.ReturnLink"].ResultStrU[VisUnitCodes.visUnitsString];
                    string[] linkInfo = returnLinkProp.Split(',');
                    string stencilName = linkInfo[0];
                    string shapeName = linkInfo[1];

                    Common.WriteToDebugWindow($"  returnLinkProp:>{returnLinkProp}< stencilName:>{stencilName}< shapeName:>{shapeName}< ");

                    try
                    {
                        Document linkStencil = app.Documents[stencilName];

                        try
                        {
                            Master linkMaster = linkStencil.Masters[shapeName];

                            // Add return link in upper left corner.  Assume 11x8.5

                            // TODO(crhodes)
                            // Get Page Size and drop in upper left
                            Shape returnLinkShape = newPage.Drop(linkMaster, 1.0, 8.0);

                            // FIX(crhodes)
                            // Decide if just want to give in and reference VNC.Core
                            //returnLinkShape.CellsU["Prop.PageName"].FormulaU = activePage.Name.WrapInDblQuotes();
                            //returnLinkShape.CellsU["Prop.HyperLink"].FormulaU = activePage.Name.WrapInDblQuotes();
                        }
                        catch (Exception ex)
                        {
                            Common.WriteToDebugWindow(string.Format("  Cannot find Master named:>{0}<", shapeName));
                        }
                    }
                    catch (Exception ex)
                    {
                        Common.WriteToDebugWindow(string.Format("  Cannot find open Stencil named:>{0}<", stencilName));
                    }
                }

                // Add a header.  May want to pick the stencil and shape from config file.
                // Or add a property to Shape.

                VNCVisioAddIn.Helpers.LoadStencil(app, "Page Shapes.vssx");
                Master headerMaster = app.Documents[@"Page Shapes.vssx"].Masters[@"Sizeable Page Header"];

                // NOTE(crhodes)
                // Doesn't really matter where header is dropped as Sizeable Page Header will position itself.

                newPage.Drop(headerMaster, 5.5, 8.0625);

                // NOTE(crhodes)
                // Add the shape that triggered the event.  User can delete if doesn't want.
                // More and more often I go back and copy it, traverse the link, and drop it.
                // Drop in middle of page for now assuming 11x8.5

                // TODO(crhodes)
                // Get Page Size and drop in center
                newPage.Drop(activeShape, 5.5, 4.25);
            }
            catch (Exception ex)
            {
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        public static void ToggleLayerLock(Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}");

            Shape activeShape = app.ActivePage.Shapes[shape];
            Page activePage = app.ActivePage;

            ToggleLayerSetting(activePage, activeShape, VisCellIndices.visLayerLock, "Prop.Lock");
        }

        public static void ToggleLayerPrint(Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}");

            Shape activeShape = app.ActivePage.Shapes[shape];
            Page activePage = app.ActivePage;

            ToggleLayerSetting(activePage, activeShape, VisCellIndices.visLayerPrint, "Prop.Print");
        }

        public static void ToggleLayerVisibility(Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}");

            Shape activeShape = app.ActivePage.Shapes[shape];
            Page activePage = app.ActivePage;

            ToggleLayerSetting(activePage, activeShape, VisCellIndices.visLayerVisible, "Prop.Visible");
        }

        public static void ToggleLayerSetting(Page activePage, Shape activeShape, VisCellIndices visCell, string cellsU)
        {
            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (Layer layer in activePage.Layers)
            {
                if (layer.Name.ToLower() == layerName.ToLower())
                {
                    var currentState = layer.CellsC[(short)visCell].ResultIU;
                    string newState = null;

                    newState = (currentState == 0) ? "1" : "0";

                    layer.CellsC[(short)visCell].Formula = newState;
                    activeShape.CellsU[cellsU].FormulaU = newState;
                }
            }
        }

        public static void UpdateLayer(Application app, string doc, string page, string shape, string shapeu)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}");

            Shape activeShape = app.ActivePage.Shapes[shape];
            Page activePage = app.ActivePage;

            string layerName = activeShape.CellsU["Prop.Layer"].ResultStrU[0];

            foreach (Layer layer in activePage.Layers)
            {
                try
                {
                    Common.WriteToDebugWindow(layer.Name);

                    if (layer.Name.ToLower() == layerName.ToLower())
                    {
                        if (activeShape.CellExistsU["Prop.Visible", 0] != 0)
                        {
                            layer.CellsC[(short)VisCellIndices.visLayerVisible].FormulaU = activeShape.CellsU["Prop.Visible"].ResultStrU[0];
                        }

                        if (activeShape.CellExistsU["Prop.Lock", 0] != 0)
                        {
                            layer.CellsC[(short)VisCellIndices.visLayerLock].FormulaU = activeShape.CellsU["Prop.Lock"].ResultStrU[0];
                        }

                        if (activeShape.CellExistsU["Prop.Print", 0] != 0)
                        {
                            layer.CellsC[(short)VisCellIndices.visLayerPrint].FormulaU = activeShape.CellsU["Prop.Print"].ResultStrU[0];
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Log.Error(ex, Common.LOG_CATEGORY);
                }
            }
        }

        #endregion

        #region Helpers

        private static Page CreatePage(Application app, string pageName, string backgroundPageName, short isBackground = 0)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}() Page:{pageName} Background:{backgroundPageName}");

            // TODO(crhodes):
            //	Error handling. Page already exists, background page doesn't exist, etc.

            int currentPageIndex = app.ActivePage.Index;

            Common.WriteToDebugWindow($"  currentPageIndex:{currentPageIndex}");

            Page newPage = app.ActiveDocument.Pages.Add();

            // Cleanup page names
            pageName = pageName.Replace("\n", " ");

            newPage.Name = pageName;

            try
            {
                newPage.BackPage = backgroundPageName;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow($"Cannot Find Background Page ({backgroundPageName})");
            }

            newPage.Index = (short)(currentPageIndex + 1);

            newPage.Background = isBackground;

            AddNavigationLinks(app, newPage);

            return newPage;
        }

        private static void AddNavigationLinks(Application app, Page page)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            // Skip Background pages and the "Navigation Links" page.

            if ((page.Background != 0) || (page.NameU == "Navigation Links"))
            {
                //Common.WriteToDebugWindow("   Skipping");
                return;
            }

            RemoveNavigationLinks(page);

            try
            {
                Window activeWindow = app.ActiveWindow;
                activeWindow.Page = app.ActiveDocument.Pages["Navigation Links"];
                activeWindow.SelectAll();
                activeWindow.Selection.Copy(VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                activeWindow.Page = page;
                page.Paste(VisCutCopyPasteCodes.visCopyPasteNoTranslate);

                //Globals.ThisAddIn.Application.Windows.ItemEx["Navigation Links"].Activate();
                //Globals.ThisAddIn.Application.ActiveWindow.SelectAll();
                //Globals.ThisAddIn.Application.ActiveWindow.Selection.Copy();
                //Globals.ThisAddIn.Application.Windows.ItemEx["Navigation Links"].Activate();


                //Visio.Page linkPage = Globals.ThisAddIn.Application.ActiveDocument.Pages["Navigation Links"];
                //linkPage.Application.
                //Globals.ThisAddIn.Application.
                //Common.WriteToDebugWindow(string.Format("  Copying {0} links", linkPage.Shapes.Count));

                //foreach (Visio.Shape shape in linkPage.Shapes)
                //{
                //    // TODO: Make this smarter about only using IsNavigationLink shapes
                //    shape.Copy(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                //    page.Paste(Visio.VisCutCopyPasteCodes.visCopyPasteNoTranslate);
                //}

                //List<Visio.Shape> links = Actions.Visio_Document.GetNavigationLinks();

                // Typically we don't print the stuff on the navigation layer.

                page.Layers["Navigation"].CellsC[(short)VisCellIndices.visLayerPrint].FormulaU = "0";
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString());
                //Log.Error(ex, Common.LOG_CATEGORY);
                // No navigation Links Page perhaps
            }
        }

        private static void RemoveNavigationLinks(Page page)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            List<Shape> navigationLinks = GetNavigationLinks(page);

            try
            {
                foreach (Shape shape in navigationLinks)
                {
                    var isNavigationLink = shape.CellExists["User.IsNavigationLink", 0];  // 0 not limited to local only

                    if (isNavigationLink != 0)
                    {
                        shape.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString());
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        private static List<Shape> GetNavigationLinks(Page page)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            List<Shape> navigationLinks = new List<Shape>();

            foreach (Shape shape in page.Shapes)
            {
                var isNavigationLink = shape.CellExists["User.IsNavigationLink", 0];

                navigationLinks.Add(shape);
            }

            return navigationLinks;
        }

        private static void SyncPageNames(Application app)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}()");

            Document doc = app.ActiveDocument;
            Page page = app.ActivePage;

            SyncPageNames(app, page);
        }

        private static void SyncPageNames(Application app, Page page)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            // NOTE(crhodes)
            // NameU is universal name and Name is localized name
            // NameU can only be changed in code.

            try
            {
                app.EventsEnabled = 0;

                page.NameU = page.Name;
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString());
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
            finally
            {
                app.EventsEnabled = 1;
            }
        }

        private static void UpdatePageNameShapes(Page page)
        {
            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name} Name:>{page.Name}< NameU:>{page.NameU}<");

            foreach (Shape shape in page.Shapes)
            {
                Visio_Shape.UpdatePageNameShape(shape, page.Name);
            }
        }

        #endregion
    }
}
