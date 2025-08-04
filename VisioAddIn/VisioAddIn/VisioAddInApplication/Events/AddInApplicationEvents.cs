using System;
using System.Linq;
using System.Reflection;

using Microsoft.Office.Interop.Visio;

namespace VisioAddInApplication.Events
{
    public class AddInApplicationEvents
    {
        private Application Application;
        public Application VisioApplication
        {
            get
            {
                return Application;
            }
            set
            {
                if (Application != null)
                {
                    // Should remove all the event handlers;
                }

                Application = value;

                // NOTE(crhodes)
                // There are events that are processed by the Application.
                // Remove the event handler from AppEvents
                // Can still call method for logging, see infra.

                if (Application != null)
                {   
                    Application.MarkerEvent += new EApplication_MarkerEventEventHandler(Application_MarkerEvent);
            
                    Application.PageChanged += new EApplication_PageChangedEventHandler(Application_PageChanged);

                    Application.ShapeAdded += new EApplication_ShapeAddedEventHandler(Application_ShapeAdded);

                    Application.WindowTurnedToPage += new EApplication_WindowTurnedToPageEventHandler(Application_WindowTurnedToPage);
                }
            }
        }

        #region Events Handled by Application Code

        // NOTE(crhodes)
        // The VisioAppEvent handlers will log event to watch window.

        void Application_MarkerEvent(Application app, int SequenceNum, string ContextString)
        {
            string message = $"{MethodInfo.GetCurrentMethod().Name} SequenceNum={SequenceNum} ContextString=>{ContextString}<";

            Common.WriteToWatchWindow(message);

            // If we got here from a RUNADDONWARGS("QueueMarkerEvent", "<Action>")
            // the ContextString should have multiple pieces showing the context of what was selected.
            // See RouteShapeSheet_QueueMarkerEvent for details

            try
            {
                if (null != ContextString)
                {
                    var context = ContextString.Split(' ');

                    if (context.Count() > 1)
                    {
                        RouteShapeSheet_QueueMarkerEvent(app, SequenceNum, context); ;
                    }
                    else
                    {
                        // Quietly ignore
                        // No context.
                    }
                }
            }
            catch (Exception ex)
            {
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
        }

        void Application_PageChanged(Page Page)
        {
            Actions.Visio_Page.PageChanged(Page);
        }

        void Application_ShapeAdded(Shape Shape)
        {
            Actions.Visio_Shape.HandleShapeAdded(Shape);
        }

        void Application_WindowTurnedToPage(Window Window)
        {
            Window.ViewFit = (int)VisWindowFit.visFitPage;
        }

        #endregion

        private void RouteShapeSheet_QueueMarkerEvent(Application app, int sequenceNum, string[] context)
        {
            //VNC.Log.Debug("", Common.LOG_CATEGORY, 0);

            Common.WriteToDebugWindow($"{MethodInfo.GetCurrentMethod().Name}()");

            try
            {
                for (int i = 0; i < context.Count(); i++)
                {
                    Common.WriteToDebugWindow($"  ci[{i}]:>{context[i]}");
                }

                // The QueueMarkerEvent provides context information for each event along with user information (action).
                // Each part of the context is preceeded by an identifier of the form /<identifier>=
                // Grab the part of the entry that is past the = sign.

                string doc = context[0].Substring(5);       // "/doc="
                string page = context[1].Substring(6);      // "/page="
                string shape = context[2].Substring(7);     // "/shape="

                Common.WriteToDebugWindow($" doc: >{doc}<  page: >{page}<  shape: >{shape}<");

                // QueueMarkerEvent from Pages does not have a shapeu

                string shapeu = "<none>";

                if (context.Count() > 3)
                {
                    shapeu = context[3].Substring(8);    // "/shapeu="
                    Common.WriteToDebugWindow($"   shapeu:>{shapeu}<");
                }

                string args = context[4].Replace("%20", " ");   // Embedded spaces
                var actionArgs = args.Split(',');

                Common.WriteToDebugWindow($"    actionArgs:>{actionArgs[0]}<");

                // TODO:
                // Add new case statement for each unique "<Action>"
                // RUNADDONWARGS("QueueMarkerEvent", "<Action>,<arg1>,<arg2>")
                // Skip(1) skips past <Action> and passes any <args> (separated by commas) that are present 

                switch (actionArgs[0])
                {
                    #region Visio_Document Actions

                    #endregion

                    #region Visio_Page Actions

                    case "CreatePageForShape":
                        Actions.Visio_Page.CreatePageForShape(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                    case "ToggleLayerLock":
                        Actions.Visio_Page.ToggleLayerLock(app, doc, page, shape, shapeu);
                        break;

                    case "ToggleLayerPrint":
                        Actions.Visio_Page.ToggleLayerPrint(app, doc, page, shape, shapeu);
                        break;

                    case "ToggleLayerVisibility":
                        Actions.Visio_Page.ToggleLayerVisibility(app, doc, page, shape, shapeu);
                        break;

                    case "UpdateLayer":
                        Actions.Visio_Page.UpdateLayer(app, doc, page, shape, shapeu);
                        break;

                    #endregion

                    #region Visio_Shape Actions

                    case "LinkShapeToPage":
                        Actions.Visio_Shape.LinkShapeToPage(app, doc, page, shape, shapeu, actionArgs.Skip(1).ToArray());
                        break;

                        #endregion
                }
            }
            catch (Exception ex)
            {
                //Log.Error(ex, Common.LOG_CATEGORY);
            }
        }
    }
}
