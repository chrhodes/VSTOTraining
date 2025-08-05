using System;

using VisioPrismAddInApplication.Visio;

namespace VisioPrismAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Int64 startTicks = Common.WriteToDebugWindow("ThisAddIn_Startup()", true);
                //MessageBox.Show("Visio - ThisAddin_Startup");

                InitializeRibbonUI();

                if (Common.EnableAppEvents)
                {
                    if (Common.AppEvents == null)
                    {
                        Common.AppEvents = new VisioPrismAddInApplication.Events.VisioAppEvents();
                        Common.AppEvents.VisioApplication = Globals.ThisAddIn.Application;
                    }
                }
                else
                {
                    Common.AppEvents = null;
                }

                // NOTE(crhodes)
                // These are the events that the AddIn depends on

                Common.AddInApplicationEvents = new VisioPrismAddInApplication.Events.AddInApplicationEvents();
                Common.AddInApplicationEvents.VisioApplication = Globals.ThisAddIn.Application;

                Common.VisioApplication = Globals.ThisAddIn.Application;

                // NOTE(crhodes)
                // Initiaze the AddInApplication.
                // This creates the WPF/Prism Environment in a VisioPrismAddInApplication.

                AddInApplication.InitializeApplication();

                Common.WriteToDebugWindow("ThisAddIn_Startup()", startTicks, true);
            }
            catch (Exception ex)
            {
                //Log.Error(ex, Common.LOG_CATEGORY);
                throw (ex);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Common.WriteToDebugWindow("ThisAddIn_Shutdown()", true);
            //MessageBox.Show("Visio - ThisAddin_Shutdown");
        }

        void InitializeRibbonUI()
        {
            Globals.Ribbons.Ribbon.rgDebug.Visible = Common.DeveloperMode = false;

            Globals.Ribbons.Ribbon.rcbEnableAppEvents.Checked = Common.EnableAppEvents = false;

            // NOTE(crhodes)
            // No need to display during normal operation.
            // More for understanding what Visio is doing during development.
            Globals.Ribbons.Ribbon.rcbDisplayEvents.Checked = Common.DisplayEvents = false;
            Globals.Ribbons.Ribbon.rcbDisplayChattyEvents.Checked = Common.DisplayChattyEvents = false;
        } 

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
