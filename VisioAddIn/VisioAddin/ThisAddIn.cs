namespace VisioAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //MessageBox.Show("Visio - ThisAddin_Startup");

            InitializeRibbonUI();

            if (Common.EnableAppEvents)
            {
                if (Common.AppEvents == null)
                {
                    Common.AppEvents = new VisioAddInApplication.Events.VisioAppEvents();
                    Common.AppEvents.VisioApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
            }

            // NOTE(crhodes)
            // These are the events that the AddIn depends on

            Common.AddInApplicationEvents = new VisioAddInApplication.Events.AddInApplicationEvents();
            Common.AddInApplicationEvents.VisioApplication = Globals.ThisAddIn.Application;
            
            Common.VisioApplication = Globals.ThisAddIn.Application;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
