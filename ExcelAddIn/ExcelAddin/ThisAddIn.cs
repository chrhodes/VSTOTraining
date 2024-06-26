﻿namespace ExcelAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            InitializeRibbonUI();

            if (Common.EnableAppEvents)
            {
                if (Common.AppEvents == null)
                {
                    Common.AppEvents = new ExcelAddInApplication.Events.ExcelAppEvents();
                    Common.AppEvents.ExcelApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents = null;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //MessageBox.Show("Excel - ThisAddin_Shutdown");
        }

        void InitializeRibbonUI()
        {
            Globals.Ribbons.Ribbon.rgDebug.Visible = Common.DeveloperMode = false;

            // NOTE(crhodes)
            // Needed for several events handled by this Addin
            Globals.Ribbons.Ribbon.rcbEnableAppEvents.Checked = Common.EnableAppEvents = true;

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
