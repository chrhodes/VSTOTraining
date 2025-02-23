using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Tools.Ribbon;

namespace VisioAddInTemplate
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        #region Debug Events

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayDebugWindow();
        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayWatchWindow();
        }

        private void rcbEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.EnableAppEvents = rcbEnableAppEvents.Checked;

            if (Common.EnableAppEvents)
            {
                if (Common.AppEvents == null)
                {
                    Common.AppEvents = new VisioAddInApplicationTemplate.Events.VisioAppEvents();
                    Common.AppEvents.VisioApplication = Globals.ThisAddIn.Application;
                }
            }
            else
            {
                Common.AppEvents.VisioApplication = null;
                Common.AppEvents = null;
            }
        }

        private void rcbDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayEvents = rcbDisplayEvents.Checked;
        }

        private void rcbDisplayChattyEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayChattyEvents = rcbDisplayChattyEvents.Checked;
        }

        private void rcbDeveloperUIMode_Click(object sender, RibbonControlEventArgs e)
        {

        }

        #endregion  

        #region Help Events

        private void btnAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayAddInInfo();
        }

        private void btnDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DeveloperMode = !Common.DeveloperMode;
            Globals.Ribbons.Ribbon.rgDebug.Visible = Common.DeveloperMode;
        }

        #endregion

        #region Private Methods

        private void DisplayAddInInfo()
        {
            // TODO(crhodes)
            // Think through how to fix this and not have to reference VNC.AssemblyHelper or just do it
            VisioAddInApplicationTemplate.AddInInfo.DisplayInfo();
        }

        private void DisplayWatchWindow()
        {
            Common.WatchWindow.Visible = !Common.WatchWindow.Visible;
        }

        private void DisplayDebugWindow()
        {
            Common.DebugWindow.Visible = !Common.DebugWindow.Visible;
        }

        #endregion
    }
}
