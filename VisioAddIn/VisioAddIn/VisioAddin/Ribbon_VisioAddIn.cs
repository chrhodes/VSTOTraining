using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

using Microsoft.Office.Tools.Ribbon;
using VisioAddInApplication;

namespace VisioAddIn
{
    public partial class Ribbon
    {
        #region EventHandlers

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("button1_Click");
            AddShapeToNewPage();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBox.Show("button1_Click");
            AddFooter();
        }

        #region Help Events

        private void btnDisplayAddInInfo_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayAddInInfo();
        }

        private void btnToggleDeveloperMode_Click(object sender, RibbonControlEventArgs e)
        {
            VNC.VSTOAddIn.Common.DeveloperMode = !VNC.VSTOAddIn.Common.DeveloperMode;
            Globals.Ribbons.Ribbon.rgDebug.Visible = VNC.VSTOAddIn.Common.DeveloperMode;
        }

        #endregion

        #region Debug Events

        private void btnDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayDebugWindow();
        }

        private void btnWatchWindow_Click(object sender, RibbonControlEventArgs e)
        {
            DisplayWatchWindow();
        }

        private void rcbLogToDebugWindow_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(System.Reflection.MethodInfo.GetCurrentMethod().Name);
        }

        private void rcbEnableAppEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.EnableAppEvents = rcbEnableAppEvents.Checked;

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
                Common.AppEvents.VisioApplication = null;
                Common.AppEvents = null;
            }
        }

        private void rcbDisplayEvents_Click(object sender, RibbonControlEventArgs e)
        {
            VisioAddInApplication.Common.DisplayEvents = rcbDisplayEvents.Checked;
        }

        private void rcbDisplayChattyEvents_Click(object sender, RibbonControlEventArgs e)
        {
            Common.DisplayChattyEvents = rcbDisplayChattyEvents.Checked;
        }

        #endregion

        #endregion

        #region Private Methods
        private void DisplayAddInInfo()
        {
            VNC.VSTOAddIn.AddInInfo.DisplayInfo();
        }

        private void DisplayWatchWindow()
        {
            VNC.VSTOAddIn.Common.WatchWindow.Visible = !VNC.VSTOAddIn.Common.WatchWindow.Visible;
        }

        private void DisplayDebugWindow()
        {
            VNC.VSTOAddIn.Common.DebugWindow.Visible = !VNC.VSTOAddIn.Common.DebugWindow.Visible;
        }

        #endregion
    }
}
