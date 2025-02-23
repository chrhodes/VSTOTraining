using System.Windows;
//using Microsoft.Office.Core;

namespace VisioAddInApplicationTemplate
{
    /// <summary>
    /// AddInInfo
    /// </summary>
    /// <remarks>
    /// This class can be used in two ways.  If calling this from a commandBar, modify
    /// the Private Const as needed and then create an instance of this class in the code
    /// that loads the command bars.
    /// 
    /// If calling this from a Ribbon Event handler, call the ActionNameGoesHere method directly.
    /// 
    /// Rename the ActionNameGoesHere Method and provide code that does something useful.
    /// </remarks>
    public class AddInInfo     {

        #region "Private Constants and Variables"

        private const string _MODULE_NAME = Common.LOG_CATEGORY + "AddInInfo";
        private const string _NAME = "AddInInfo";
        private const string _BITMAP_NAME = "AddInInfo.bmp";
        private const string _CAPTION = "AddInInfo";
        private const string _TOOL_TIP_TEXT = "Click for AddInInfo";
        private const string _DESCRIPTION = "AddInInfo does ...";
        #endregion

        #region "Public Methods"

        public static void DisplayInfo()
        {
            //AssemblyHelper.AssemblyInformation info = new AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetExecutingAssembly());

            // FIX(crhodes)
            // Isn't their stuff in VNC.Core that does this.  Can we deprecate VNC.AssemblyHelper

            VNC.AssemblyHelper.AssemblyInformation info = new VNC.AssemblyHelper.AssemblyInformation(System.Reflection.Assembly.GetCallingAssembly());
            MessageBox.Show(info.ToString());
        }

        #endregion

    }
}