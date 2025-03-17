using System.Windows;

namespace VA1Application
{
    public class AddInInfo
    {

        //#region "Private Constants and Variables"

        //private const string _MODULE_NAME = Common.LOG_CATEGORY + "AddInInfo";
        //private const string _NAME = "AddInInfo";
        //private const string _BITMAP_NAME = "AddInInfo.bmp";
        //private const string _CAPTION = "AddInInfo";
        //private const string _TOOL_TIP_TEXT = "Click for AddInInfo";
        //private const string _DESCRIPTION = "AddInInfo does ...";

        //#endregion

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