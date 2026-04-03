using System.Windows;
//using Microsoft.Office.Core;

namespace VisioAddInApplicationTemplate
{

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