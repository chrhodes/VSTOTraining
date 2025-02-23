using System.Diagnostics;
using System;
using System.Windows;

namespace VisioAddInTemplate
{
    public class Common : VisioAddInApplicationTemplate.Common
    {
        new public const string LOG_CATEGORY = "VisioAddInTemplate";

        // NOTE(crhodes)
        // If we want to log anything in Ribbon_Load or ThisAddIn_Startup / ThisAddIn_Shutdown
        // we will have to face referencing VNC.Core, ugh.
        // Maybe just leave as commented out MessageBox.Show for when developing.
        // Everything else can be seen by updating DeveloperMode
    }
}
