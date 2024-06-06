using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Visio;

namespace VisioAddIn
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VisioAddIn";

        public static Events.VisioAppEvents AppEvents;
    }
}
