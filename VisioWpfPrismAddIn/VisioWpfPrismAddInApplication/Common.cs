namespace VisioWpfPrismAddInApplication
{
    public class Common : VNC.VSTOAddIn.Common
    {
        public static Events.VisioAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static Microsoft.Office.Interop.Visio.Application VisioApplication { get; set; }
    }
}
