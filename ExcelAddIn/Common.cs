namespace ExcelAddIn
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "ExcelAddIn";

        public static Events.ExcelAppEvents AppEvents;
    }
}
