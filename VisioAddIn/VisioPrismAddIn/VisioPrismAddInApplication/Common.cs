using System.Windows;

using Prism.Events;

namespace VisioPrismAddInApplication
{
    public class Common : VNC.VSTOAddIn.Common
    {
        new public const string LOG_CATEGORY = "VisioPrismAddInApplication";

        //public static Boolean EnableAppEvents = false;  // Custom Header and Footer need this enabled.
        //public static Boolean DisplayEvents = false;
        //public static Boolean DisplayChattyEvents = false;

        public static Events.VisioAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static IEventAggregator EventAggregator = new EventAggregator();
        public static VisioPrismAddInApplication.Bootstrapper ApplicationBootstrapper;

        public static Microsoft.Office.Interop.Visio.Application VisioApplication { get; set; }

        //public static Boolean EnableLogging
        //{
        //    get;
        //    set;
        //}

        //public static Boolean DeveloperMode
        //{
        //    get;
        //    set;
        //}

        //private static frmDebugWindow _DebugWindow;
        //public static frmDebugWindow DebugWindow
        //{
        //    get
        //    {
        //        if (_DebugWindow == null)
        //        {
        //            _DebugWindow = new frmDebugWindow();
        //        }

        //        return _DebugWindow;
        //    }
        //    set
        //    {
        //        _DebugWindow = value;
        //    }
        //}

        //private static frmWatchWindow _WatchWindow;
        //public static frmWatchWindow WatchWindow
        //{
        //    get
        //    {
        //        if (_WatchWindow == null)
        //        {
        //            _WatchWindow = new frmWatchWindow();
        //        }
        //        return _WatchWindow;
        //    }
        //    set
        //    {
        //        _WatchWindow = value;
        //    }
        //}

        //public static Visibility DeveloperUIMode
        //{
        //    get;
        //    set;
        //}

        //public static long WriteToWatchWindow(string message)
        //{
        //    if (DeveloperMode)
        //    {
        //        WatchWindow.AddOutputLine(message);
        //    }

        //    return Stopwatch.GetTimestamp();
        //}

        //public static long WriteToWatchWindow(string message, long startTicks)
        //{
        //    if (DeveloperMode)
        //    {
        //        WatchWindow.AddOutputLine(message + "-" + (Stopwatch.GetTimestamp() - startTicks) / Stopwatch.Frequency);
        //    }

        //    return Stopwatch.GetTimestamp();
        //}

        //public static long WriteToDebugWindow(string message)
        //{
        //    if (DeveloperMode)
        //    {
        //        DebugWindow.AddOutputLine(message);
        //    }

        //    return Stopwatch.GetTimestamp();
        //}

        //public static long WriteToDebugWindow(string message, long startTicks)
        //{

        //    if (DeveloperMode)
        //    {
        //        DebugWindow.AddOutputLine(message + "-" + (Stopwatch.GetTimestamp() - startTicks) / Stopwatch.Frequency);
        //    }

        //    return Stopwatch.GetTimestamp();
        //}

    }
}
