using System;
using System.Diagnostics;
using System.Windows;

namespace VA1Application
{
    public class Common //: VNC.Core.Common
    {
        new public const string LOG_CATEGORY = "VA1Application";

        public static Boolean EnableAppEvents = false;  // Custom Header and Footer need this enabled.
        public static Boolean DisplayEvents = false;
        public static Boolean DisplayChattyEvents = false;

        public static Events.VisioAppEvents AppEvents;
        public static Events.AddInApplicationEvents AddInApplicationEvents;

        public static Microsoft.Office.Interop.Visio.Application VisioApplication { get; set; }

        public static Boolean EnableLogging
        {
            get;
            set;
        }

        public static Boolean DeveloperMode
        {
            get;
            set;
        }

        private static Presentation.frmDebugWindow _DebugWindow;
        public static Presentation.frmDebugWindow DebugWindow
        {
            get
            {
                if (_DebugWindow == null)
                {
                    _DebugWindow = new Presentation.frmDebugWindow();
                }

                return _DebugWindow;
            }
            set
            {
                _DebugWindow = value;
            }
        }

        private static Presentation.frmWatchWindow _WatchWindow;
        public static Presentation.frmWatchWindow WatchWindow
        {
            get
            {
                if (_WatchWindow == null)
                {
                    _WatchWindow = new Presentation.frmWatchWindow();
                }
                return _WatchWindow;
            }
            set
            {
                _WatchWindow = value;
            }
        }

        public static Visibility DeveloperUIMode
        {
            get;
            set;
        }

        public static long WriteToWatchWindow(string message)
        {
            if (DeveloperMode)
            {
                WatchWindow.AddOutputLine(message);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToWatchWindow(string message, long startTicks)
        {
            if (DeveloperMode)
            {
                WatchWindow.AddOutputLine(message + "-" + (Stopwatch.GetTimestamp() - startTicks) / Stopwatch.Frequency);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToDebugWindow(string message)
        {
            if (DeveloperMode)
            {
                DebugWindow.AddOutputLine(message);
            }

            return Stopwatch.GetTimestamp();
        }

        public static long WriteToDebugWindow(string message, long startTicks)
        {

            if (DeveloperMode)
            {
                DebugWindow.AddOutputLine(message + "-" + (Stopwatch.GetTimestamp() - startTicks) / Stopwatch.Frequency);
            }

            return Stopwatch.GetTimestamp();
        }

    }
}
