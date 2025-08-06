using System;

using Microsoft.Office.Tools.Ribbon;

namespace TestVisioAddIn
{
    public partial class Ribbon
    {
        // NOTE(crhodes)
        // This was moved out of designer so we can log

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            Int64 startTicks = Common.WriteToDebugWindow("Ribbon()", true);
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
    }
}
