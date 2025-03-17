using System;
using System.Windows.Forms;

namespace VA1Application.Presentation
{
    public partial class frmWatchWindow : Form
    {
        public frmWatchWindow()
        {
            InitializeComponent();
        }

        internal void AddOutputLine(string outputLine)
        {
            txtOutput.AppendText(outputLine + Environment.NewLine);

            if (Common.EnableLogging)
            {
                // TODO(crhodes)
                // Handle where to send logging output
                //Log.Trace(outputLine, Common.LOG_CATEGORY);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtOutput.Clear();
        }

        private void btnSaveToFile_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Not Implemented Yet");
        }

        private void frmWatchWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            Common.WatchWindow = null;
        }
    }
}
