using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            MessageBox.Show("UserStorySimilarityAddIn Loaded!");
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Cleanup if needed
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

