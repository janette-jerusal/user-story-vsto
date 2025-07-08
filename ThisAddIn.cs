using System;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace UserStorySimilarityAddIn
{
    public partial class ThisAddIn
    {
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return Globals.Factory.GetRibbonFactory().CreateRibbonManager(
                new IRibbonExtension[] { new Ribbon1(Globals.Factory.GetRibbonFactory()) });
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
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
