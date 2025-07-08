using System;

namespace UserStorySimilarityAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e) { }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }
    }
}
