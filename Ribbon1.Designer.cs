namespace UserStorySimilarityAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton compareButton;

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.compareButton = this.Factory.CreateRibbonButton();

            // Tab
            this.tab1.Label = "User Story Tools";
            this.tab1.Groups.Add(this.group1);

            // Group
            this.group1.Label = "Similarity";
            this.group1.Items.Add(this.compareButton);

            // Button
            this.compareButton.Label = "Compare User Stories";
            this.compareButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.compareButton_Click);

            this.Tabs.Add(this.tab1);
        }
    }
}

