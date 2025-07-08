namespace UserStorySimilarityAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        private RibbonTab tab1;
        private RibbonGroup group1;
        internal RibbonButton compareButton;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.compareButton = this.Factory.CreateRibbonButton();

            // 
            // tab1
            // 
            this.tab1.Label = "UserStory Tools";
            this.tab1.Name = "tab1";
            this.tab1.Groups.Add(this.group1);

            // 
            // group1
            // 
            this.group1.Label = "Actions";
            this.group1.Name = "group1";
            this.group1.Items.Add(this.compareButton);

            // 
            // compareButton
            // 
            this.compareButton.Label = "Compare";
            this.compareButton.Name = "compareButton";
            this.compareButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CompareButton_Click);

            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        #endregion
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
