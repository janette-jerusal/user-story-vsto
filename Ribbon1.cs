// Ribbon1.cs
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

public partial class Ribbon1 : RibbonBase
{
    public Ribbon1() : base(Globals.Factory.GetRibbonFactory())
    {
        InitializeComponent();
    }

    private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

    private void CompareButton_Click(object sender, RibbonControlEventArgs e)
    {
        Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
        var stories = ExcelReader.ReadUserStories(sheet);
        var results = UserStoryComparer.Compare(stories);
        ExcelWriter.WriteSimilarities(sheet, results);
    }
}

