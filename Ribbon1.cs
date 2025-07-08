using Microsoft.Office.Tools.Ribbon;
using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class Ribbon1 : RibbonBase
    {
        public Ribbon1(RibbonFactory factory) : base(factory)
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void compareButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                var sheets = workbook.Sheets;

                if (sheets.Count < 2)
                {
                    MessageBox.Show("Workbook must have at least 2 sheets.");
                    return;
                }

                Excel.Worksheet sheet1 = (Excel.Worksheet)sheets[1];
                Excel.Worksheet sheet2 = (Excel.Worksheet)sheets[2];

                DataTable df1 = ExcelReader.ReadSheetToDataTable(sheet1);
                DataTable df2 = ExcelReader.ReadSheetToDataTable(sheet2);

                DataTable result = UserStoryComparer.CompareUserStories(df1, df2, 0.5); // Change threshold if needed

                Excel.Worksheet resultSheet = (Excel.Worksheet)sheets.Add();
                resultSheet.Name = "Similarity Results";

                ExcelWriter.WriteDataTableToSheet(result, resultSheet);

                MessageBox.Show("Comparison complete. Results written to new sheet.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}

