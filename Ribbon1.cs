using Microsoft.Office.Tools.Ribbon;
using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        private void compareButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                var sheets = workbook.Sheets;

                if (sheets.Count < 2)
                {
                    MessageBox.Show("Workbook must have at least 2 sheets.");
                    return;
                }

                var df1 = ExcelReader.ReadSheetToDataTable(sheets[1]);
                var df2 = ExcelReader.ReadSheetToDataTable(sheets[2]);

                var result = UserStoryComparer.CompareUserStories(df1, df2, 0.5);

                ExcelWriter.WriteDataTableToSheet(result, sheets.Add());
                MessageBox.Show("Comparison complete!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}

