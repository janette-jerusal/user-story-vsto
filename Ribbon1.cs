using System;
using System.Data;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace UserStorySimilarityAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void CompareButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;

                if (workbook == null)
                {
                    MessageBox.Show("No workbook is open.");
                    return;
                }

                var sheets = workbook.Sheets;
                if (sheets.Count < 2)
                {
                    MessageBox.Show("Please ensure at least two sheets are present.");
                    return;
                }

                var df1 = ExcelReader.ReadSheetToDataTable(sheets[1]);
                var df2 = ExcelReader.ReadSheetToDataTable(sheets[2]);

                var result = UserStoryComparer.CompareUserStories(df1, df2, 0.5); // 0.5 = similarity threshold

                var newSheet = workbook.Sheets.Add();
                newSheet.Name = "ComparisonResult";

                ExcelWriter.WriteDataTableToSheet(result, newSheet);

                MessageBox.Show("Comparison complete!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}

