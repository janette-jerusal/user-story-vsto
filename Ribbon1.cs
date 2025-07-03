using System;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;

namespace UserStorySimilarityAddIn
{
    public partial class Ribbon1 : RibbonBase
    {
        private void btnUploadAndCompare_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var dlg = new OpenFileDialog { Filter = "Excel Files|*.xlsx", Multiselect = true })
                {
                    if (dlg.ShowDialog() == DialogResult.OK && dlg.FileNames.Length == 2)
                    {
                        var df1 = ExcelReader.ReadUserStories(dlg.FileNames[0]);
                        var df2 = ExcelReader.ReadUserStories(dlg.FileNames[1]);
                        var results = UserStoryComparer.CompareDataFrames(df1, df2, 0.75);
                        ExcelWriter.WriteResultsToNewSheet(results);
                        MessageBox.Show("Done! Check the new worksheet.", "Success");
                    }
                    else
                        MessageBox.Show("Please select exactly two .xlsx files.", "Oops");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show($"Error: {ex.Message}", "Error");
            }
        }
    }
}
