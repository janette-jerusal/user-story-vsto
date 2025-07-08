// ExcelReader.cs
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelReader
{
    public static List<string> ReadUserStories(Excel.Worksheet worksheet)
    {
        var stories = new List<string>();
        int row = 1;
        while (worksheet.Cells[row, 1].Value2 != null)
        {
            stories.Add(worksheet.Cells[row, 1].Value2.ToString());
            row++;
        }
        return stories;
    }
}
