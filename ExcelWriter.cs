// ExcelWriter.cs
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelWriter
{
    public static void WriteSimilarities(Excel.Worksheet worksheet, List<string> results)
    {
        for (int i = 0; i < results.Count; i++)
        {
            worksheet.Cells[i + 1, 2].Value2 = results[i];
        }
    }
}
