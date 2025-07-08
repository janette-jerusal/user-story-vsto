using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelReader
{
    public static DataTable ReadSheetToDataTable(object sheetObj)
    {
        var sheet = (Excel.Worksheet)sheetObj;
        Excel.Range usedRange = sheet.UsedRange;
        DataTable dt = new DataTable();

        int rowCount = usedRange.Rows.Count;
        int colCount = usedRange.Columns.Count;

        // Add columns
        for (int c = 1; c <= colCount; c++)
            dt.Columns.Add(usedRange.Cells[1, c].Value?.ToString() ?? $"Column{c}");

        // Add rows
        for (int r = 2; r <= rowCount; r++)
        {
            var row = dt.NewRow();
            for (int c = 1; c <= colCount; c++)
                row[c - 1] = usedRange.Cells[r, c].Value?.ToString() ?? "";
            dt.Rows.Add(row);
        }

        return dt;
    }
}
