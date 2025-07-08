using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

public static class ExcelWriter
{
    public static void WriteDataTableToSheet(DataTable dt, object sheetObj)
    {
        var sheet = (Excel.Worksheet)sheetObj;

        // Headers
        for (int c = 0; c < dt.Columns.Count; c++)
            sheet.Cells[1, c + 1] = dt.Columns[c].ColumnName;

        // Data
        for (int r = 0; r < dt.Rows.Count; r++)
            for (int c = 0; c < dt.Columns.Count; c++)
                sheet.Cells[r + 2, c + 1] = dt.Rows[r][c];
    }
}
