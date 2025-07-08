using System;
using System.Data;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace UserStorySimilarityAddIn
{
    public partial class Ribbon1 : OfficeRibbon, Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e) { }

        public void OnCompareClick(Office.IRibbonControl control)
        {
            var ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            var used = ws.UsedRange;
            var dt1 = new DataTable();
            var dt2 = new DataTable();

            dt1.Columns.Add("ID");
            dt1.Columns.Add("Desc");
            dt2.Columns.Add("ID");
            dt2.Columns.Add("Desc");

            int rows = used.Rows.Count;
            int mid = rows / 2;

            for (int i = 2; i <= mid; i++)
                dt1.Rows.Add(
                    (used.Cells[i, 1] as Excel.Range)?.Text?.ToString() ?? "",
                    (used.Cells[i, 2] as Excel.Range)?.Text?.ToString() ?? ""
                );

            for (int i = mid + 1; i <= rows; i++)
                dt2.Rows.Add(
                    (used.Cells[i, 1] as Excel.Range)?.Text?.ToString() ?? "",
                    (used.Cells[i, 2] as Excel.Range)?.Text?.ToString() ?? ""
                );

            var result = UserStoryComparer.CompareUserStories(dt1, dt2, 0.75);

            var ns = Globals.ThisAddIn.Application.Worksheets.Add();
            ns.Name = "Similarity Results";

            for (int c = 0; c < result.Columns.Count; c++)
                ns.Cells[1, c + 1] = result.Columns[c].ColumnName;

            for (int r = 0; r < result.Rows.Count; r++)
                for (int c = 0; c < result.Columns.Count; c++)
                    ns.Cells[r + 2, c + 1] = result.Rows[r][c]?.ToString();
        }
    }
}


