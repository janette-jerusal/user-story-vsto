using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace UserStorySimilarityAddIn
{
    public static class ExcelReader
    {
        public static List<(string ID, string Desc)> ReadUserStories(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var list = new List<(string ID, string Desc)>();
            using (var pkg = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = pkg.Workbook.Worksheets[0];
                int rows = ws.Dimension.End.Row;
                for (int r = 2; r <= rows; r++)
                {
                    var id = ws.Cells[r, 1].Text;
                    var desc = ws.Cells[r, 2].Text;
                    if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(desc))
                        list.Add((id, desc));
                }
            }
            return list;
        }
    }
}
