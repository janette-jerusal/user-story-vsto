using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace UserStorySimilarityAddIn
{
    public static class ExcelWriter
    {
        public static void WriteResultsToNewSheet(List<(string idA, string idB, double score)> matches)
        {
            var app = Globals.ThisAddIn.Application;
            var sheet = app.Worksheets.Add();
            sheet.Name = "Similarity Results";
            sheet.Cells[1,1] = "Story A ID";
            sheet.Cells[1,2] = "Story B ID";
            sheet.Cells[1,3] = "Similarity Score";
            for (int i = 0; i < matches.Count; i++)
            {
                sheet.Cells[i+2, 1] = matches[i].idA;
                sheet.Cells[i+2, 2] = matches[i].idB;
                sheet.Cells[i+2, 3] = matches[i].score;
            }
        }
    }
}
