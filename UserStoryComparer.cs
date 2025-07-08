using System;
using System.Data;

namespace UserStorySimilarityAddIn
{
    public static class UserStoryComparer
    {
        public static DataTable Compare(DataTable df1, DataTable df2, double threshold = 0.5)
        {
            // Create result table
            DataTable result = new DataTable();
            result.Columns.Add("Story1");
            result.Columns.Add("Story2");
            result.Columns.Add("Similarity");

            // Dummy comparison for now
            foreach (DataRow row1 in df1.Rows)
            {
                foreach (DataRow row2 in df2.Rows)
                {
                    string s1 = row1["Desc"].ToString();
                    string s2 = row2["Desc"].ToString();

                    double similarity = DummySimilarity(s1, s2);

                    if (similarity >= threshold)
                    {
                        result.Rows.Add(s1, s2, similarity.ToString("F2"));
                    }
                }
            }

            return result;
        }

        // Dummy similarity logic (placeholder)
        private static double DummySimilarity(string a, string b)
        {
            return a.Equals(b, StringComparison.OrdinalIgnoreCase) ? 1.0 : 0.0;
        }
    }
}
