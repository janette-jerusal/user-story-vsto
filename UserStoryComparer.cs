using System;
using System.Data;
using System.Linq;

public static class UserStoryComparer
{
    public static DataTable CompareUserStories(DataTable df1, DataTable df2, double threshold)
    {
        var result = new DataTable();
        result.Columns.Add("Story1");
        result.Columns.Add("Story2");
        result.Columns.Add("Similarity");

        foreach (DataRow r1 in df1.Rows)
        {
            string s1 = r1["Desc"].ToString();

            foreach (DataRow r2 in df2.Rows)
            {
                string s2 = r2["Desc"].ToString();
                double sim = JaccardSimilarity(s1, s2);

                if (sim >= threshold)
                    result.Rows.Add(s1, s2, sim.ToString("0.00"));
            }
        }

        return result;
    }

    private static double JaccardSimilarity(string a, string b)
    {
        var set1 = a.Split(' ').Select(x => x.ToLower()).ToHashSet();
        var set2 = b.Split(' ').Select(x => x.ToLower()).ToHashSet();

        var intersection = set1.Intersect(set2).Count();
        var union = set1.Union(set2).Count();

        return union == 0 ? 0 : (double)intersection / union;
    }
}
