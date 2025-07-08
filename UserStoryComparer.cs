using System;
using System.Data;
using System.Linq;
using Accord.MachineLearning.Text;
using Accord.Math.Distances;

namespace UserStorySimilarityAddIn
{
    public static class UserStoryComparer
    {
        public static DataTable CompareUserStories(DataTable df1, DataTable df2, double threshold)
        {
            var combined = df1.Rows.Cast<DataRow>().Select(r => r["Desc"].ToString())
                          .Concat(df2.Rows.Cast<DataRow>().Select(r => r["Desc"].ToString()))
                          .ToArray();

            var tfidf = new TfIdfVectorizer().Learn(combined);
            var vectors = combined.Select(s => tfidf.Transform(s)).ToArray();

            int len1 = df1.Rows.Count;
            int len2 = df2.Rows.Count;
            var results = new DataTable();

            results.Columns.Add("Story A ID");
            results.Columns.Add("Story A Desc");
            results.Columns.Add("Story B ID");
            results.Columns.Add("Story B Desc");
            results.Columns.Add("Similarity Score");

            var cosine = new Cosine();

            for (int i = 0; i < len1; i++)
                for (int j = 0; j < len2; j++)
                {
                    double sim = 1.0 - cosine.Distance(vectors[i], vectors[len1 + j]);
                    if (sim >= threshold)
                    {
                        results.Rows.Add(
                            df1.Rows[i]["ID"], df1.Rows[i]["Desc"],
                            df2.Rows[j]["ID"], df2.Rows[j]["Desc"],
                            Math.Round(sim, 3)
                        );
                    }
                }

            return results;
        }
    }
}
