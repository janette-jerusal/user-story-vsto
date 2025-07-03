using System;
using System.Collections.Generic;
using System.Linq;

namespace UserStorySimilarityAddIn
{
    public static class UserStoryComparer
    {
        public static List<(string idA, string idB, double score)> CompareDataFrames(
            List<(string ID, string Desc)> df1,
            List<(string ID, string Desc)> df2,
            double threshold)
        {
            var results = new List<(string, string, double)>();
            foreach (var a in df1)
                foreach (var b in df2)
                {
                    double sim = ComputeSimilarity(a.Desc, b.Desc);
                    if (sim >= threshold)
                        results.Add((a.ID, b.ID, Math.Round(sim, 3)));
                }
            return results;
        }

        public static double ComputeSimilarity(string t1, string t2)
        {
            var v1 = GetVector(t1);
            var v2 = GetVector(t2);
            return CosineSimilarity(v1, v2);
        }

        private static Dictionary<string,int> GetVector(string text) =>
            text.Split(new[] {' ','.',',',';','!','?'}, StringSplitOptions.RemoveEmptyEntries)
                .Select(w => w.ToLowerInvariant())
                .GroupBy(w => w)
                .ToDictionary(g=>g.Key, g=>g.Count());

        private static double CosineSimilarity(Dictionary<string,int> v1, Dictionary<string,int> v2)
        {
            var all = new HashSet<string>(v1.Keys.Concat(v2.Keys));
            double dot = all.Sum(k => v1.GetValueOrDefault(k) * v2.GetValueOrDefault(k));
            double mag1 = Math.Sqrt(v1.Values.Sum(x=>x*x));
            double mag2 = Math.Sqrt(v2.Values.Sum(x=>x*x));
            return (mag1*mag2 == 0) ? 0 : dot/(mag1*mag2);
        }
    }
}
