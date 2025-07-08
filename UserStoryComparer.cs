// UserStoryComparer.cs
using System;
using System.Collections.Generic;
using System.Linq;

public static class UserStoryComparer
{
    public static List<string> Compare(List<string> stories)
    {
        var results = new List<string>();
        for (int i = 0; i < stories.Count; i++)
        {
            string comparison = $"Story {i + 1} is similar to: ";
            var similar = new List<string>();

            for (int j = 0; j < stories.Count; j++)
            {
                if (i == j) continue;
                if (stories[i].Split(' ').Intersect(stories[j].Split(' ')).Count() > 2)
                {
                    similar.Add((j + 1).ToString());
                }
            }
            comparison += string.Join(", ", similar);
            results.Add(comparison);
        }
        return results;
    }
}
