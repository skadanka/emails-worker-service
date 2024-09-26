using System;
using System.Collections.Generic;
using System.Linq;

public class ClassifyCategory
{
    // Property to store keyword lists
    public Dictionary<string, List<string>> KeywordDict { get; private set; }

    // Constructor to load keywords
    public ClassifyCategory()
    {
        KeywordDict = KeywordsLists.keywordDict;
    }

    // Method to classify a given text and return the category with the maximum score
    public string ClassifyText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            Console.WriteLine("Provided text is empty or null.");
            return string.Empty;
        }

        Dictionary<string, double> categoryScores = new Dictionary<string, double>();

        // Initialize scores and perform classification
        foreach (var entry in KeywordDict)
        {
            string category = entry.Key;
            List<string> keywords = entry.Value;

            double score = keywords.Count(keyword => text.Contains(keyword, StringComparison.OrdinalIgnoreCase));
            // Normalize the score by the number of keywords in the category
            categoryScores[category] = score / keywords.Count;
        }

        // Determine the category with the maximum score
        var maxScoreCategory = categoryScores.OrderByDescending(sc => sc.Value).FirstOrDefault();

        return maxScoreCategory.Key;  // Return the category with the highest score
    }
}
