using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

public static class KeywordsLists
{
    public static Dictionary<string, List<string>> keywordDict;

    // Static constructor to initialize the dictionary
    static KeywordsLists()
    {
        LoadKeywordsFromFile(@"C:\Users\recruitment\source\repos\emails-worker service\emails-worker service\Candidate scoring\heb_eng_keywords");
    }

    private static void LoadKeywordsFromFile(string filePath)
    {
        try
        {
            // Ensure the file exists
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Keyword file not found.");
                return;
            }

            // Read JSON file into a string
            string jsonData = File.ReadAllText(filePath);

            // Deserialize the JSON string into a Dictionary
            keywordDict = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(jsonData);

            if (keywordDict == null)
            {
                Console.WriteLine("Failed to load keywords.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Public method to get keywords by category
    public static List<string> GetKeywordsByCategory(string category)
    {
        if (keywordDict != null && keywordDict.ContainsKey(category))
        {
            return keywordDict[category];
        }
        else
        {
            Console.WriteLine($"No keywords found for category: {category}");
            return new List<string>(); // Return an empty list if the category is not found
        }
    }
}
