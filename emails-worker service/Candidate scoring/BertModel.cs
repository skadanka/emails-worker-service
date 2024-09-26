using Microsoft.ML;
using Microsoft.ML.Data;
using Microsoft.ML.Transforms.Text;
using System;
using System.Collections.Generic;
using System.Linq;

namespace emails_worker_service.candidate_scroing
{

    public class BertFeatureExtraction
    {
        private readonly MLContext _mlContext;
        private readonly ITransformer _model;

        public BertFeatureExtraction()
        {
            _mlContext = new MLContext();
            var pipeline = _mlContext.Transforms.Text.TokenizeIntoWords("Tokens", "Text")
                .Append(_mlContext.Transforms.Text.ApplyWordEmbedding("Features", "Tokens",
                    WordEmbeddingEstimator.PretrainedModelKind.SentimentSpecificWordEmbedding));

            var emptyData = _mlContext.Data.LoadFromEnumerable(new List<Document>());
            _model = pipeline.Fit(emptyData);
        }

        public float[] ExtractFeatures(string text)
        {
            var data = _mlContext.Data.LoadFromEnumerable(new List<Document> { new Document { Text = text } });
            var transformedData = _model.Transform(data);
            var features = transformedData.GetColumn<float[]>("Features").FirstOrDefault();
            return features;
        }

        private class Document
        {
            public string Text { get; set; }
        }
    }
    

/*    public class Program
    {
        public static void Main()
        {
            // Sample data
            var cvs = new List<string> { "Candidate 1 CV text...", "Candidate 2 CV text...", "Candidate 3 CV text..." };
            var jobDescription = "Job description text...";

            // Rank candidates
            var candidateScoring = new CandidateScoring();
            var rankedCandidates = candidateScoring.RankCandidates(cvs, jobDescription);

            // Display ranked candidates
            foreach (var candidate in rankedCandidates)
            {
                Console.WriteLine($"CV: {candidate.Key}, Score: {candidate.Value}");
            }
        }
    }*/
}