using emails_worker_service.candidate_scroing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emails_worker_service.Candidate_scoring
{
    public class CandidateScoring
    {
        private readonly BertFeatureExtraction _featureExtractor;

        public CandidateScoring()
        {
            _featureExtractor = new BertFeatureExtraction();
        }

        public List<KeyValuePair<string, double>> RankCandidates(List<string> cvs, string jobDescription)
        {
            var cleanedJobDescription = TextPreprocessing.CleanText(jobDescription);
            var jobFeatures = _featureExtractor.ExtractFeatures(cleanedJobDescription);

            var scores = new List<KeyValuePair<string, double>>();
            foreach (var cv in cvs)
            {
                var cleanedCv = TextPreprocessing.CleanText(cv);
                var cvFeatures = _featureExtractor.ExtractFeatures(cleanedCv);
                var similarity = SimilarityCalculation.CosineSimilarity(jobFeatures, cvFeatures);
                scores.Add(new KeyValuePair<string, double>(cv, similarity));
            }

            return scores.OrderByDescending(s => s.Value).ToList();
        }
    }
}
