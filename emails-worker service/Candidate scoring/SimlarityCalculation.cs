using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace emails_worker_service.Candidate_scoring
{
    public class SimilarityCalculation
    {
        public static double CosineSimilarity(float[] vectorA, float[] vectorB)
        {
            double dotProduct = 0.0;
            double normA = 0.0;
            double normB = 0.0;
            for (int i = 0; i < vectorA.Length; i++)
            {
                dotProduct += vectorA[i] * vectorB[i];
                normA += Math.Pow(vectorA[i], 2);
                normB += Math.Pow(vectorB[i], 2);
            }
            return dotProduct / (Math.Sqrt(normA) * Math.Sqrt(normB));
        }
    }
}
