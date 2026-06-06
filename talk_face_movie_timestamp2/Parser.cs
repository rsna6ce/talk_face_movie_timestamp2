using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace BudouX
{
    public class Parser
    {
        private readonly Dictionary<string, Dictionary<string, int>> model;

        public Parser(Dictionary<string, Dictionary<string, int>> model)
        {
            this.model = model;
        }

        private int GetScore(string featureKey, string sequence)
        {
            if (model.TryGetValue(featureKey, out var group) && group.TryGetValue(sequence, out var score))
            {
                return score;
            }
            return 0;
        }

        public List<string> Parse(string sentence)
        {
            if (string.IsNullOrEmpty(sentence))
            {
                return new List<string>();
            }

            var result = new List<string> { sentence[0].ToString() };
            int totalScore = model.Values.Sum(group => group.Values.Sum());
            double baseScore = -totalScore * 0.5;

            for (int i = 1; i < sentence.Length; i++)
            {
                double score = baseScore;

                if (i > 2)
                {
                    score += GetScore("UW1", sentence.Substring(i - 3, 1));
                }
                if (i > 1)
                {
                    score += GetScore("UW2", sentence.Substring(i - 2, 1));
                }
                score += GetScore("UW3", sentence.Substring(i - 1, 1));
                score += GetScore("UW4", sentence.Substring(i, 1));
                if (i + 1 < sentence.Length)
                {
                    score += GetScore("UW5", sentence.Substring(i + 1, 1));
                }
                if (i + 2 < sentence.Length)
                {
                    score += GetScore("UW6", sentence.Substring(i + 2, 1));
                }

                if (i > 1)
                {
                    score += GetScore("BW1", sentence.Substring(i - 2, 2));
                }
                score += GetScore("BW2", sentence.Substring(Math.Max(0, i - 1), Math.Min(2, sentence.Length - Math.Max(0, i - 1))));
                if (i + 1 < sentence.Length)
                {
                    score += GetScore("BW3", sentence.Substring(i, Math.Min(2, sentence.Length - i)));
                }

                if (i > 2)
                {
                    score += GetScore("TW1", sentence.Substring(i - 3, 3));
                }
                if (i > 1 && i + 1 < sentence.Length)
                {
                    score += GetScore("TW2", sentence.Substring(i - 2, 3));
                }
                if (i + 2 < sentence.Length)
                {
                    score += GetScore("TW3", sentence.Substring(Math.Max(0, i - 1), Math.Min(3, sentence.Length - Math.Max(0, i - 1))));
                }
                if (i + 2 < sentence.Length)
                {
                    score += GetScore("TW4", sentence.Substring(i, Math.Min(3, sentence.Length - i)));
                }

                if (score > 0)
                {
                    result.Add(sentence[i].ToString());
                }
                else
                {
                    result[result.Count - 1] += sentence[i];
                }
            }

            return result;
        }

        public static Parser LoadDefaultJapaneseParser(string modelPath)
        {
            var json = File.ReadAllText(modelPath);
            var model = ParseJsonModel(json);
            return new Parser(model);
        }
        
        private static Dictionary<string, Dictionary<string, int>> ParseJsonModel(string json)
        {
            var model = new Dictionary<string, Dictionary<string, int>>();
            
            var featureGroupPattern = new Regex("\"([^\"]+)\"\\s*:\\s*\\{([^}]+)\\}");
            var featureMatches = featureGroupPattern.Matches(json);
            
            foreach (Match featureMatch in featureMatches)
            {
                if (featureMatch.Groups.Count >= 3)
                {
                    string featureKey = featureMatch.Groups[1].Value;
                    string featureContent = featureMatch.Groups[2].Value;
                    
                    var featureDict = new Dictionary<string, int>();
                    var featurePattern = new Regex("\"([^\"]+)\"\\s*:\\s*(-?\\d+)");
                    var matches = featurePattern.Matches(featureContent);
                    
                    foreach (Match match in matches)
                    {
                        if (match.Groups.Count >= 3)
                        {
                            string key = match.Groups[1].Value;
                            int value = int.Parse(match.Groups[2].Value);
                            featureDict[key] = value;
                        }
                    }
                    
                    model[featureKey] = featureDict;
                }
            }
            
            return model;
        }

        public string TranslateHtmlString(string html)
        {
            string textContent = html; // 実際にはHTMLからテキストを抽出
            var chunks = Parse(textContent);
            return string.Join("\u200b", chunks); // ゼロ幅スペースで区切る
        }
    }
}
