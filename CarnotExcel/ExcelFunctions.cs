using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CarnotExcel
{
    public class ExcelFunctions : IExcelAddIn
    {
        public void AutoOpen() => IntelliSenseServer.Install();
        public void AutoClose() => IntelliSenseServer.Uninstall();

        [ExcelFunction(Name = "PRIMAIÚSCULA.CARNOT", Description = "Função para colocar a primeira letra de cada palavra em maiúsculo ignorando algumas palavras")]
        public static string capitalize([ExcelArgument(Name = "Texto", Description = "Texto para aplicar a função")] string name)
        {
            string lowerString = name.ToLower();

            string[] words = lowerString.Split(' ');

            string[] wordsToIgnore = new string[]
            {
                "a", "o", "as", "os", "de", "de", "das", "do", "da", "e", "ou", "para", "por", "no", "na", "nos", "nas", "dos"
            };

            string result = "";

            foreach (string word in words)
            {
                if (Array.IndexOf(wordsToIgnore, word) > -1)
                    result += $"{word} ";
                else
                    result += $"{word[0].ToString().ToUpper()}{word.Substring(1)} ";
            }

            return result.Trim();
        }

        [ExcelFunction(Name = "PROCURA.CARNOT", Description = "Função para uma busca aproximada")]
        public static object[,] NGramsDistance([ExcelArgument(Name = "Elemento buscado", Description = "Elemento que será buscado na outra matriz")] string searchValue, [ExcelArgument(Name = "Matriz", Description = "Matriz onde será procurado")] object[,] dataMatrix)
        {
            List<double> distances = new List<double>();

            object[,] result = new object[1, 2];

            for (int row = 0; row < dataMatrix.GetLength(0); row++)
            {
                distances.Add(CompareSophisticated(searchValue, dataMatrix[row, 0].ToString()));
            }

            int index = distances.IndexOf(distances.Max());
            result[0, 0] = dataMatrix[index, 0];
            result[0, 1] = distances.Max();
            
            return result;
        }

        [ExcelFunction(Description = "Retorna uma matriz 2D de valores.")]
        public static object[,] GenerateMatrix(int rows, int cols)
        {
            object[,] result = new object[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    result[i, j] = (i + 1) * (j + 1);
                }
            }

            return result;
        }

        public static double JaroWinklerDistance(string s1, string s2)
        {
            double jaroDistance = JaroDistance(s1, s2);

            // Penalização para strings curtas
            int prefixLength = 0;
            for (int i = 0; i < Math.Min(s1.Length, s2.Length); i++)
            {
                if (s1[i] == s2[i])
                    prefixLength++;
                else
                    break;
            }
            prefixLength = Math.Min(4, prefixLength); // O prefixo máximo considerado é de 4 caracteres

            return jaroDistance + (0.1 * prefixLength * (1 - jaroDistance));
        }

        private static double JaroDistance(string s1, string s2)
        {
            int s1_len = s1.Length;
            int s2_len = s2.Length;

            if (s1_len == 0 || s2_len == 0)
            {
                return 0;
            }

            int match_distance = Math.Max(s1_len, s2_len) / 2 - 1;

            bool[] s1_matches = new bool[s1_len];
            bool[] s2_matches = new bool[s2_len];

            int matches = 0;
            int transpositions = 0;

            for (int i = 0; i < s1_len; i++)
            {
                int start = Math.Max(0, i - match_distance);
                int end = Math.Min(i + match_distance + 1, s2_len);

                for (int j = start; j < end; j++)
                {
                    if (s2_matches[j]) continue;
                    if (s1[i] != s2[j]) continue;
                    s1_matches[i] = true;
                    s2_matches[j] = true;
                    matches++;
                    break;
                }
            }

            if (matches == 0)
            {
                return 0;
            }

            int k = 0;
            for (int i = 0; i < s1_len; i++)
            {
                if (!s1_matches[i]) continue;
                while (!s2_matches[k]) k++;
                if (s1[i] != s2[k]) transpositions++;
                k++;
            }

            transpositions /= 2;

            return ((double)matches / s1_len + (double)matches / s2_len + (double)(matches - transpositions) / matches) / 3.0;
        }

        private static int LevenshteinDistance(string a, string b)
        {
            a = a.ToLower();
            b = b.ToLower();

            int n = a.Length;
            int m = b.Length;
            int[,] dp = new int[n + 1, m + 1];

            for (int i = 0; i <= n; i++)
            {
                dp[i, 0] = i;
            }
            for (int j = 0; j <= m; j++)
            {
                dp[0, j] = j;
            }

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                    dp[i, j] = Math.Min(
                        Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1),
                        dp[i - 1, j - 1] + cost);
                }
            }

            return dp[n, m];
        }

        public static double CompareSophisticated(string text1, string text2)
        {
            text1 = text1.ToLower();
            text2 = text2.ToLower();

            string processedText1 = Preprocess(text1);
            string processedText2 = Preprocess(text2);

            var tokens1 = processedText1.Split(' ');
            var tokens2 = processedText2.Split(' ');

            double totalSimilarity = 0;
            int totalWeight = 0;

            for (int i = 0; i < Math.Max(tokens1.Length, tokens2.Length); i++)
            {
                string token1 = i < tokens1.Length ? tokens1[i] : "";
                string token2 = i < tokens2.Length ? tokens2[i] : "";

                int weight = (i == 0) ? 3 : 1;
                totalWeight += weight;

                totalSimilarity += (1 - (double)LevenshteinDistance(token1, token2) / Math.Max(token1.Length, token2.Length)) * weight;
            }

            return (totalSimilarity / totalWeight);
        }

        private static string Preprocess(string text)
        {
            return text.ToLower().Replace("(", "").Replace(")", "").Replace(",", "");
        }
    }
}

