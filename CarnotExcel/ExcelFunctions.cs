using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CarnotExcel
{
    public class ExcelFunctions : IExcelAddIn
    {
        public void AutoOpen() => IntelliSenseServer.Install();
        public void AutoClose() => IntelliSenseServer.Uninstall();

        [ExcelFunction(Name = "PRI.MAIÚSCULA.CARNOT", Description = "Função para colocar a primeira letra de cada palavra em maiúsculo ignorando algumas palavras")]
        public static string Capitalize([ExcelArgument(Name = "Texto", Description = "Texto para aplicar a função")] string name, [ExcelArgument(Name = "Palavras ignoradas", Description = "Lista de palavras adicionais aos casos padrões separadas por vírgula que serão ignoradas")] string wordsToIgnoreString = null)
        {
            if (string.IsNullOrEmpty(wordsToIgnoreString))
            {
                wordsToIgnoreString = "a,o,as,os,de,das,do,dos,da,e,ou,para,por,no,na,nos,nas,à";
            }
            else
            {
                wordsToIgnoreString = "a,o,as,os,de,das,do,dos,da,e,ou,para,por,no,na,nos,nas,à," + wordsToIgnoreString.ToLower();
            }

            string lowerString = name.ToLower();

            string[] words = lowerString.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            string[] wordsToIgnore = wordsToIgnoreString.Split(',');

            StringBuilder result = new StringBuilder();

            for (int i = 0; i < words.Length; i++)
            {
                string word = words[i];

                if (Array.IndexOf(wordsToIgnore, word) > -1 && i != 0)
                {
                    result.Append(word);
                }
                else
                {
                    result.Append(char.ToUpper(word[0]) + word.Substring(1));
                }

                if (i < words.Length - 1)
                {
                    result.Append(" ");
                }
            }

            return result.ToString();
        }

        [ExcelFunction(Name = "PROCURA.CARNOT", Description = "Função para realizar uma busca aproximada em uma matriz, identificando o valor mais semelhante ao termo buscado e indicando seu grau de similaridade.")]
        public static object[,] FuzzyMatchSearch([ExcelArgument(Name = "Valor a ser buscado", Description = "O valor que você deseja encontrar de forma aproximada na matriz. Pode ser um texto ou uma palavra-chave que esteja presente na matriz de dados.")] string searchValue, [ExcelArgument(Name = "Matriz de Dados", Description = "Matriz (intervalo de células) onde será realizada a busca pelo valor mais semelhante ao valor informado.")] object[,] dataMatrix)
        {
            List<double> proximity = new List<double>();

            object[,] result = new object[1, 2];

            for (int row = 0; row < dataMatrix.GetLength(0); row++)
            {
                proximity.Add(CompareTextAdvanced(searchValue, dataMatrix[row, 0].ToString()));
            }

            int index = proximity.IndexOf(proximity.Max());
            result[0, 0] = dataMatrix[index, 0];
            result[0, 1] = proximity.Max();

            return result;
        }

        public static double CompareTextAdvanced(string text1, string text2)
        {
            var tokens1 = ExtractRelevantWords(text1);
            var tokens2 = ExtractRelevantWords(text2);

            double totalSimilarity = 0;
            int totalWeight = 0;

            for (int i = 0; i < Math.Max(tokens1.Count, tokens2.Count); i++)
            {
                string token1 = i < tokens1.Count ? tokens1[i] : "";
                string token2 = i < tokens2.Count ? tokens2[i] : "";

                int weight = (i == 0) ? 5 : 1;
                totalWeight += weight;

                double levenshteinSim = 1 - (double)LevenshteinDistance(token1, token2) / Math.Max(token1.Length, token2.Length);
                double jaroWinklerSim = JaroWinklerDistance(token1, token2);

                double combinedSim = (levenshteinSim + jaroWinklerSim) / 2;

                totalSimilarity += combinedSim * weight;
            }

            return (totalSimilarity / totalWeight);
        }

        // Função de distância de Levenshtein
        private static int LevenshteinDistance(string a, string b)
        {
            int n = a.Length;
            int m = b.Length;
            int[,] dp = new int[n + 1, m + 1];

            for (int i = 0; i <= n; i++)
                dp[i, 0] = i;

            for (int j = 0; j <= m; j++)
                dp[0, j] = j;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                    dp[i, j] = Math.Min(Math.Min(dp[i - 1, j] + 1, dp[i, j - 1] + 1), dp[i - 1, j - 1] + cost);
                }
            }

            return dp[n, m];
        }

        public static double JaroWinklerDistance(string s1, string s2)
        {
            double jaroDistance = JaroDistance(s1, s2);

            int prefixLength = 0;
            for (int i = 0; i < Math.Min(s1.Length, s2.Length); i++)
            {
                if (s1[i] == s2[i])
                    prefixLength++;
                else
                    break;
            }
            prefixLength = Math.Min(4, prefixLength);

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

        private static List<string> ExtractRelevantWords(string text)
        {
            string[] commonWords = { "of", "the", "and", "state", "plurinational" };
            return text.Split(new char[] { ' ', ',', '(', ')' }, StringSplitOptions.RemoveEmptyEntries)
                       .Where(word => !commonWords.Contains(word.ToLower()))
                       .ToList();
        }
    }
}

