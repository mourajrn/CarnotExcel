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
            // Preprocessar os textos (remover parênteses, ignorar maiúsculas/minúsculas)
            string processedText1 = Preprocess(text1);
            string processedText2 = Preprocess(text2);

            // Separar os textos em tokens (palavras)
            var tokens1 = ExtractRelevantWords(processedText1);
            var tokens2 = ExtractRelevantWords(processedText2);

            // Comparar as palavras principais com Levenshtein e Jaro-Winkler
            double totalSimilarity = 0;
            int totalWeight = 0;

            for (int i = 0; i < Math.Max(tokens1.Count, tokens2.Count); i++)
            {
                string token1 = i < tokens1.Count ? tokens1[i] : "";
                string token2 = i < tokens2.Count ? tokens2[i] : "";

                // Ponderar a primeira palavra como a mais importante (nome principal)
                int weight = (i == 0) ? 5 : 1; // Dar peso 5 para a primeira palavra (nome principal)
                totalWeight += weight;

                // Calcular a similaridade combinada usando Levenshtein e Jaro-Winkler
                double levenshteinSim = (1 - (double)LevenshteinDistance(token1, token2) / Math.Max(token1.Length, token2.Length));
                double jaroWinklerSim = JaroWinklerDistance(token1, token2);

                // Combinar as duas similaridades com pesos iguais
                double combinedSim = (levenshteinSim + jaroWinklerSim) / 2;

                // Adicionar à similaridade total, aplicando o peso
                totalSimilarity += combinedSim * weight;
            }

            // Retornar a similaridade como percentual
            return (totalSimilarity / totalWeight);
        }

        // Função de pré-processamento (remover parênteses, ignorar maiúsculas/minúsculas)
        private static string Preprocess(string text)
        {
            return text.ToLower().Replace("(", "").Replace(")", "").Replace(",", "");
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

        // Função que calcula a distância de Jaro-Winkler
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

        // Função que implementa a distância de Jaro
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

        // Função auxiliar para remover palavras irrelevantes como "State", "of", "the"
        private static List<string> ExtractRelevantWords(string text)
        {
            string[] commonWords = { "of", "the", "and", "state", "plurinational" };
            return text.Split(new char[] { ' ', ',', '(', ')' }, StringSplitOptions.RemoveEmptyEntries)
                       .Where(word => !commonWords.Contains(word.ToLower()))
                       .ToList();
        }
    }
}

