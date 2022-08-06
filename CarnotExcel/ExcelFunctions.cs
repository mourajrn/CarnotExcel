using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using System;

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
    }
}
