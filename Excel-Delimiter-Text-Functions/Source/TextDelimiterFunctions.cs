using System;
using ExcelDna.Integration;

namespace TextDelimiterAddIn
{
    public static class TextDelimiterFunctions
    {
        [ExcelFunction(Description = "Returns the text to the left of the Nth occurrence of a delimiter.")]
        public static string TextLeft(
            [ExcelArgument(Description = "Text to process. Ex: 'A-B-C'")] string txt,
            [ExcelArgument(Description = "Delimiter. Ex: '-'")] string delim,
            [ExcelArgument(Description = "Occurrence to stop at. Ex: 2 -> 'A-B'")] int n)
        {
            if (string.IsNullOrEmpty(txt) || string.IsNullOrEmpty(delim) || n <= 0)
                return "";

            int pos = GetNthDelimiterPosition(txt, delim, n);
            if (pos == -1) return "";

            return txt.Substring(0, pos).Trim();
        }

        [ExcelFunction(Description = "Returns the text to the right of the Nth occurrence of a delimiter.")]
        public static string TextRight(
            [ExcelArgument(Description = "Text to process. Ex: 'A-B-C'")] string txt,
            [ExcelArgument(Description = "Delimiter. Ex: '-'")] string delim,
            [ExcelArgument(Description = "Occurrence to start from. Ex: 1 -> 'B-C'")] int n)
        {
            if (string.IsNullOrEmpty(txt) || string.IsNullOrEmpty(delim) || n <= 0)
                return "";

            int pos = GetNthDelimiterPosition(txt, delim, n);
            if (pos == -1) return "";

            return txt.Substring(pos + delim.Length).Trim();
        }

        [ExcelFunction(Description = "Returns the text between the N1th and N2th occurrence of a delimiter.")]
        public static string TextMid(
            [ExcelArgument(Description = "Text to process. Ex: 'A-B-C'")] string txt,
            [ExcelArgument(Description = "Delimiter. Ex: '-'")] string delim,
            [ExcelArgument(Description = "Start occurrence. Ex: 1")] int n1,
            [ExcelArgument(Description = "End occurrence. Ex: 2 -> 'B'")] int n2)
        {
            if (string.IsNullOrEmpty(txt) || string.IsNullOrEmpty(delim) || n1 <= 0 || n2 <= n1)
                return "";

            int p1 = GetNthDelimiterPosition(txt, delim, n1);
            int p2 = GetNthDelimiterPosition(txt, delim, n2);

            if (p1 == -1 || p2 == -1) return "";

            int start = p1 + delim.Length;
            int length = p2 - start;

            if (length < 0) return "";

            return txt.Substring(start, length).Trim();
        }

        private static int GetNthDelimiterPosition(string txt, string delim, int n)
        {
            int pos = -1;
            for (int i = 0; i < n; i++)
            {
                pos = txt.IndexOf(delim, pos + 1, StringComparison.Ordinal); // Case-sensitive or Ordinal? VBA was Binary (Ordinal).
                if (pos == -1) return -1;
            }
            return pos;
        }
    }
}
