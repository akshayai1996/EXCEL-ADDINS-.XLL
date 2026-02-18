using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace ReplaceManyAddIn
{
    public static class ReplaceManyFunctions
    {
        [ExcelFunction(Description = "Bulk full-word replacement using a mapping table")]
        public static object REPLACE_MANY(
            [ExcelArgument(Description = "Input Range")] object InputRange,
            [ExcelArgument(Description = "Map Range (2 columns: From, To)")] object MapRange,
            [ExcelArgument(Description = "Optional Delimiters string")] object Delims)
        {
            try
            {
                object[,] inArr = To2D(InputRange);
                int rowsCount = inArr.GetLength(0);
                int colsCount = inArr.GetLength(1);

                object[,] mapArr = To2D(MapRange);
                
                // Build Dictionary
                Dictionary<string, string> dict = BuildDictionary(mapArr, false); // Default case-insensitive for UDF? VBA says: case-insensitive by default.
                if (dict.Count == 0) return inArr; // No map, return input

                string delimiters = (Delims is string s && !string.IsNullOrEmpty(s)) ? s : DefaultDelims();

                // Process
                object[,] outArr = new object[rowsCount, colsCount];

                for (int r = 0; r < rowsCount; r++)
                {
                    for (int c = 0; c < colsCount; c++)
                    {
                        outArr[r, c] = ReplaceInCell(inArr[r, c], dict, delimiters);
                    }
                }

                return outArr;
            }
            catch (Exception)
            {
                return ExcelError.ExcelErrorValue; 
            }
        }

        private static object[,] To2D(object input)
        {
            if (input is object[,] arr2d) return arr2d;
            if (input is object[] arr1d)
            {
                // Convert 1D to 2D col
                object[,] res = new object[arr1d.Length, 1];
                for (int i = 0; i < arr1d.Length; i++) res[i, 0] = arr1d[i];
                return res;
            }
            if (input is ExcelMissing || input == null) return new object[1, 1] { { "" } }; // Or error?
            
            // Scalar
            return new object[1, 1] { { input } };
        }

        private static Dictionary<string, string> BuildDictionary(object[,] mapArr, bool caseSensitive)
        {
            // Dictionary logic: Sort keys by length desc
            var list = new List<Tuple<string, object, int>>();
            
            int rowStart = mapArr.GetLowerBound(0);
            int rowEnd = mapArr.GetUpperBound(0);
            int colStart = mapArr.GetLowerBound(1);
            int cols = mapArr.GetLength(1);

            if (cols < 2) return new Dictionary<string, string>();

            for (int i = rowStart; i <= rowEnd; i++)
            {
                object kObj = mapArr[i, colStart];
                if (kObj == null || kObj is ExcelError || kObj is ExcelMissing) continue;
                string k = kObj.ToString().Trim();
                if (string.IsNullOrEmpty(k)) continue;

                object vObj = mapArr[i, colStart + 1];
                
                list.Add(new Tuple<string, object, int>(k, vObj, k.Length));
            }

            // Sort descending by length
            list.Sort((a, b) => b.Item3.CompareTo(a.Item3));

            StringComparer comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
            var dict = new Dictionary<string, string>(comparer);

            foreach (var item in list)
            {
                if (!dict.ContainsKey(item.Item1))
                {
                    dict.Add(item.Item1, item.Item2?.ToString() ?? "");
                }
            }
            return dict;
        }

        private static string DefaultDelims()
        {
            return " ,.;:!?'\"()[]{}<>/\\|-_+=*&^%$#@~`\n\t";
        }

        public static object ReplaceInCell(object val, Dictionary<string, string> dict, string delims)
        {
            if (val == null || val is ExcelError || val is ExcelMissing) return val;
            string s = val.ToString();
            if (string.IsNullOrEmpty(s)) return "";

            // Pad delimiters
            // Efficiency warning: complicated string manipulation in loop
            // Optimization: Tokenize directly?
            // VBA logic: Pad delims with spaces, Replace, Then remove spaces.
            
            // To faithfully port, we follow the logic:
            // 1. Pad delims
            // 2. Split by space
            // 3. Replace tokens
            // 4. Join
            // 5. Unpad delims
            
            // Note: This logic assumes delimiters are single chars.
            
            StringBuilder sb = new StringBuilder(" " + s + " ");
            foreach (char c in delims)
            {
                sb.Replace(c.ToString(), " " + c + " ");
            }

            string[] tokens = sb.ToString().Split(new char[] { ' ' }, StringSplitOptions.None);
            
            for (int i = 0; i < tokens.Length; i++)
            {
                if (tokens[i].Length > 0)
                {
                    if (dict.TryGetValue(tokens[i], out string replacement))
                    {
                        tokens[i] = replacement;
                    }
                }
            }

            string joined = string.Join(" ", tokens);
            
            StringBuilder res = new StringBuilder(joined);
            // Reverse pad
            foreach (char c in delims)
            {
                res.Replace(" " + c, c.ToString());
            }

            // Tidy open brackets logic from VBA
            string openers = "([{<";
            foreach (char c in openers)
            {
                res.Replace(c + " ", c.ToString());
            }

            return res.ToString().Trim();
        }
        
        // Expose helper for Controller
        public static Dictionary<string, string> GetDictionary(object[,] mapArr, bool caseSensitive)
        {
            return BuildDictionary(mapArr, caseSensitive);
        }
        
        public static string GetDefaultDelims() { return DefaultDelims(); }
    }
}
