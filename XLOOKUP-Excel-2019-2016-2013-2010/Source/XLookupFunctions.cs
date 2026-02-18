using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;

public class XLookupFunctions
{
    [ExcelFunction(Description = "Searches a range or an array, and returns an item corresponding to the first match it finds. If a match doesn't exist, then XLOOKUP can return the closest (approximate) match.")]
    public static object XLOOKUP(
        [ExcelArgument(Description = "The value to search for")] object lookup_value,
        [ExcelArgument(Description = "The array or range to search")] object lookup_array,
        [ExcelArgument(Description = "The array or range to return")] object return_array,
        [ExcelArgument(Description = "The value to return if not found")] object if_not_found,
        [ExcelArgument(Description = "0 - Exact match (default)\n-1 - Exact match or next smaller item\n1 - Exact match or next larger item\n2 - Wildcard match")] object match_mode,
        [ExcelArgument(Description = "1 - Search first-to-last (default)\n-1 - Search last-to-first\n2 - Binary search (sorted ascending)\n-2 - Binary search (sorted descending)")] object search_mode)
    {
        // 1. Parse Arguments
        // ------------------
        
        // Match Mode Defaults to 0 (Exact Match)
        int matchModeInt = 0;
        if (match_mode != null && !(match_mode is ExcelMissing))
        {
            if (match_mode is double d) matchModeInt = (int)d;
            else if (match_mode is int i) matchModeInt = i;
            // Handle parsing errors or other types if necessary, for now default to 0 or throw #VALUE
        }

        // Search Mode Defaults to 1 (First-to-Last)
        int searchModeInt = 1;
        if (search_mode != null && !(search_mode is ExcelMissing))
        {
            if (search_mode is double d) searchModeInt = (int)d;
            else if (search_mode is int i) searchModeInt = i;
        }

        // Handle arrays/ranges
        object[] lookup = Flatten(lookup_array);
        object[] results = Flatten(return_array);

        if (lookup.Length == 0) return ExcelError.ExcelErrorValue;
        
        // If return array is smaller, we can only return up to its length, or error? 
        // Excel XLOOKUP returns #N/A if sizes don't match usually or trims? 
        // Actually XLOOKUP requires lookup_array and return_array to be same size in one dimension usually.
        // For simplicity here, we assume 1:1 mapping based on index.
        int maxIndex = Math.Min(lookup.Length, results.Length);

        // 2. Search Logic
        // ---------------

        int foundIndex = -1;

        if (searchModeInt == 1 || searchModeInt == -1) // Linear Search
        {
            if (matchModeInt == 2) // Wildcard
            {
                 foundIndex = LinearSearchWildcard(lookup_value, lookup, searchModeInt, maxIndex);
            }
            else // Exact or Approximate Linear
            {
                 foundIndex = LinearSearch(lookup_value, lookup, matchModeInt, searchModeInt, maxIndex);
            }
        }
        else if (searchModeInt == 2 || searchModeInt == -2) // Binary Search
        {
            // Binary search requires sorted data. 
            // Implementation of binary search for XLOOKUP is complex due to type handling.
            // For this MVP, we will fallback to linear or implement a basic binary search if critical.
            // Given the complexity of implementing robust binary search across mixed types in C#, 
            // and the user likely just needs functional XLOOKUP, linear search is often sufficient 
            // but we should try to honor the request or fallback.
            
            // NOTE: Implementing robust generic binary search on object[] compatible with Excel's sort rules is non-trivial.
            // We'll map to linear for now or implement a simple version. 
            // Ideally: return LinearSearch... but efficient XLOOKUP relies on binary for speed.
            // For this task, correctness > speed. 
             foundIndex = LinearSearch(lookup_value, lookup, matchModeInt, 1, maxIndex); 
        }

        // 3. Return Result
        // ----------------

        if (foundIndex >= 0 && foundIndex < results.Length)
        {
            return results[foundIndex];
        }

        // Not Found
        if (if_not_found != null && !(if_not_found is ExcelMissing))
        {
            return if_not_found;
        }

        return ExcelError.ExcelErrorNA;
    }

    private static int LinearSearch(object value, object[] array, int matchMode, int searchMode, int length)
    {
        // Simple linear search implementation
        // matchMode: 0 = exact, -1 = exact or smaller, 1 = exact or larger
        
        int start = (searchMode == 1) ? 0 : length - 1;
        int end = (searchMode == 1) ? length : -1;
        int step = (searchMode == 1) ? 1 : -1;

        int bestMatchIndex = -1;
        // For approximate matches, we need to track the "closest" found so far
        // But XLOOKUP logic for -1/1 is specific:
        // 0: Must match exactly.
        // -1: Exact match, or else the largest item less than value.
        // 1: Exact match, or else the smallest item greater than value.
        
        // Optimization: For exact match (0), return immediately.
        // For others, we might need to scan all if not sorted? 
        // Actually XLOOKUP definition for -1/1 doesn't require sorted data? 
        // Wait, XLOOKUP *does not* require sorted data for -1/1 if doing linear search, 
        // it finds the *best* match in the direction specified? 
        // actually no, standard XLOOKUP approximate match usually expects sorted data for binary search,
        // but for linear search, does it find the *absolute* closest in the whole array?
        // Documentation says: 
        // "If match_mode is -1... looks for exact match... if not found, returns the next smaller item."
        // In a linear search (search_mode 1), does it mean the *first* item it encounters that is smaller? No.
        // It effectively searches for the value. 
        
        // Let's refine: The most common use case is Exact Match (0).
        
        if (matchMode == 0)
        {
            for (int i = start; i != end; i += step)
            {
                if (IsMatch(value, array[i])) return i;
            }
            return -1;
        }
        
        // For approximate match (-1 or 1) with linear search, XLOOKUP *does* define behavior.
        // But implementing full "closest value" logic over arbitrary types (strings vs numbers) is tricky.
        // Let's implement logic for Numbers specifically, as that's the 99% use case for approximate lookups.
        
        double targetVal = 0;
        bool isTargetNumber = IsNumber(value, out targetVal);
        
        // We need to find the "best" index.
        // If we find Exact match, return immediately.
        // Otherwise keep track of best.
        
        object bestVal = null;
        
        for (int i = start; i != end; i += step)
        {
            object current = array[i];
            
            // Check Exact
            if (IsMatch(value, current)) return i;
            
            // Check Approximate
            if (isTargetNumber && IsNumber(current, out double currentVal))
            {
                if (matchMode == -1) // Exact or next smaller
                {
                    if (currentVal < targetVal)
                    {
                        // We found a smaller item. Is it "better" (closer to target) than previous best?
                        // Or does XLOOKUP just take the *greatest* value that is < target? Yes.
                        // We need the max of value < target.
                        if (bestMatchIndex == -1 || currentVal > (double)bestVal) // bestVal holds the number
                        {
                            bestMatchIndex = i;
                            bestVal = currentVal;
                        }
                    }
                }
                else if (matchMode == 1) // Exact or next larger
                {
                     if (currentVal > targetVal)
                    {
                        // We need the min of value > target.
                        if (bestMatchIndex == -1 || currentVal < (double)bestVal)
                        {
                            bestMatchIndex = i;
                            bestVal = currentVal;
                        }
                    }
                }
            }
        }
        
        return bestMatchIndex;
    }
    
    private static int LinearSearchWildcard(object value, object[] array, int searchMode, int length)
    {
        string pattern = value?.ToString() ?? "";
        // Convert Excel wildcard to Regex? 
        // * -> .*, ? -> .
        // Minimal implementation required.
        // For now, simple exact string match if no wildcards, or simple startsWith/Contains check?
        // Proper wildcard support is complex regex. 
        // Let's try native VB-like operators or simplified Regex.
        
        string regexPattern = "^" + System.Text.RegularExpressions.Regex.Escape(pattern)
            .Replace("\\*", ".*")
            .Replace("\\?", ".") + "$";
            
        var regex = new System.Text.RegularExpressions.Regex(regexPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        int start = (searchMode == 1) ? 0 : length - 1;
        int end = (searchMode == 1) ? length : -1;
        int step = (searchMode == 1) ? 1 : -1;

        for (int i = start; i != end; i += step)
        {
             string s = array[i]?.ToString() ?? "";
             if (regex.IsMatch(s)) return i;
        }
        return -1;
    }

    private static bool IsMatch(object target, object candidate)
    {
        if (target == null && candidate == null) return true;
        if (target == null || candidate == null) return false;
        
        // Excel matches string case-insensitively usually
        if (target is string st && candidate is string sc)
        {
            return st.Equals(sc, StringComparison.OrdinalIgnoreCase);
        }
        
        // Handle numbers
        if (IsNumber(target, out double nt) && IsNumber(candidate, out double nc))
        {
            return Math.Abs(nt - nc) < 1e-9;
        }

        return target.Equals(candidate);
    }
    
    private static bool IsNumber(object val, out double d)
    {
        d = 0;
        if (val == null) return false;
        if (val is double dub) { d = dub; return true; }
        if (val is int i) { d = i; return true; }
        // Attempt parse string? Excel usually purely based on type, but "1" vs 1 behavior varies. 
        // XLOOKUP usually type sensitive unless value is number-stored-as-text issues.
        return false;
    }

    private static object[] Flatten(object arg)
    {
        if (arg is object[,] arr)
        {
            // Flatten 2D array (Range) to 1D
            // Row-major or Column-major? 
            // XLOOKUP works on vectors. If 1 row multiple cols -> flatten.
            // If multiple rows 1 col -> flatten.
            // If Matrix -> XLOOKUP usually errors or processes first row/col?
            // Actually XLOOKUP can return arrays.
            // But for Lookup Array, it must be a vector (1 row or 1 col).
            
            // Let's just list all items.
            List<object> list = new List<object>();
            foreach (var item in arr) list.Add(item);
            return list.ToArray();
        }
        if (arg is object[] vec) return vec;
        if (arg is ExcelMissing) return new object[0];
        return new object[] { arg }; // Single value
    }
}
