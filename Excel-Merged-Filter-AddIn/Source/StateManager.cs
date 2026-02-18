using System;
using System.Collections.Generic;

namespace ExcelMergedFilter
{
    public static class StateManager
    {
        // Tracks selected values for the current filter operation
        // Equivalent to: Public gSelectedValues As Object
        public static Dictionary<string, bool> SelectedValues { get; set; }

        // Tracks hidden rows to restore them later: RowIndex -> IsHidden
        // Equivalent to: Private gOriginalHiddenState As Object
        public static Dictionary<int, bool> OriginalHiddenState { get; set; } = new Dictionary<int, bool>();

        // Tracks which columns are filtered and what values they are filtered by
        // Equivalent to: Private gFilteredColumns As Object
        // ColumnIndex -> List of Selected Values
        public static Dictionary<int, List<string>> FilteredColumns { get; set; } = new Dictionary<int, List<string>>();

        public static void InitTrackers()
        {
            if (OriginalHiddenState == null) OriginalHiddenState = new Dictionary<int, bool>();
            if (FilteredColumns == null) FilteredColumns = new Dictionary<int, List<string>>();
        }

        public static void ClearSelectedValues()
        {
            SelectedValues = null;
        }

        public static void ResetAndClear()
        {
            SelectedValues = null;
            OriginalHiddenState = new Dictionary<int, bool>();
            FilteredColumns = new Dictionary<int, List<string>>();
        }
    }
}
