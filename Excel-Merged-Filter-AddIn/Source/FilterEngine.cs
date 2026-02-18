using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelMergedFilter
{
    public static class FilterEngine
    {
        public static void ApplyFilter(Range specificTarget = null)
        {
            StateManager.InitTrackers();

            Range headerCell = specificTarget;
            Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)specificTarget?.Application ?? (Microsoft.Office.Interop.Excel.Application)ExcelDna.Integration.ExcelDnaUtil.Application;

            if (headerCell == null)
            {
                try
                {
                    object result = app.InputBox("Select Header:", "Filter v9.2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                    if (result is Range r)
                    {
                        headerCell = r;
                    }
                    else
                    {
                        return; // Cancelled or invalid
                    }
                }
                catch
                {
                    return; // InputBox cancelled
                }
            }

            if (headerCell == null) return;

            Worksheet ws = (Worksheet)headerCell.Worksheet;
            int colIndex = headerCell.Column;

            app.ScreenUpdating = false;

            try
            {
                // 2. Snapshot -> Reset -> Re-apply Others -> Edit Target
                var savedFilters = new Dictionary<int, List<string>>(StateManager.FilteredColumns);
                ClearFilter(ws, false);

                List<string> targetPreSelections = null;

                foreach (var kvp in savedFilters)
                {
                    if (kvp.Key != colIndex)
                    {
                        Range otherHeader = (Range)ws.Cells[headerCell.Row, kvp.Key];
                        Internal_FilterBlocksWithDropdown(kvp.Value, otherHeader, true);
                    }
                    else
                    {
                        targetPreSelections = kvp.Value;
                    }
                }

                Internal_FilterBlocksWithDropdown(targetPreSelections, headerCell, false);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        public static void ClearFilter(Worksheet ws, bool showMsg = true)
        {
            StateManager.InitTrackers();

            if (StateManager.OriginalHiddenState != null && StateManager.OriginalHiddenState.Count > 0)
            {
                foreach (var kvp in StateManager.OriginalHiddenState)
                {
                    if (kvp.Value == false)
                    {
                        try { ((Range)ws.Rows[kvp.Key]).Hidden = false; } catch { }
                    }
                }
                StateManager.OriginalHiddenState = new Dictionary<int, bool>();
            }

            ExcelUtils.RemoveAllMarkersAndRestoreFormats(ws);
            StateManager.FilteredColumns = new Dictionary<int, List<string>>();

            try
            {
                Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)ws.Application;
                Microsoft.Office.Interop.Excel.Window win = app.ActiveWindow;
                win.FreezePanes = false;
                win.SplitRow = 0;
            }
            catch { }

            if (showMsg) MessageBox.Show("Filters cleared.", "Filter v9.2", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void Internal_FilterBlocksWithDropdown(List<string> preSelections, Range headerCell, bool silentMode)
        {
            Worksheet ws = (Worksheet)headerCell.Worksheet;
            int headerRow = headerCell.Row;
            int filterCol = headerCell.Column;

            Range usedRange = ws.UsedRange;
            int lastRow = usedRange.Row + usedRange.Rows.Count - 1;

            if (lastRow <= headerRow) return;

            // 1. Optimized Structural Scan (X-Ray Vision)
            string[] blockValues = new string[lastRow + 1]; // 1-based index to match Excel rows roughly

            // We can read the column in one go for performance, but we need to handle MergeArea.
            // Reading cell by cell is slow in COM.
            // Optimization: Read value array, then for each cell check if new value or continues.
            // But VBA code uses MergeArea.Cells(1,1).Value.
            // In C#, getting MergeArea for every cell is EXPENSIVE.
            // Optimization: Read the whole column values and merge cells info if possible.
            // Or just stick to VBA logic for correctness first.

            // To speed up, we can use Range.Value2 to get 2D array.
            Range colRange = (Range)ws.Range[ws.Cells[headerRow + 1, filterCol], ws.Cells[lastRow, filterCol]];
            object[,] values = (object[,])colRange.Value2; // 1-based array relative to range

            // We still need to know about Merged Cells.
            // Iterating range cell by cell is slow.
            // Faster: Iterate 1 to count, check if cell is merged.
            // BUT checking .MergeCells property is also a COM call.
            // So we will just do the loop. It might be slower than VBA or similar.

            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                Range c = (Range)ws.Cells[r, filterCol];
                // Optimization: Handle MergeArea only if needed?
                // The VBA does: If c.MergeCells Then Value = c.MergeArea.Cells(1,1).Value
                // In Excel, if you access a cell in a merged area that is NOT the top-left, Value is empty?
                // No, Value is empty for non-top-left cells in a merged area.
                // So if Value is null/empty, it MIGHT be part of a merge area, or just empty.
                // If it is part of a merge area, we need the value of the top-left.

                // Let's stick to VBA logic:
                string val = "";
                if ((bool)c.MergeCells)
                {
                    Range ma = c.MergeArea;
                    Range tl = (Range)ma.Cells[1, 1];
                    object v = tl.Value2;
                    val = v != null ? v.ToString().Trim() : "";
                }
                else
                {
                    object v = c.Value2;
                    val = v != null ? v.ToString().Trim() : "";
                }
                blockValues[r] = val;
            }

            // 2. Collect Intersection Options
            SortedDictionary<string, bool> uniqueSet = new SortedDictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            for (int r = headerRow + 1; r <= lastRow; r++)
            {
                Range row = (Range)ws.Rows[r];
                if (!(bool)row.Hidden)
                {
                    string v = blockValues[r];
                    if (!string.IsNullOrEmpty(v))
                    {
                        if (!uniqueSet.ContainsKey(v)) uniqueSet.Add(v, true);
                    }
                }
            }

            if (uniqueSet.Count == 0) return;

            // 3. Selection Logic
            Dictionary<string, bool> selectedSet = null;

            if (silentMode)
            {
                selectedSet = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
                if (preSelections != null)
                {
                    foreach (var s in preSelections) selectedSet[s] = true;
                }
            }
            else
            {
                StateManager.SelectedValues = null;

                // Show Form
                // Must run on UI thread? Excel Add-in is on UI thread.
                using (FilterForm frm = new FilterForm())
                {
                    frm.InitWithKeys(uniqueSet.Keys);
                    if (preSelections != null) frm.SetSelections(preSelections);

                    DialogResult dr = frm.ShowDialog(); // Modular
                    if (dr == DialogResult.OK)
                    {
                        selectedSet = StateManager.SelectedValues; // Returns Dict<string, bool>
                    }
                }

                if (selectedSet == null) return;
            }

            // 4. Apply Batch Hiding
            Range hideRange = null;
            int batchCount = 0;
            Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)ws.Application;
            XlCalculation calc = app.Calculation;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = XlCalculation.xlCalculationManual;

                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    Range row = (Range)ws.Rows[r];
                    if (!(bool)row.Hidden)
                    {
                        string val = blockValues[r];
                        if (val != null && !selectedSet.ContainsKey(val)) // Hide if NOT selected (and not empty? VBA hides everything not selected, but uniqueSet only includes non-empty. If blockValues is empty, it's not in uniqueSet, so not in selectedSet -> Hidden?)
                        {
                            // VBA: If Not selectedSet.exists(blockValues(r)) Then Hide.
                            // If blockValues(r) is "", it's not in uniqueSet (VBA: If Len > 0 Then Add).
                            // So empty rows are Hidden?
                            // Yes, in VBA logic: If Not exists -> Hide.

                            // Track Original State
                            if (!StateManager.OriginalHiddenState.ContainsKey(r))
                            {
                                StateManager.OriginalHiddenState[r] = (bool)row.Hidden; // Should be false here
                            }

                            if (hideRange == null) hideRange = row;
                            else hideRange = app.Union(hideRange, row);

                            batchCount++;
                            if (batchCount >= 50)
                            {
                                if (hideRange != null) { hideRange.EntireRow.Hidden = true; hideRange = null; }
                                batchCount = 0;
                            }
                        }
                    }
                }
                if (hideRange != null) hideRange.EntireRow.Hidden = true;

                if (!silentMode)
                {
                    ws.Activate();
                    ((Range)ws.Rows[headerRow]).Hidden = false;
                    ExcelUtils.SafeFreezeTopRows(ws, headerRow);
                }

                ExcelUtils.MarkColumnFiltered(ws, headerRow, filterCol, selectedSet.Keys.ToArray());
            }
            finally
            {
                app.Calculation = calc;
                // app.ScreenUpdating = true; // Handled by caller or outer block
            }
        }
    }
}
