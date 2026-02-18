using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace ExcelFilterCopy
{
    public static class CopyEngine
    {
        public static void CopyFilteredWithFormat()
        {
            Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Worksheet ws = (Worksheet)app.ActiveSheet;

            // 1. Validate source selection
            if (!(app.Selection is Range srcRange))
            {
                MessageBox.Show("Select SOURCE (filtered) range first.", "Copy Filtered", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            Range visRange;
            try
            {
                visRange = srcRange.SpecialCells(XlCellType.xlCellTypeVisible);
            }
            catch
            {
                MessageBox.Show("No visible cells found.", "Copy Filtered", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (visRange == null)
            {
                MessageBox.Show("No visible cells found.", "Copy Filtered", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 2. Pick Destination Column
            Range destCell = null;
            try
            {
                object input = app.InputBox(
                    Prompt: "1. Scroll to destination\n2. Click ANY cell in the target column",
                    Title: "Select Destination Column",
                    Type: 8); // Type 8 returns Range

                if (input is Range r) destCell = r;
            }
            catch
            {
                return; // Cancelled
            }

            if (destCell == null) return;
            int destCol = destCell.Column;

            // 3. Performance optimizations
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            XlCalculation calc = app.Calculation;
            app.Calculation = XlCalculation.xlCalculationManual;

            HashSet<string> processedMerges = new HashSet<string>();

            try
            {
                // 4. Copy loop
                // Iterating cells in a large range can be slow. 
                // SpecialCells(xlCellTypeVisible) returns a multipart range (Areas).
                // It's better to iterate Areas first, then rows/cells in areas?
                // The VBA iterates "visRange.Cells".

                foreach (Range area in visRange.Areas)
                {
                    foreach (Range cell in area.Cells)
                    {
                        try
                        {
                            if ((bool)cell.MergeCells)
                            {
                                Range mergeArea = cell.MergeArea;
                                string addr = mergeArea.Address;

                                if (!processedMerges.Contains(addr))
                                {
                                    processedMerges.Add(addr);
                                    mergeArea.Copy();

                                    Range target = (Range)ws.Cells[mergeArea.Row, destCol];
                                    target.PasteSpecial(XlPasteType.xlPasteAll);

                                    // Re-merge destination if needed (PasteAll usually handles merge, but VBA re-merges explicitly)
                                    // VBA says: .Resize(...).Merge
                                    // PasteSpecial sometimes unmerges target? Or maybe just to be safe.
                                    // Let's follow VBA logic.
                                    target.Resize[mergeArea.Rows.Count, mergeArea.Columns.Count].Merge();
                                }
                            }
                            else
                            {
                                cell.Copy();
                                Range target = (Range)ws.Cells[cell.Row, destCol];
                                target.PasteSpecial(XlPasteType.xlPasteAll);
                            }
                        }
                        catch (Exception ex)
                        {
                            // In VBA: MsgBox "Paste failed..." -> GoTo Cleanup
                            // We will just log or continue? VBA shows MsgBox and exits.
                            app.ScreenUpdating = true;
                            MessageBox.Show($"Paste failed at row {cell.Row}. Destination may be protected.\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unexpected error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Cleanup
                app.CutCopyMode = 0; // False
                app.Calculation = calc;
                app.EnableEvents = true;
                app.ScreenUpdating = true;
            }
        }
    }
}
