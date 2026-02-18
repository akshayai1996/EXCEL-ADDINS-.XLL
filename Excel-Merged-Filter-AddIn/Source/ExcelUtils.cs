using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using ExcelDna.Integration;
using System.Runtime.InteropServices;

namespace ExcelMergedFilter
{
    public static class ExcelUtils
    {
        private const int HEADER_FILL_COLOR = 13434879; // RGB(255, 255, 204)
        private const int HEADER_FONT_COLOR = 0; // vbBlack
        private const string SHAPE_PREFIX = "MacroFilterMarker_";
        private const string HELPER_SHEET_NAME = "__MF_Helper";

        public static Worksheet GetHelperSheet(Workbook wb)
        {
            try
            {
                return (Worksheet)wb.Worksheets[HELPER_SHEET_NAME];
            }
            catch
            {
                Worksheet ws = (Worksheet)wb.Worksheets.Add();
                ws.Name = HELPER_SHEET_NAME;
                ws.Visible = XlSheetVisibility.xlSheetVeryHidden;
                return ws;
            }
        }

        public static void SaveHeaderFormatIfNotMarked(Worksheet ws, int rIndex, int cIndex)
        {
            if (HasHeaderMarkerShape(ws, cIndex)) return;

            Worksheet helper = GetHelperSheet((Workbook)ws.Parent);
            int rowKey = EnsureHelperRow(helper, ws);

            try
            {
                // Copy formats
                Range source = (Range)ws.Cells[rIndex, cIndex];
                Range dest = (Range)helper.Cells[rowKey, cIndex];
                dest.ClearFormats();
                source.Copy();
                dest.PasteSpecial(XlPasteType.xlPasteFormats);
                ((Application)ws.Application).CutCopyMode = 0;
            }
            catch { }
        }

        public static void RestoreHeaderFormatFromHelper(Worksheet ws, int rIndex, int cIndex)
        {
            Worksheet helper = GetHelperSheet((Workbook)ws.Parent);
            int rowKey = EnsureHelperRow(helper, ws);

            try
            {
                Range source = (Range)helper.Cells[rowKey, cIndex];
                Range dest = (Range)ws.Cells[rIndex, cIndex];
                source.Copy();
                dest.PasteSpecial(XlPasteType.xlPasteFormats);
                ((Application)ws.Application).CutCopyMode = 0;
            }
            catch { }
        }

        private static int EnsureHelperRow(Worksheet helper, Worksheet ws)
        {
            string key = ((Workbook)ws.Parent).Name + "!" + ws.Name;
            int lastRow = helper.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            // Allow for empty sheet case or just scan
            // For simplicity and porting exactly, we linear scan column A
            // Optimization: Filter or Match could be faster but linear is fine for helper
            for (int r = 1; r <= lastRow; r++)
            {
                Range cell = (Range)helper.Cells[r, 1];
                if (cell.Value != null && cell.Value.ToString().Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    return r;
                }
            }

            int newRow = lastRow + 1;
            // If sheet was empty, lastRow might be 1 but empty. Check.
            if (lastRow == 1 && ((Range)helper.Cells[1, 1]).Value == null) newRow = 1;
            else if (lastRow > 0 && ((Range)helper.Cells[lastRow, 1]).Value != null) newRow = lastRow + 1;

            ((Range)helper.Cells[newRow, 1]).Value = key;
            return newRow;
        }

        public static void SafeFreezeTopRows(Worksheet ws, int splitR)
        {
            try
            {
                Application app = (Application)ws.Application;
                Window win = app.ActiveWindow;
                win.FreezePanes = false;
                win.SplitRow = splitR;
                win.FreezePanes = true;
            }
            catch { }
        }

        public static void AddHeaderShapeMarker(Worksheet ws, Range hdr, string sName)
        {
            try
            {
                // Delete existing first
                try { ws.Shapes.Item(sName).Delete(); } catch { }

                double s = 7;
                double l = (double)hdr.Left + (double)hdr.Width - s - 2;
                double t = (double)hdr.Top + 2;

                Shape shp = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeIsoscelesTriangle, (float)l, (float)t, (float)s, (float)s);
                shp.Name = sName;
                shp.Rotation = 180;
                shp.Fill.ForeColor.RGB = 6579300; // RGB(100, 100, 100) -> 100 + 100*256 + 100*65536 = 100+25600+6553600 = 6579300
                shp.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                shp.Placement = XlPlacement.xlMoveAndSize;
            }
            catch { }
        }

        public static bool HasHeaderMarkerShape(Worksheet ws, int cIndex)
        {
            try
            {
                Shape shp = ws.Shapes.Item(SHAPE_PREFIX + cIndex);
                return shp != null;
            }
            catch
            {
                return false;
            }
        }

        public static void AddOrReplaceNote(Range tgt, string text)
        {
            try
            {
                if (tgt.Comment != null) tgt.Comment.Delete();
                tgt.AddComment(text);
                tgt.Comment.Shape.TextFrame.AutoSize = true;
            }
            catch { }
        }

        public static void RemoveOurNote(Range tgt)
        {
            try
            {
                if (tgt.Comment != null)
                {
                    if (tgt.Comment.Text().IndexOf("Filtered:", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        tgt.Comment.Delete();
                    }
                }
            }
            catch { }
        }

        public static void AddAmberCF(Range hdr)
        {
            try
            {
                RemoveAmberCF(hdr);
                FormatCondition fc = (FormatCondition)hdr.FormatConditions.Add(XlFormatConditionType.xlExpression, Formula1: "=TRUE");
                fc.StopIfTrue = true;
                fc.SetFirstPriority();
                fc.Interior.Color = HEADER_FILL_COLOR;
                fc.Font.Color = HEADER_FONT_COLOR;
            }
            catch { }
        }

        public static void RemoveAmberCF(Range hdr)
        {
            try
            {
                int count = hdr.FormatConditions.Count;
                for (int i = count; i >= 1; i--)
                {
                    FormatCondition fc = (FormatCondition)hdr.FormatConditions[i];
                    if (fc.Type == (int)XlFormatConditionType.xlExpression) // VBA check Type=xlExpression
                    {
                        if (fc.Formula1 == "=TRUE")
                        {
                            fc.Delete();
                        }
                    }
                }
            }
            catch { }
        }

        public static void RemoveAllMarkersAndRestoreFormats(Worksheet ws)
        {
            List<string> toDelete = new List<string>();
            try
            {
                foreach (Shape shp in ws.Shapes)
                {
                    if (shp.Name.StartsWith(SHAPE_PREFIX))
                    {
                        int headerRow = shp.TopLeftCell.Row;
                        int colIndex = shp.TopLeftCell.Column;
                        Range hdr = (Range)ws.Cells[headerRow, colIndex];

                        RemoveAmberCF(hdr);
                        RestoreHeaderFormatFromHelper(ws, headerRow, colIndex);
                        RemoveOurNote(hdr);

                        toDelete.Add(shp.Name);
                    }
                }
            }
            catch { }

            foreach (string name in toDelete)
            {
                try { ws.Shapes.Item(name).Delete(); } catch { }
            }
        }

        public static void MarkColumnFiltered(Worksheet ws, int rIndex, int cIndex, string[] selectedValues)
        {
            Range hdr = (Range)ws.Cells[rIndex, cIndex];

            // Save state
            List<string> vals = new List<string>(selectedValues);
            if (StateManager.FilteredColumns == null) StateManager.FilteredColumns = new Dictionary<int, List<string>>();
            StateManager.FilteredColumns[cIndex] = vals;

            SaveHeaderFormatIfNotMarked(ws, rIndex, cIndex);

            hdr.Interior.Pattern = XlPattern.xlPatternSolid;
            hdr.Interior.Color = HEADER_FILL_COLOR;
            hdr.Font.Color = HEADER_FONT_COLOR;

            AddAmberCF(hdr);
            AddOrReplaceNote(hdr, "Filtered:\n{" + string.Join(", ", selectedValues) + "}");

            AddHeaderShapeMarker(ws, hdr, SHAPE_PREFIX + cIndex);
        }
    }
}
