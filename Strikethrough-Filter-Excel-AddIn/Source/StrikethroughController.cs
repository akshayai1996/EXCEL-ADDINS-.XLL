using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace StrikethroughFilterAddIn
{
    [ComVisible(true)]
    public class StrikethroughController : ExcelRibbon
    {
        private const string HELPER_MARKER = "___STRIKE_FILTER_HELPER___";
        private const string FLAG_KEEP = "KEEP";
        private const string FLAG_HIDE = "HIDE";

        public override string GetCustomUI(string RibbonID)
        {
            return @"
              <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabStrikethrough' label='Filter Tools'>
                    <group id='grpStrike' label='Strikethrough'>
                      <button id='btnToggleStrike' label='Toggle Filter' onAction='OnToggleFilter' size='large' imageMso='Filter' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnToggleFilter(IRibbonControl control)
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet is Excel.Worksheet ws)
            {
                ToggleStrikethroughFilter(app, ws);
            }
        }

        private void ToggleStrikethroughFilter(Excel.Application app, Excel.Worksheet ws)
        {
            // 1. Check if already applied (Toggle OFF)
            int helperCol = FindHelperColumn(ws);
            if (helperCol > 0)
            {
                try
                {
                    app.ScreenUpdating = false;
                    if (ws.AutoFilterMode)
                    {
                        if (ws.FilterMode) ws.ShowAllData();
                        ws.AutoFilterMode = false;
                    }

                    // Retrieve state to clear comment
                    Excel.Range stateCol = (Excel.Range)ws.Cells[2, helperCol];
                    Excel.Range stateRow = (Excel.Range)ws.Cells[3, helperCol];
                    if (stateCol.Value2 is double c && stateRow.Value2 is double r)
                    {
                         Excel.Range headerCell = (Excel.Range)ws.Cells[(int)r, (int)c];
                         if (headerCell.Comment != null) headerCell.Comment.Delete();
                    }
                    ((Excel.Range)ws.Columns[helperCol]).Delete();
                    app.StatusBar = false;
                }
                finally
                {
                    app.ScreenUpdating = true;
                }
                return;
            }

            // 2. Toggle ON
            try
            {
                if (app.ActiveCell == null) return;
                Excel.Range activeCell = app.ActiveCell;
                int targetCol = activeCell.Column;
                int headerRow;

                Excel.Range selection = app.Selection as Excel.Range;
                // Auto-detect header row
                if (selection != null && selection.Rows.Count == 1 && selection.Columns.Count == ws.Columns.Count)
                {
                     // User selected entire row
                     headerRow = selection.Row;
                }
                else
                {
                     // Default to CurrentRegion top row
                     headerRow = activeCell.CurrentRegion.Row;
                }

                int lastRow = ((Excel.Range)ws.Cells[ws.Rows.Count, targetCol]).End[Excel.XlDirection.xlUp].Row;
                if (lastRow <= headerRow) 
                {
                     // Fallback/Safety: If CurrentRegion failed or single cell, maybe user meant active cell is header?
                     // But let's try to be smart. If 1 cell, and lastRow <= headerRow, implies empty or 1-row data.
                     // Let's just return for now to avoid crashes.
                     return;
                }

                helperCol = ((Excel.Range)ws.Cells[headerRow, ws.Columns.Count]).End[Excel.XlDirection.xlToLeft].Column + 1; // Check from HeaderRow level

                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                // Loop and check strikethrough
                // Range indexing in C# is typically [row, col] but for Range objects it's Range["A1"] or Cells[r, c]
                Excel.Range dataRng = ws.Range[ws.Cells[headerRow + 1, targetCol], ws.Cells[lastRow, targetCol]];
                object[,] arr = new object[dataRng.Rows.Count, 1];
                
                // Read individually (slow but safe for formatting)
                int r = 0;
                // dataRng.Cells is an IEnumerable
                foreach (Excel.Range cell in dataRng.Cells)
                {
                     // Explicit cast to bool? or check for DBNull
                     object strike = cell.Font.Strikethrough;
                     bool isStrike = false;
                     if (strike is bool b) isStrike = b;
                     // Excel returns DBNull or null for mixed, true/false otherwise.
                     
                     arr[r, 0] = isStrike ? FLAG_KEEP : FLAG_HIDE;
                     r++;
                }

                // Header
                Excel.Range headerCell = (Excel.Range)ws.Cells[headerRow, targetCol];
                // Visual Indicator (Comment)
                if (headerCell.Comment != null) headerCell.Comment.Delete();
                headerCell.AddComment("Filtered by Strikethrough");
                
                ((Excel.Range)ws.Cells[headerRow, helperCol]).Value = "FILTER";
                
                // Write data
                Excel.Range writeRng = ((Excel.Range)ws.Cells[headerRow + 1, helperCol]).Resize[arr.GetLength(0), 1];
                writeRng.Value = arr;
                
                // Marker and State
                ((Excel.Range)ws.Cells[1, helperCol]).Value = HELPER_MARKER;
                ((Excel.Range)ws.Cells[2, helperCol]).Value = targetCol; // Save target col
                ((Excel.Range)ws.Cells[3, helperCol]).Value = headerRow; // Save header row

                // Apply Filter
                Excel.Range filterRng = ((Excel.Range)ws.Cells[headerRow, helperCol]).Resize[arr.GetLength(0) + 1, 1];
                filterRng.AutoFilter(1, FLAG_KEEP);
                
                ((Excel.Range)ws.Columns[helperCol]).Hidden = true;
                
                // StatusBar
                app.StatusBar = "Strikethrough Filter ON (Col " + targetCol + ")";

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.ScreenUpdating = true;
            }
        }

        private int FindHelperColumn(Excel.Worksheet ws)
        {
            // Scanning row 1 for marker
            // UsedRange logic
            Excel.Range used = ws.UsedRange; 
            int maxCol = used.Column + used.Columns.Count + 5; 

            for (int c = 1; c <= maxCol; c++)
            {
                Excel.Range cell = (Excel.Range)ws.Cells[1, c];
                string val = Convert.ToString(cell.Value2);
                if (val == HELPER_MARKER)
                {
                    return c;
                }
            }
            return 0;
        }
    }
}
