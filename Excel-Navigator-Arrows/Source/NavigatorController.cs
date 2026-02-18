using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace NavigatorArrowsAddIn
{
    [ComVisible(true)]
    public class NavigatorController : ExcelRibbon
    {
        private const string NAV_SHEET = "Navigator_List";
        private const int DATA_START_ROW = 6;
        private const float ARROW_WIDTH_PTS = 12f;
        private const string HEADER_ARROW_LEFT_NAME = "Nav_Header_Arrow_L";
        private const string HEADER_ARROW_RIGHT_NAME = "Nav_Header_Arrow_R";
        internal const string PASSWORD_RANGE_NAME = "_NavPassword"; // accessible to lifecycle class

        private string _currentPassword; // password for the current session

        // Event handlers
        private Excel.AppEvents_SheetSelectionChangeEventHandler _selectionChangeHandler;
        private Excel.AppEvents_SheetBeforeDoubleClickEventHandler _doubleClickHandler;
        private Excel.AppEvents_WorkbookBeforeCloseEventHandler _beforeCloseHandler;

        // --- Ribbon UI definition ---
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab idMso='TabAddIns' visible='false' />
                  <tab id='tabNavigator' label='Navigator Tools'>
                    <group id='grpNav' label='Focus Mode'>
                      <button id='btnStart' label='Start Focus' onAction='OnStartFocus' size='large' imageMso='Filter' />
                      <button id='btnStop' label='Stop Focus' onAction='OnStopFocus' size='large' imageMso='FilterClear' />
                      <separator id='sep1' />
                      <button id='btnPrev' label='Previous' onAction='OnPrevRow' size='large' imageMso='LeftArrow2' />
                      <button id='btnNext' label='Next' onAction='OnNextRow' size='large' imageMso='RightArrow2' />
                      <separator id='sep2' />
                      <button id='btnShowList' label='Show List' onAction='OnShowList' size='normal' imageMso='TableStyleClear' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        // --- Ribbon button handlers ---
        public void OnStartFocus(IRibbonControl control)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet == null) return;
            Excel.Worksheet activeSheet = (Excel.Worksheet)app.ActiveSheet;

            Excel.Range rngFilter = GetFilterRangeOrRegion(activeSheet);
            if (rngFilter == null)
            {
                System.Windows.Forms.MessageBox.Show("Place cursor inside filtered data.", "Navigator Arrows",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                return;
            }

            // Capture the ACTIVE column (where the user clicked)
            int activeCol = app.ActiveCell.Column;

            BuildNavigatorList(rngFilter, activeSheet, activeCol);
            UpdateHeaderArrow(activeSheet, activeCol, rngFilter.Row);
            HookEvents(app);

            // Protect the workbook with a random password and store it
            ProtectWorkbookWithRandomPassword(app.ActiveWorkbook);
        }

        public void OnStopFocus(IRibbonControl control)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveSheet == null) return;
            Excel.Worksheet activeSheet = (Excel.Worksheet)app.ActiveSheet;

            RemoveHeaderArrow(activeSheet);

            try
            {
                // 1. Unprotect the workbook using the stored password
                UnprotectWorkbook(app.ActiveWorkbook);

                // 2. Find and delete the Navigator_List sheet
                Excel.Worksheet navSh = null;
                try { navSh = (Excel.Worksheet)app.Worksheets[NAV_SHEET]; } catch { }

                if (navSh != null)
                {
                    // Retrieve the data sheet name before deleting
                    string sheetName = Convert.ToString(((Excel.Range)navSh.Range["F1"]).Value);
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        try
                        {
                            Excel.Worksheet dataSh = (Excel.Worksheet)app.Worksheets[sheetName];
                            if (dataSh != null)
                            {
                                dataSh.Rows.Hidden = false;
                                RemoveHeaderArrow(dataSh);
                            }
                        }
                        catch { }
                    }

                    // Delete the navigator sheet
                    navSh.Visible = Excel.XlSheetVisibility.xlSheetVisible; // ensure visible before delete
                    app.DisplayAlerts = false; // Disable alerts to prevent delete prompt
                    try { navSh.Delete(); } finally { app.DisplayAlerts = true; }
                }

                // 3. Remove the hidden password named range
                try
                {
                    Excel.Name passwordName = app.ActiveWorkbook.Names.Item(PASSWORD_RANGE_NAME);
                    if (passwordName != null)
                        passwordName.Delete();
                }
                catch { }

                // 4. Clear session password and unhook events
                _currentPassword = null;
                UnhookEvents(app);

                app.StatusBar = false;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error stopping focus: " + ex.Message, "Navigator");
            }
        }

        public void OnPrevRow(IRibbonControl control)
        {
            if (!CheckNavigatorSheet()) return;
            NavigateByIndex(-1);
        }

        public void OnNextRow(IRibbonControl control)
        {
            if (!CheckNavigatorSheet()) return;
            NavigateByIndex(1);
        }

        public void OnShowList(IRibbonControl control)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            if (!CheckNavigatorSheet()) return;

            // Handle Protection FIRST
            Excel.Workbook wb = app.ActiveWorkbook;
            string pwd = GetStoredPassword(wb);
            bool wasProtected = wb.ProtectStructure;

            if (wasProtected && !string.IsNullOrEmpty(pwd))
            {
                try { wb.Unprotect(pwd); } catch { }
            }

            try
            {
                // Just show the existing list - DO NOT REBUILD
                // per user request: "dont rebuilt the list on show list button just we will hide/unhide"

                Excel.Worksheet navSh = (Excel.Worksheet)app.Worksheets[NAV_SHEET];
                navSh.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                navSh.Activate();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("OnShowList error: " + ex.Message);
            }
            finally
            {
                // ALWAYS Reprotect if it was protected
                if (wasProtected && !string.IsNullOrEmpty(pwd))
                {
                    try { wb.Protect(pwd, true, false); } catch { }
                }
            }
        }

        // --- Helper methods for navigation and list management ---
        private bool CheckNavigatorSheet()
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            try
            {
                var sh = app.Worksheets[NAV_SHEET];
                return true;
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("Navigation list is missing. Please click 'Start Focus' again.", "Navigator");
                return false;
            }
        }

        private void HookEvents(Excel.Application app)
        {
            try
            {
                UnhookEvents(app);
                _selectionChangeHandler = new Excel.AppEvents_SheetSelectionChangeEventHandler(OnSelectionChange);
                app.SheetSelectionChange += _selectionChangeHandler;

                // Add Double Click
                _doubleClickHandler = new Excel.AppEvents_SheetBeforeDoubleClickEventHandler(OnBeforeDoubleClick);
                app.SheetBeforeDoubleClick += _doubleClickHandler;

                _beforeCloseHandler = new Excel.AppEvents_WorkbookBeforeCloseEventHandler(OnWorkbookBeforeClose);
                app.WorkbookBeforeClose += _beforeCloseHandler;
            }
            catch { }
        }

        private void UnhookEvents(Excel.Application app)
        {
            try
            {
                if (_selectionChangeHandler != null)
                    app.SheetSelectionChange -= _selectionChangeHandler;
                if (_doubleClickHandler != null)
                    app.SheetBeforeDoubleClick -= _doubleClickHandler; // Hook Remove
                if (_beforeCloseHandler != null)
                    app.WorkbookBeforeClose -= _beforeCloseHandler;
            }
            catch { }
            finally
            {
                _selectionChangeHandler = null;
                _doubleClickHandler = null;
                _beforeCloseHandler = null;
            }
        }

        private void OnSelectionChange(object Sh, Excel.Range Target)
        {
            // Just update status bar, DO NOT JUMP
            Excel.Worksheet sheet = (Excel.Worksheet)Sh;
            if (sheet.Name != NAV_SHEET) return;
            if (Target.Row < DATA_START_ROW) return;
            try
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                app.StatusBar = "Navigator: Double-click row or press GO to jump.";
            }
            catch { }
        }

        private void OnBeforeDoubleClick(object Sh, Excel.Range Target, ref bool Cancel)
        {
            Excel.Worksheet sheet = (Excel.Worksheet)Sh;
            if (sheet.Name != NAV_SHEET) return;

            int r = Target.Row;
            if (r < DATA_START_ROW) return;

            try
            {
                Cancel = true; // Prevent entering edit mode
                Excel.Range cells = (Excel.Range)((Excel._Worksheet)sheet).Cells;
                Excel.Range cell = (Excel.Range)cells[r, 1];
                object idxVal = cell.Value;

                if (idxVal != null && idxVal is double)
                {
                    JumpToIndex(Convert.ToInt32(idxVal));
                }
            }
            catch { }
        }

        private void OnWorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            // If this workbook is protected with our password, unprotect it before closing
            if (Wb.ProtectStructure)
            {
                string pwd = GetStoredPassword(Wb);
                if (!string.IsNullOrEmpty(pwd))
                {
                    try { Wb.Unprotect(pwd); } catch { }
                }
            }
        }

        [ExcelCommand(MenuName = "Navigator", MenuText = "Jump To Index")]
        public static void NavigatorJumpButton_Click()
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet activeSh = (Excel.Worksheet)app.ActiveSheet;
            if (activeSh.Name != NAV_SHEET) return;

            Excel.Range activeCell = app.ActiveCell;
            int r = activeCell.Row;
            if (r < DATA_START_ROW) return;

            try
            {
                Excel.Range cells = (Excel.Range)((Excel._Worksheet)activeSh).Cells;
                Excel.Range cell = (Excel.Range)cells[r, 1];
                object idxVal = cell.Value;

                if (idxVal != null && idxVal is double)
                {
                    JumpToIndex(Convert.ToInt32(idxVal));
                }
            }
            catch { }
        }

        private static void JumpToIndex(int idx)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet navSh = null;
            try
            {
                navSh = (Excel.Worksheet)app.Worksheets[NAV_SHEET];
                ((Excel.Range)navSh.Range["E1"]).Value = idx;

                Excel.Range cells = (Excel.Range)((Excel._Worksheet)navSh).Cells;
                Excel.Range rowInBCell = (Excel.Range)cells[idx + DATA_START_ROW - 1, 2];
                int rowInB = Convert.ToInt32(rowInBCell.Value);

                ApplySingleRowFocus(rowInB, navSh);
                UpdateStatusBar(idx, navSh);

                // Auto-Hide Logic with Protection Handling
                Excel.Workbook wb = app.ActiveWorkbook;
                string pwd = GetStoredPassword(wb);

                if (wb.ProtectStructure && !string.IsNullOrEmpty(pwd))
                {
                    try
                    {
                        wb.Unprotect(pwd);
                        navSh.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                        wb.Protect(pwd, true, false); // Reprotect Structure
                    }
                    catch { }
                }
                else
                {
                    // Fallback if not protected or no password found
                    try { navSh.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden; } catch { }
                }
            }
            catch { }
            finally
            {
                // Double safety: Ensure screen updating is back on
                try { app.ScreenUpdating = true; } catch { }
            }
        }

        private void NavigateByIndex(int direction)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            try
            {
                Excel.Worksheet navSh = (Excel.Worksheet)app.Worksheets[NAV_SHEET];
                object curIdxObj = ((Excel.Range)navSh.Range["E1"]).Value;
                if (curIdxObj == null) return;

                int curIdx = Convert.ToInt32(curIdxObj);

                Excel.Range cells = (Excel.Range)((Excel._Worksheet)navSh).Cells;
                Excel.Range lastCell = (Excel.Range)cells[navSh.Rows.Count, 1];
                int lastIdx = lastCell.End[Excel.XlDirection.xlUp].Row - (DATA_START_ROW - 1);

                int idx = curIdx + direction;
                if (idx < 1) idx = lastIdx;
                if (idx > lastIdx) idx = 1;

                JumpToIndex(idx);
            }
            catch { }
        }

        private static void ApplySingleRowFocus(int targetRow, Excel.Worksheet navSh)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            string dataSheetName = Convert.ToString(((Excel.Range)navSh.Range["F1"]).Value);
            int filterCol = Convert.ToInt32(((Excel.Range)navSh.Range["G1"]).Value);
            Excel.Worksheet dataSh = (Excel.Worksheet)app.Worksheets[dataSheetName];

            Excel.Range navCells = (Excel.Range)((Excel._Worksheet)navSh).Cells;
            Excel.Range lastCellB = (Excel.Range)navCells[navSh.Rows.Count, 2];
            int lastRowList = lastCellB.End[Excel.XlDirection.xlUp].Row;
            if (lastRowList < DATA_START_ROW) return;

            Excel.Range firstRowCell = (Excel.Range)navCells[DATA_START_ROW, 2];
            Excel.Range lastRowCell = (Excel.Range)navCells[lastRowList, 2];

            int firstRow = Convert.ToInt32(firstRowCell.Value);
            int lastRowVal = Convert.ToInt32(lastRowCell.Value);

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                Excel.Range dataRows = (Excel.Range)dataSh.Rows;
                Excel.Range r1 = (Excel.Range)dataRows[firstRow];
                Excel.Range r2 = (Excel.Range)dataRows[lastRowVal];
                Excel.Range rngToHide = dataSh.Range[r1, r2];
                rngToHide.Hidden = true;

                ((Excel.Range)dataRows[targetRow]).Hidden = false;

                dataSh.Activate();
                Excel.Range cells = (Excel.Range)((Excel._Worksheet)dataSh).Cells;
                ((Excel.Range)cells[targetRow, filterCol]).Select();

                UpdateHeaderArrowFromNavSheet(dataSh, navSh);
            }
            catch { }
            finally
            {
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.ScreenUpdating = true;
            }
        }


        private static void UpdateStatusBar(int idx, Excel.Worksheet navSh)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Range cells = (Excel.Range)((Excel._Worksheet)navSh).Cells;
            Excel.Range lastCell = (Excel.Range)cells[navSh.Rows.Count, 1];
            int total = lastCell.End[Excel.XlDirection.xlUp].Row - (DATA_START_ROW - 1);
            app.StatusBar = "Navigator: " + idx + " / " + total;
        }

        private void BuildNavigatorList(Excel.Range rngFilter, Excel.Worksheet dataSheet, int activeCol)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Worksheet navSh = GetOrCreateNavigatorSheet(app);

            navSh.Cells.Clear();
            DeleteExistingShapes(navSh);
            navSh.Activate();
            app.ActiveWindow.FreezePanes = false;

            ((Excel.Range)navSh.Range["E1"]).Value = 1;
            ((Excel.Range)navSh.Range["F1"]).Value = dataSheet.Name;
            ((Excel.Range)navSh.Range["G1"]).Value = activeCol; // Store Active Column
            ((Excel.Range)navSh.Range["H1"]).Value = rngFilter.Row; // Store Header Row

            int hdrRow = DATA_START_ROW - 1;
            Excel.Range navCells = (Excel.Range)((Excel._Worksheet)navSh).Cells;
            Excel.Range hdrRng = (Excel.Range)navCells[hdrRow, 1];
            Excel.Range resizeHdr = hdrRng.Resize[1, 3];
            resizeHdr.Value = new object[,] { { "Index", "Row", "Item" } };
            resizeHdr.Font.Bold = true;
            resizeHdr.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(240, 240, 240));

            Excel.Range dataRng = rngFilter.Offset[1, 0].Resize[rngFilter.Rows.Count - 1, Type.Missing];

            List<int> visibleRows = new List<int>();
            try
            {
                Excel.Range vis = dataRng.SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                foreach (Excel.Range area in vis.Areas)
                {
                    foreach (Excel.Range r in area.Rows)
                    {
                        visibleRows.Add(r.Row);
                    }
                }
            }
            catch { }

            if (visibleRows.Count == 0) return;

            object[,] output = new object[visibleRows.Count, 3];
            // Use ACTIVE COL for data extraction
            int filterCol = activeCol;
            Excel.Range dataCells = (Excel.Range)((Excel._Worksheet)dataSheet).Cells;

            for (int i = 0; i < visibleRows.Count; i++)
            {
                int r = visibleRows[i];
                output[i, 0] = i + 1;
                output[i, 1] = r;
                Excel.Range c = (Excel.Range)dataCells[r, filterCol];
                output[i, 2] = c.Value;
            }

            Excel.Range startCell = (Excel.Range)navCells[DATA_START_ROW, 1];
            Excel.Range writeRange = startCell.Resize[visibleRows.Count, 3];
            writeRange.Value = output;
            ((Excel.Range)navSh.Columns["A:C"]).AutoFit();

            try
            {
                AddGoButton(navSh);
            }
            catch (Exception ex)
            {
                // Just ignore button failure so we don't break the whole flow
                System.Diagnostics.Debug.WriteLine("Button creation failed: " + ex.Message);
            }

            ((Excel.Range)navSh.Range["E1:G1"]).Locked = false;
            // ...
        }

        private void IgnoreThisMethod() { } // Dummy to keep offset valid if needed, but replacing logic below

        private void AddGoButton(Excel.Worksheet sh)
        {
            try
            {
                Excel.Range target = sh.Range["F2:H4"];
                double left = (double)target.Left;
                double top = (double)target.Top;
                double width = (double)target.Width;
                double height = (double)target.Height;

                // Use AddFormControl which creates a standard button
                Excel.Shape btnShape = sh.Shapes.AddFormControl(Excel.XlFormControl.xlButtonControl, (int)left, (int)top, (int)width, (int)height);
                btnShape.Name = "Nav_Go_Button";
                btnShape.OnAction = "NavigatorJumpButton_Click";

                // Set text on the button object itself
                // The Shape contains a ControlFormat or OLEFormat, but for Form Controls, we can typically access properties via OLEFormat.Object or by keeping it simple
                // Actually, for Form Control Buttons, we can usually just set the characters text.
                // But safest is accessing the underlying object dynamically if needed, or simply:
                btnShape.TextFrame.Characters(Type.Missing, Type.Missing).Text = "GO / JUMP";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("AddFormControl failed: " + ex.Message);
            }
        }

        private void DeleteExistingShapes(Excel.Worksheet sh)
        {
            try
            {
                // Delete specific shapes we might have added, or all shapes if safe
                // Since we clear the sheet, we can delete all shapes
                Excel.Shapes shapes = sh.Shapes;
                for (int i = shapes.Count; i >= 1; i--)
                {
                    try { shapes.Item(i).Delete(); } catch { }
                }
            }
            catch { }
        }

        private static void UpdateHeaderArrow(Excel.Worksheet sheet, int column, int headerRowIndex)
        {
            RemoveHeaderArrow(sheet);

            try
            {
                Excel.Range cells = (Excel.Range)((Excel._Worksheet)sheet).Cells;
                Excel.Range headerCell;

                // Use the passed headerRowIndex if valid, otherwise fallback
                if (headerRowIndex > 0)
                {
                    headerCell = (Excel.Range)cells[headerRowIndex, column];
                }
                else if (sheet.AutoFilterMode && sheet.AutoFilter != null)
                {
                    headerCell = (Excel.Range)cells[sheet.AutoFilter.Range.Row, column];
                }
                else
                {
                    // Fallback to row 1 if we really don't know
                    headerCell = (Excel.Range)cells[1, column];
                }

                double cellLeft = (double)headerCell.Left;
                double cellTop = (double)headerCell.Top;
                double cellWidth = (double)headerCell.Width;
                double cellHeight = (double)headerCell.Height;

                // PLACEMENT: INSIDE the cell (Left +, Left + Width -)
                // DIRECTION: Pointing OUTWARD (<- Text ->)

                // Left Arrow -> msoShapeLeftArrow (Points <)
                Excel.Shape arrowL = sheet.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeLeftArrow, (float)(cellLeft + 2), (float)(cellTop + (cellHeight - ARROW_WIDTH_PTS) / 2), ARROW_WIDTH_PTS, ARROW_WIDTH_PTS);
                arrowL.Name = HEADER_ARROW_LEFT_NAME;
                arrowL.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                arrowL.Line.Visible = Office.MsoTriState.msoFalse;

                // Right Arrow -> msoShapeRightArrow (Points >)
                Excel.Shape arrowR = sheet.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRightArrow, (float)(cellLeft + cellWidth - ARROW_WIDTH_PTS - 2), (float)(cellTop + (cellHeight - ARROW_WIDTH_PTS) / 2), ARROW_WIDTH_PTS, ARROW_WIDTH_PTS);
                arrowR.Name = HEADER_ARROW_RIGHT_NAME;
                arrowR.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                arrowR.Line.Visible = Office.MsoTriState.msoFalse;
            }
            catch { }
        }

        private static void UpdateHeaderArrowFromNavSheet(Excel.Worksheet dataSh, Excel.Worksheet navSh)
        {
            try
            {
                object colVal = ((Excel.Range)navSh.Range["G1"]).Value;
                object rowVal = ((Excel.Range)navSh.Range["H1"]).Value; // Read Header Row

                if (colVal != null)
                {
                    int col = Convert.ToInt32(colVal);
                    int row = 0;
                    if (rowVal != null && rowVal is double) row = Convert.ToInt32(rowVal);

                    UpdateHeaderArrow(dataSh, col, row);
                }
            }
            catch { }
        }

        private static void RemoveHeaderArrow(Excel.Worksheet sheet)
        {
            try
            {
                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    if (shape.Name == HEADER_ARROW_LEFT_NAME || shape.Name == HEADER_ARROW_RIGHT_NAME || shape.Name == "Nav_Header_Arrow")
                    {
                        shape.Delete();
                    }
                }
            }
            catch { }
        }

        private Excel.Worksheet GetOrCreateNavigatorSheet(Excel.Application app)
        {
            try
            {
                return (Excel.Worksheet)app.Worksheets[NAV_SHEET];
            }
            catch
            {
                Excel.Worksheet sh = (Excel.Worksheet)app.Worksheets.Add();
                sh.Name = NAV_SHEET;
                return sh;
            }
        }


        private Excel.Range GetFilterRangeOrRegion(Excel.Worksheet sh)
        {
            if (sh.AutoFilterMode && sh.AutoFilter != null)
            {
                return sh.AutoFilter.Range;
            }
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
            if (app.ActiveCell != null)
            {
                return app.ActiveCell.CurrentRegion;
            }
            return null;
        }

        // --- Password management ---
        private void ProtectWorkbookWithRandomPassword(Excel.Workbook wb)
        {
            try
            {
                Random rnd = new Random();
                int pwdNum = rnd.Next(10000000, 100000000);
                _currentPassword = pwdNum.ToString();

                StorePasswordInWorkbook(wb, _currentPassword);

                if (!wb.ProtectStructure)
                {
                    wb.Protect(Password: _currentPassword, Structure: true, Windows: false);
                }
            }
            catch { }
        }

        private void StorePasswordInWorkbook(Excel.Workbook wb, string password)
        {
            try
            {
                // Delete any existing password range
                try { wb.Names.Item(PASSWORD_RANGE_NAME).Delete(); } catch { }

                // Store password in a safe cell (IV1) of the first actual worksheet
                Excel.Worksheet anySheet = null;
                foreach (var s in wb.Worksheets)
                {
                    if (s is Excel.Worksheet)
                    {
                        anySheet = (Excel.Worksheet)s;
                        break;
                    }
                }

                if (anySheet != null)
                {
                    Excel.Range storeCell = (Excel.Range)anySheet.Cells[1, 256];
                    storeCell.Value = password;
                    storeCell.NumberFormat = "@";

                    Excel.Name name = wb.Names.Add(PASSWORD_RANGE_NAME, storeCell);
                    name.Visible = false;

                    // Try to hide column safely, ignore if fails
                    try { ((Excel.Range)anySheet.Columns[256]).Hidden = true; } catch { }
                }
            }
            catch { }
        }

        private static string GetStoredPassword(Excel.Workbook wb)
        {
            try
            {
                Excel.Name name = wb.Names.Item(PASSWORD_RANGE_NAME);
                if (name == null) return null;
                Excel.Range rng = name.RefersToRange;
                return rng?.Value?.ToString();
            }
            catch { return null; }
        }

        private void UnprotectWorkbook(Excel.Workbook wb)
        {
            if (!wb.ProtectStructure) return;

            string pwd = _currentPassword ?? GetStoredPassword(wb);
            if (!string.IsNullOrEmpty(pwd))
            {
                try { wb.Unprotect(pwd); } catch { }
            }
        }

        // --- Static methods for crash recovery (called by AddinLifecycle) ---
        internal static void CheckAndUnprotectWorkbook(Excel.Workbook wb)
        {
            try
            {
                Excel.Name name = null;
                try { name = wb.Names.Item(PASSWORD_RANGE_NAME); } catch { }
                if (name != null)
                {
                    Excel.Range rng = name.RefersToRange;
                    string pwd = rng?.Value?.ToString();
                    if (!string.IsNullOrEmpty(pwd) && wb.ProtectStructure)
                    {
                        wb.Unprotect(pwd);
                    }
                }
            }
            catch { }
        }

        internal static void OnWorkbookOpen(Excel.Workbook Wb)
        {
            CheckAndUnprotectWorkbook(Wb);
        }
    }
}
