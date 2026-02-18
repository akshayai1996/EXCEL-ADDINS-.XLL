using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReplaceManyAddIn
{
    [ComVisible(true)]
    public class ReplaceManyController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabReplaceMany' label='Replace Many'>
                    <group id='grpReplace' label='Bulk Replace'>
                      <button id='btnReplacePopup' label='Replace Many Tool' onAction='OnShowTool' size='large' imageMso='ReplaceDialog' />
                      <separator id='sep1' />
                      <button id='btnInsertReplaceMany' label='REPLACE_MANY()' onAction='OnInsertReplaceMany' size='large' imageMso='FunctionWizard' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnInsertReplaceMany(IRibbonControl control)
        {
            SendKeys.SendWait("=REPLACE_MANY(");
        }

        public void OnShowTool(IRibbonControl control)
        {
            Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;

            // 1. Prompt for Map Range
            Excel.Range mapRange = PromptForRange(app, "Select the 2-column mapping (From, To). Example: Map!A:B", "REPLACE_MANY - Mapping");
            if (mapRange == null) return;
            
            // Validate map range
            if (mapRange.Columns.Count < 2)
            {
                MessageBox.Show("Please select a 2-column map (From, To)", "REPLACE_MANY", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 2. Show Options Form
            ReplaceManyOptionsForm form = new ReplaceManyOptionsForm();
            
            try
            {
                if (app.ActiveWorkbook != null) form.WorkbookName = app.ActiveWorkbook.Name;
                if (app.ActiveSheet is Excel.Worksheet sh) form.SheetName = sh.Name;
                if (app.Selection is Excel.Range sel) form.SelectionAddress = sel.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            }
            catch { }

            if (form.ShowDialog() != DialogResult.OK) return;

            int scope = form.Scope;
            bool includeFormulas = form.IncludeFormulas;
            bool caseSensitive = form.CaseSensitive;

            // 3. Build Dictionary
            // Need to read values from Map Range
            object[,] mapVals = ReadRangeValues(mapRange);
            Dictionary<string, string> dict = ReplaceManyFunctions.GetDictionary(mapVals, caseSensitive); // Need to expose helper
            
            if (dict.Count == 0)
            {
                 MessageBox.Show("No valid mapping keys found.", "REPLACE_MANY", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                 return;
            }

            string delims = ReplaceManyFunctions.GetDefaultDelims();

            // 4. Process based on Scope
            Excel.Range targetRng = null;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                if (scope == 1) // Specific Range
                {
                    targetRng = PromptForRange(app, "Select the target range to process.", "REPLACE_MANY - Target Range", Type.Missing); 
                    // Type.Missing or default to selection?
                    // PromptForRange handles InputBox.
                    if (targetRng == null)
                    {
                         // If null, try selection?
                         if (app.Selection is Excel.Range sel) targetRng = sel;
                         else return;
                    }
                    ApplyToRange(targetRng, dict, delims, includeFormulas);
                }
                else if (scope == 2) // Active Sheet
                {
                    Excel.Worksheet activeSh = (Excel.Worksheet)app.ActiveSheet;
                    targetRng = activeSh.UsedRange;
                    ApplyToRange(targetRng, dict, delims, includeFormulas);
                }
                else if (scope == 3) // Workbook
                {
                    foreach (Excel.Worksheet sh in app.Worksheets)
                    {
                        app.StatusBar = "REPLACE_MANY: " + sh.Name;
                        if (sh.UsedRange.Cells.Count > 0) // Approximation
                        {
                            ApplyToRange(sh.UsedRange, dict, delims, includeFormulas);
                        }
                    }
                }
                
                MessageBox.Show("REPLACE_MANY completed.", "REPLACE_MANY", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.EnableEvents = true;
                app.StatusBar = false;
            }
        }

        private Excel.Range PromptForRange(Excel.Application app, string prompt, string title, object defaultRange = null)
        {
            try
            {
                object selection = defaultRange;
                if (selection == null || selection == Type.Missing)
                {
                     if (app.Selection is Excel.Range sel) selection = sel.Address;
                }
                
                object rng = app.InputBox(prompt, title, selection, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (rng is Excel.Range r) return r;
            }
            catch
            {
                // Cancelled
            }
            return null;
        }

        private object[,] ReadRangeValues(Excel.Range rng)
        {
             if (rng == null) return new object[0, 0];
             
             // Optimization: Intersect with UsedRange to avoid reading millions of empty rows if user selected A:B
             try
             {
                 Excel.Worksheet ws = (Excel.Worksheet)rng.Worksheet;
                 Excel.Range used = ws.UsedRange;
                 Excel.Range intersected = ((Excel.Application)ExcelDnaUtil.Application).Intersect(rng, used);
                 if (intersected != null) rng = intersected;
             }
             catch { }

             object val = rng.Value2;
             if (val is object[,] arr) return arr;
             
             // Single cell case: create a 1x2 array so BuildDictionary logic doesn't break
             // even though a 1-row mapping is unlikely, it's safer.
             object[,] res = new object[1, 2];
             // Detect if it's 1-based or 0-based result (Interop is 1-based)
             // But here we are creating it manually, so we can pick. 
             // BuildDictionary uses GetLowerBound, so it's fine.
             res[0, 0] = val;
             res[0, 1] = null;
             return res;
        }

        private void ApplyToRange(Excel.Range rng, Dictionary<string, string> dict, string delims, bool includeFormulas)
        {
            if (rng == null) return;

            // Fix: Process EXACTLY what is selected. Do NOT restrict to UsedRange, 
            // as UsedRange might be stale or not include new data.
            
            // Safety check for massive selections (e.g. entire column A:A)
            // Limit to reasonable cell count to prevent freeze
            // 500,000 cells is plenty (e.g. 50 cols x 10,000 rows)
            
            Excel.Range processRng = rng;
            
            // Only use Intersect if row count is massive (entire sheet selection)
            if (Convert.ToInt64(rng.CountLarge) > 1000000)
            {
                try
                {
                   Excel.Worksheet ws = (Excel.Worksheet)rng.Worksheet;
                   processRng = ((Excel.Application)ExcelDnaUtil.Application).Intersect(rng, ws.UsedRange);
                   if (processRng == null) return;
                }
                catch { }
            }

            foreach (Excel.Range area in processRng.Areas)
            {
                // Reading Value2 is faster and doesn't format
                object val = area.Value2;
                
                if (val == null) continue;

                if (area.Cells.Count == 1)
                {
                    // Single cell
                    string resStr = ReplaceManyFunctions.ReplaceInCell(val, dict, delims)?.ToString();
                    area.Value2 = resStr;
                }
                else if (val is object[,] arr)
                {
                    int rows = arr.GetLength(0);
                    int cols = arr.GetLength(1);
                    bool changed = false;

                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            object cellVal = arr[r, c];
                            if (cellVal == null) continue;
                            
                            object newVal = ReplaceManyFunctions.ReplaceInCell(cellVal, dict, delims);
                            arr[r, c] = newVal;
                            changed = true;
                        }
                    }
                    if (changed) area.Value2 = arr;
                }
            }
        }
    }
}
