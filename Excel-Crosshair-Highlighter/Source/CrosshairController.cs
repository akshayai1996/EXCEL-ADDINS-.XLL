using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCrosshairAddIn
{
    [ComVisible(true)]
    public class CrosshairController
    {
        private Excel.Application _app;
        private bool _enabled = false;
        private System.Windows.Forms.Timer _timer;
        private string _lastVisibleRangeAddress;
        private const string ROW_SHAPE_NAME = "XH_Row";
        private const string COL_SHAPE_NAME = "XH_Col";

        public void Initialize(Excel.Application app)
        {
            _app = app;
            _app.SheetSelectionChange += OnSheetSelectionChange;
            
            // Timer for checking scrolling
            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 200; // 200ms check
            _timer.Tick += OnTimerTick;
            _timer.Start();

            _enabled = true; // Auto-start
        }

        public void Terminate()
        {
            if (_timer != null)
            {
                _timer.Stop();
                _timer.Dispose();
                _timer = null;
            }

            if (_app != null)
            {
                _app.SheetSelectionChange -= OnSheetSelectionChange;
                DeleteShapes(_app.ActiveSheet as Excel.Worksheet);
                _app = null;
            }
        }

        public void Toggle()
        {
            _enabled = !_enabled;
            if (_enabled)
            {
                UpdateCrosshair();
                _timer.Start();
            }
            else
            {
                DeleteShapes(_app.ActiveSheet as Excel.Worksheet);
                _timer.Stop();
            }
        }

        private void OnSheetSelectionChange(object Sh, Excel.Range Target)
        {
            if (_enabled)
            {
                UpdateCrosshair();
            }
        }

        private void OnTimerTick(object sender, EventArgs e)
        {
            if (!_enabled || _app == null) return;

            try
            {
                Excel.Window window = _app.ActiveWindow;
                if (window == null) return;
                
                Excel.Range visibleRange = window.VisibleRange;
                if (visibleRange == null) return;

                string currentAddress = visibleRange.Address;
                if (currentAddress != _lastVisibleRangeAddress)
                {
                    _lastVisibleRangeAddress = currentAddress;
                    UpdateCrosshair();
                }
            }
            catch
            {
                // Ignore errors (e.g. no active window)
            }
        }

        private void UpdateCrosshair()
        {
            try
            {
                if (_app == null || _app.ActiveSheet == null) return;
                Excel.Worksheet ws = _app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;
                
                Excel.Range activeCell = _app.ActiveCell;
                if (activeCell == null) return;
                
                Excel.Window window = _app.ActiveWindow;
                if (window == null) return;

                Excel.Range visibleRange = window.VisibleRange;

                CreateOrMoveShapes(ws, activeCell, visibleRange);
            }
            catch
            {
                // Silent fail
            }
        }

        private void CreateOrMoveShapes(Excel.Worksheet ws, Excel.Range activeCell, Excel.Range visibleRange)
        {
            // Consistent shape properties
            int color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(248, 248, 208)); // RGB(248, 248, 208)
            float transparency = 0.65f; // 0.35 in VBA is opacity or transparency? VBA Transparency 0.35 means 35% transparent. C# usually 0-1.

            // Row Shape
            Excel.Shape rowShape = GetShape(ws, ROW_SHAPE_NAME);
            if (rowShape == null)
            {
                rowShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                    (float)(double)visibleRange.Left, (float)(double)activeCell.Top, (float)(double)visibleRange.Width, (float)(double)activeCell.Height);
                rowShape.Name = ROW_SHAPE_NAME;
                rowShape.Fill.ForeColor.RGB = color;
                rowShape.Fill.Transparency = transparency;
                rowShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                rowShape.Placement = Excel.XlPlacement.xlFreeFloating;
                rowShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
            }
            else
            {
                 rowShape.Top = (float)(double)activeCell.Top;
                 rowShape.Left = (float)(double)visibleRange.Left;
                 rowShape.Width = (float)(double)visibleRange.Width;
                 rowShape.Height = (float)(double)activeCell.Height;
                 // Ensure it's behind
                 // rowShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack); // Can be slow if called every time
            }

            // Column Shape (similar logic)
            Excel.Shape colShape = GetShape(ws, COL_SHAPE_NAME);
             if (colShape == null)
            {
                colShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 
                    (float)(double)activeCell.Left, (float)(double)visibleRange.Top, (float)(double)activeCell.Width, (float)(double)visibleRange.Height);
                colShape.Name = COL_SHAPE_NAME;
                colShape.Fill.ForeColor.RGB = color;
                colShape.Fill.Transparency = transparency;
                colShape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                colShape.Placement = Excel.XlPlacement.xlFreeFloating;
                 colShape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack);
            }
            else
            {
                 colShape.Top = (float)(double)visibleRange.Top;
                 colShape.Left = (float)(double)activeCell.Left;
                 colShape.Width = (float)(double)activeCell.Width;
                 colShape.Height = (float)(double)visibleRange.Height;
            }
        }

        private Excel.Shape GetShape(Excel.Worksheet ws, string name)
        {
            try
            {
                return ws.Shapes.Item(name);
            }
            catch
            {
                return null;
            }
        }

        private void DeleteShapes(Excel.Worksheet ws)
        {
            if (ws == null) return;
            try { ws.Shapes.Item(ROW_SHAPE_NAME).Delete(); } catch { }
            try { ws.Shapes.Item(COL_SHAPE_NAME).Delete(); } catch { }
        }
    }
}
