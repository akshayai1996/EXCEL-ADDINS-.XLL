using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace NavigatorArrowsAddIn
{
    public class AddinLifecycle : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Queue on the main Excel thread
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                    if (app != null)
                    {
                        // Hook future workbook opens
                        app.WorkbookOpen += NavigatorController.OnWorkbookOpen;

                        // Process already open workbooks (in case of crash recovery)
                        foreach (Excel.Workbook wb in app.Workbooks)
                        {
                            NavigatorController.CheckAndUnprotectWorkbook(wb);
                        }
                    }
                }
                catch { }
            });
        }

        public void AutoClose()
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                    if (app != null)
                    {
                        app.WorkbookOpen -= NavigatorController.OnWorkbookOpen;
                    }
                }
                catch { }
            });
        }
    }
}
