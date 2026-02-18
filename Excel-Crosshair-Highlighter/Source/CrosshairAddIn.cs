using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelCrosshairAddIn
{
    public class CrosshairAddIn : IExcelAddIn
    {
        public static CrosshairController Controller { get; private set; }

        public void AutoOpen()
        {
            try
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Controller = new CrosshairController();
                Controller.Initialize(app);
            }
            catch (Exception ex)
            {
                // Simple error logging or ignore
                System.Diagnostics.Debug.WriteLine("Error in CrosshairAddIn AutoOpen: " + ex.Message);
            }
        }

        public void AutoClose()
        {
            if (Controller != null)
            {
                Controller.Terminate();
                Controller = null;
            }
        }
    }
}
