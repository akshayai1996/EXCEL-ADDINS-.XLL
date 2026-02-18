using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelMergedFilter
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
        }

        public void AutoClose()
        {
        }
    }

    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
            <customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabMergedFilter' label='Merged Filter'>
                    <group id='grpFilter' label='Filter Actions'>
                      <button id='btnApply' label='Apply Merged Filter' size='large' onAction='OnApplyFilter' imageMso='Filter' />
                      <button id='btnClear' label='Clear All Filters' size='large' onAction='OnClearFilter' imageMso='FilterClear' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnApplyFilter(IRibbonControl control)
        {
            FilterEngine.ApplyFilter();
        }

        public void OnClearFilter(IRibbonControl control)
        {
            Application app = (Application)ExcelDnaUtil.Application;
            if (app.ActiveWorkbook != null && app.ActiveSheet != null)
            {
                if (app.ActiveSheet is Worksheet ws)
                {
                    FilterEngine.ClearFilter(ws, true);
                }
            }
        }
    }
}
