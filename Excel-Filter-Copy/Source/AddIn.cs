using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Runtime.InteropServices;

namespace ExcelFilterCopy
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
                  <tab id='tabMergedFilter' label='Copy Filtered'>
                    <group id='grpCopy' label='Copy Actions'>
                      <button id='btnCopy' label='Copy Visible &amp; Merged' size='large' onAction='OnCopyFiltered' imageMso='Copy' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnCopyFiltered(IRibbonControl control)
        {
            CopyEngine.CopyFilteredWithFormat();
        }
    }
}
