using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace ExcelCrosshairAddIn
{
    [ComVisible(true)]
    public class RibbonController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
              <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabCrosshair' label='Crosshair'>
                    <group id='grpControls' label='Controls'>
                      <button id='btnToggle' label='Toggle Highlighter' onAction='OnToggle' size='large' imageMso='SelectionPane' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnToggle(IRibbonControl control)
        {
            if (ExcelCrosshairAddIn.CrosshairAddIn.Controller != null)
            {
                ExcelCrosshairAddIn.CrosshairAddIn.Controller.Toggle();
            }
        }
    }
}
