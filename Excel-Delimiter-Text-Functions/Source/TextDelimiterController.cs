using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using System.Windows.Forms;

namespace TextDelimiterAddIn
{
    [ComVisible(true)]
    public class TextDelimiterController : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"
              <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon>
                <tabs>
                  <tab id='tabTextDelimiter' label='Text Tools'>
                    <group id='grpText' label='Delimiter Tools'>
                      <button id='btnShowTextTool' label='Text Delimiter Tool' onAction='OnShowForm' size='large' imageMso='TextToTable' />
                      <separator id='sep1' />
                      <button id='btnInsertLeft' label='TextLeft()' onAction='OnInsertTextLeft' size='normal' imageMso='FunctionWizard' />
                      <button id='btnInsertMid' label='TextMid()' onAction='OnInsertTextMid' size='normal' imageMso='FunctionWizard' />
                      <button id='btnInsertRight' label='TextRight()' onAction='OnInsertTextRight' size='normal' imageMso='FunctionWizard' />
                    </group>
                  </tab>
                </tabs>
              </ribbon>
            </customUI>";
        }

        public void OnShowForm(IRibbonControl control)
        {
            TextDelimiterForm form = new TextDelimiterForm();
            form.Show();
        }

        public void OnInsertTextLeft(IRibbonControl control)
        {
            SendKeys.SendWait("=TextLeft(");
        }

        public void OnInsertTextMid(IRibbonControl control)
        {
            SendKeys.SendWait("=TextMid(");
        }

        public void OnInsertTextRight(IRibbonControl control)
        {
            SendKeys.SendWait("=TextRight(");
        }
    }
}
