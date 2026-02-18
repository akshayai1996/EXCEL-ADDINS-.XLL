using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;

namespace ExcelMergedFilter
{
    public partial class FilterForm : Form
    {
        private List<string> mKeys;

        public FilterForm()
        {
            InitializeComponent();
        }

        public void InitWithKeys(IEnumerable<string> keys)
        {
            mKeys = keys.ToList();
            mKeys.Sort();

            cboFilter.Items.Clear();
            foreach (var key in mKeys)
            {
                cboFilter.Items.Add(key);
            }
            
            ApplyFilter("");
        }

        private void ApplyFilter(string term)
        {
            string needle = term.Trim();
            lstValues.Items.Clear();

            if (mKeys == null) return;

            lstValues.BeginUpdate(); // Performance
            foreach (var key in mKeys)
            {
                if (string.IsNullOrEmpty(needle) || key.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    lstValues.Items.Add(key);
                }
            }
            lstValues.EndUpdate();
        }

        private void cboFilter_TextChanged(object sender, EventArgs e)
        {
            ApplyFilter(cboFilter.Text);
        }

        public void SetSelections(IEnumerable<string> arrToSelect)
        {
            if (arrToSelect == null) return;
            
            HashSet<string> toSelect = new HashSet<string>(arrToSelect, StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < lstValues.Items.Count; i++)
            {
                if (toSelect.Contains(lstValues.Items[i].ToString()))
                {
                    lstValues.SetSelected(i, true);
                }
            }
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            Dictionary<string, bool> sel = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            
            foreach (var item in lstValues.SelectedItems)
            {
                sel[item.ToString()] = true;
            }
            
            StateManager.SelectedValues = sel;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cmdCancel_Click(object sender, EventArgs e)
        {
            StateManager.SelectedValues = null;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
