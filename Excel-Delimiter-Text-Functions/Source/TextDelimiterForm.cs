using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace TextDelimiterAddIn
{
    public class TextDelimiterForm : Form
    {
        private RadioButton rbLeft;
        private RadioButton rbRight;
        private RadioButton rbMid;
        private TextBox txtSource;
        private TextBox txtDest;
        private TextBox txtDelim;
        private NumericUpDown numN1;
        private NumericUpDown numN2;
        private Label lblN1;
        private Label lblN2;
        private Button btnRun;
        private Button btnClose;

        public TextDelimiterForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Text Delimiter Tool";
            this.Size = new Size(400, 350);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            // Operation Group
            GroupBox grpOp = new GroupBox();
            grpOp.Text = "Operation";
            grpOp.Bounds = new Rectangle(10, 10, 360, 60);
            
            rbLeft = new RadioButton() { Text = "Get Left", Location = new Point(20, 20), AutoSize = true, Checked = true };
            rbRight = new RadioButton() { Text = "Get Right", Location = new Point(120, 20), AutoSize = true };
            rbMid = new RadioButton() { Text = "Get Mid", Location = new Point(220, 20), AutoSize = true };
            
            rbLeft.CheckedChanged += OpChanged;
            rbRight.CheckedChanged += OpChanged;
            rbMid.CheckedChanged += OpChanged;

            grpOp.Controls.Add(rbLeft);
            grpOp.Controls.Add(rbRight);
            grpOp.Controls.Add(rbMid);
            this.Controls.Add(grpOp);

            // Selection Inputs
            Label lblSource = new Label() { Text = "Source Range:", Location = new Point(20, 90), AutoSize = true };
            txtSource = new TextBox() { Location = new Point(120, 87), Width = 150 };
            Button btnSelSource = new Button() { Text = "Select", Location = new Point(280, 85), Width = 60 };
            btnSelSource.Click += (s, e) => txtSource.Text = PromptRange("Select Source Range");

            Label lblDest = new Label() { Text = "Dest Cell:", Location = new Point(20, 120), AutoSize = true };
            txtDest = new TextBox() { Location = new Point(120, 117), Width = 150 };
            Button btnSelDest = new Button() { Text = "Select", Location = new Point(280, 115), Width = 60 };
            btnSelDest.Click += (s, e) => txtDest.Text = PromptRange("Select Destination Cell");

            this.Controls.AddRange(new Control[] { lblSource, txtSource, btnSelSource, lblDest, txtDest, btnSelDest });

            // Parameters
            Label lblDelim = new Label() { Text = "Delimiter:", Location = new Point(20, 160), AutoSize = true };
            txtDelim = new TextBox() { Text = " ", Location = new Point(120, 157), Width = 50 };

            lblN1 = new Label() { Text = "Occurrence (N):", Location = new Point(20, 190), AutoSize = true };
            numN1 = new NumericUpDown() { Location = new Point(120, 187), Width = 50, Minimum = 1, Maximum = 100, Value = 1 };

            lblN2 = new Label() { Text = "End Occ (N2):", Location = new Point(200, 190), AutoSize = true, Visible = false };
            numN2 = new NumericUpDown() { Location = new Point(280, 187), Width = 50, Minimum = 1, Maximum = 100, Value = 2, Visible = false };

            this.Controls.AddRange(new Control[] { lblDelim, txtDelim, lblN1, numN1, lblN2, numN2 });

            // Buttons
            btnRun = new Button() { Text = "Run", Location = new Point(100, 240), Width = 80, Height = 30 };
            btnRun.Click += BtnRun_Click;

            btnClose = new Button() { Text = "Close", Location = new Point(200, 240), Width = 80, Height = 30 };
            btnClose.Click += (s, e) => this.Close();

            this.Controls.Add(btnRun);
            this.Controls.Add(btnClose);
        }

        private void OpChanged(object sender, EventArgs e)
        {
            if (rbMid.Checked)
            {
                lblN1.Text = "Start Occ (N1):";
                lblN2.Visible = true;
                numN2.Visible = true;
            }
            else
            {
                lblN1.Text = "Occurrence (N):";
                lblN2.Visible = false;
                numN2.Visible = false;
            }
        }

        private string PromptRange(string title)
        {
            this.Hide();
            try
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                object result = app.InputBox("Select range:", title, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                if (result is Excel.Range rng)
                {
                    return rng.get_Address(false, false, Excel.XlReferenceStyle.xlA1, true, false);
                }
            }
            catch { }
            finally
            {
                this.Show();
            }
            return null;
        }

        private void BtnRun_Click(object sender, EventArgs e)
        {
            string srcAddr = txtSource.Text;
            string destAddr = txtDest.Text;
            string delim = txtDelim.Text;
            int n1 = (int)numN1.Value;
            int n2 = (int)numN2.Value;

            if (string.IsNullOrEmpty(srcAddr) || string.IsNullOrEmpty(destAddr)) return;

            try
            {
                Excel.Application app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Range srcRange = app.Range[srcAddr];
                Excel.Range destRef = app.Range[destAddr]; // Destination start cell

                // For relative placement, we need to know the top-left of source
                int startRow = srcRange.Row;
                int startCol = srcRange.Column;
                
                // Working sheet (assume source sheet is active or specified)
                // If External reference, InputBox returns [Book]Sheet!Range. app.Range handles it.
                // However, writing back is safer if we use the object we got.
                
                // Read values
                object[,] values;
                int rows, cols;

                if (srcRange.Count == 1)
                {
                    rows = 1;
                    cols = 1;
                    values = new object[1, 1]; // 0-based
                    values[0, 0] = srcRange.Value2;
                }
                else
                {
                    object val = srcRange.Value2;
                    if (val is object[,] arr)
                    {
                        values = arr; // 1-based from Excel
                        rows = values.GetLength(0);
                        cols = values.GetLength(1);
                    }
                    else
                    {
                        // Fallback
                        rows = 1;
                        cols = 1;
                         values = new object[1, 1];
                         values[0, 0] = val;
                    }
                }

                object[,] results = new object[rows, cols]; // 0-based result array

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        // Handle 1-based vs 0-based access for 'values'
                        object cellVal;
                        if (values.GetLowerBound(0) == 1)
                        {
                            // It's the Excel array (1-based)
                            cellVal = values[r + 1, c + 1];
                        }
                        else
                        {
                            // It's our manual array (0-based)
                            cellVal = values[r, c];
                        }

                        string txt = Convert.ToString(cellVal);
                        string res = "";

                        if (rbLeft.Checked)
                        {
                            string delimN = string.IsNullOrEmpty(delim) ? " " : delim;
                             res = TextDelimiterFunctions.TextLeft(txt, delimN, n1);
                        }
                        else if (rbRight.Checked)
                        {
                             string delimN = string.IsNullOrEmpty(delim) ? " " : delim;
                             res = TextDelimiterFunctions.TextRight(txt, delimN, n1);
                        }
                        else if (rbMid.Checked)
                        {
                             string delimN = string.IsNullOrEmpty(delim) ? " " : delim;
                             res = TextDelimiterFunctions.TextMid(txt, delimN, n1, n2);
                        }

                        results[r, c] = res;
                    }
                }

                // Write back
                // Resize property expects relative size. 
                // We need to construct the range object correctly.
                // destRef is the top-left cell.
                Excel.Range destRange = destRef.Resize[rows, cols];
                destRange.Value2 = results;

                MessageBox.Show("Done!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message + "\nStack: " + ex.StackTrace);
            }
        }
    }
}
