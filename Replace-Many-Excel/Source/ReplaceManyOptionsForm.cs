using System;
using System.Windows.Forms;
using System.Drawing;

namespace ReplaceManyAddIn
{
    public class ReplaceManyOptionsForm : Form
    {
        public bool IncludeFormulas { get; private set; }
        public bool CaseSensitive { get; private set; }
        public int Scope { get; private set; } // 1=Range, 2=Sheet, 3=Workbook

        public string WorkbookName { set { lblWorkbook.Text = "Workbook: " + value; } }
        public string SheetName { set { lblSheet.Text = "Sheet: " + value; } }
        public string SelectionAddress { set { lblSelection.Text = "Selection: " + value; } }
        
        private RadioButton rbRange;
        private RadioButton rbSheet;
        private RadioButton rbWorkbook;
        private CheckBox chkFormulas;
        private CheckBox chkCase;
        private Button btnOk;
        private Button btnCancel;
        
        private Label lblWorkbook;
        private Label lblSheet;
        private Label lblSelection;

        public ReplaceManyOptionsForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Replace Many Options";
            this.Size = new Size(400, 450); // Increased Height
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Context Group
            GroupBox grpContext = new GroupBox();
            grpContext.Text = "Current Context";
            grpContext.Location = new Point(10, 10);
            grpContext.Size = new Size(360, 90);

            lblWorkbook = new Label() { Location = new Point(10, 20), AutoSize = true, Text = "Workbook: " };
            lblSheet = new Label() { Location = new Point(10, 40), AutoSize = true, Text = "Sheet: " };
            lblSelection = new Label() { Location = new Point(10, 60), AutoSize = true, Text = "Selection: " };
            
            grpContext.Controls.Add(lblWorkbook);
            grpContext.Controls.Add(lblSheet);
            grpContext.Controls.Add(lblSelection);

            // Scope Group
            GroupBox grpScope = new GroupBox();
            grpScope.Text = "Scope";
            grpScope.Location = new Point(10, 110);
            grpScope.Size = new Size(360, 100);

            rbRange = new RadioButton();
            rbRange.Text = "Specific Range / Selection";
            rbRange.Location = new Point(10, 20);
            rbRange.AutoSize = true;
            rbRange.Checked = true;

            rbSheet = new RadioButton();
            rbSheet.Text = "Active Sheet (Used Range)";
            rbSheet.Location = new Point(10, 45);
            rbSheet.AutoSize = true;

            rbWorkbook = new RadioButton();
            rbWorkbook.Text = "Entire Workbook";
            rbWorkbook.Location = new Point(10, 70);
            rbWorkbook.AutoSize = true;

            grpScope.Controls.Add(rbRange);
            grpScope.Controls.Add(rbSheet);
            grpScope.Controls.Add(rbWorkbook);

            chkFormulas = new CheckBox();
            chkFormulas.Text = "Include Text Formulas (Convert to Values)";
            chkFormulas.Location = new Point(15, 220);
            chkFormulas.AutoSize = true;

            chkCase = new CheckBox();
            chkCase.Text = "Case Sensitive Matching";
            chkCase.Location = new Point(15, 245);
            chkCase.AutoSize = true;

            btnOk = new Button();
            btnOk.Text = "OK";
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Location = new Point(110, 350); // Moved down
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Location = new Point(190, 350); // Moved down

            this.Controls.Add(grpContext);
            this.Controls.Add(grpScope);
            this.Controls.Add(chkFormulas);
            this.Controls.Add(chkCase);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
            
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            IncludeFormulas = chkFormulas.Checked;
            CaseSensitive = chkCase.Checked;
            if (rbRange.Checked) Scope = 1;
            else if (rbSheet.Checked) Scope = 2;
            else Scope = 3;
        }
    }
}
