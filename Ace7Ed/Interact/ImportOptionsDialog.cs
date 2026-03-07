using Ace7Ed;
using System;
using System.Windows.Forms;

namespace Ace7Ed.Interact
{
    public partial class ImportOptionsDialog : Form
    {
        public bool OverwriteExistingStrings => OverwriteExistingStringsCheckBox.Checked;

        public ImportOptionsDialog()
        {
            InitializeComponent();
            BackColor = Theme.WindowColor;
            ForeColor = Theme.WindowTextColor;
            Theme.SetDarkThemeCheckBox(OverwriteExistingStringsCheckBox);
            Theme.SetDarkThemeButton(OkButton);
            Theme.SetDarkThemeButton(CancelBtn);
        }

        private void OkButton_Click(object? sender, System.EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
