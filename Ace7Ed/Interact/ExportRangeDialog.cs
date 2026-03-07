using Ace7Ed;
using Ace7LocalizationFormat.Formats;
using CMN = Ace7LocalizationFormat.Formats.CmnFile;
using System;
using System.Windows.Forms;

namespace Ace7Ed.Interact
{
    public partial class ExportRangeDialog : Form
    {
        public int StartNumber => Convert.ToInt32(StartNumberNumericUpDown.Value);
        public int EndNumber => Convert.ToInt32(EndNumberNumericUpDown.Value);

        public ExportRangeDialog(CMN cmn)
        {
            InitializeComponent();
            BackColor = Theme.WindowColor;
            ForeColor = Theme.WindowTextColor;
            Theme.SetDarkThemeLabel(StartNumberLabel);
            Theme.SetDarkThemeLabel(EndNumberLabel);
            Theme.SetDarkThemeButton(OkButton);
            Theme.SetDarkThemeButton(CancelBtn);

            StartNumberNumericUpDown.Maximum = cmn.MaxStringNumber;
            EndNumberNumericUpDown.Minimum = 1;
            EndNumberNumericUpDown.Maximum = cmn.MaxStringNumber + 1;
            EndNumberNumericUpDown.Value = Math.Min(cmn.MaxStringNumber + 1, EndNumberNumericUpDown.Maximum);
        }

        private void StartNumberNumericUpDown_ValueChanged(object? sender, EventArgs e)
        {
            EndNumberNumericUpDown.Minimum = StartNumberNumericUpDown.Value + 1;
        }

        private void OkButton_Click(object? sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
