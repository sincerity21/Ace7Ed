namespace Ace7Ed.Interact
{
    partial class ExportRangeDialog
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            StartNumberLabel = new Label();
            EndNumberLabel = new Label();
            StartNumberNumericUpDown = new NumericUpDown();
            EndNumberNumericUpDown = new NumericUpDown();
            OkButton = new Button();
            CancelBtn = new Button();
            ((System.ComponentModel.ISupportInitialize)StartNumberNumericUpDown).BeginInit();
            ((System.ComponentModel.ISupportInitialize)EndNumberNumericUpDown).BeginInit();
            SuspendLayout();
            //
            // StartNumberLabel
            //
            StartNumberLabel.AutoSize = true;
            StartNumberLabel.Location = new Point(12, 15);
            StartNumberLabel.Name = "StartNumberLabel";
            StartNumberLabel.Size = new Size(78, 15);
            StartNumberLabel.TabIndex = 0;
            StartNumberLabel.Text = "Start Number";
            //
            // EndNumberLabel
            //
            EndNumberLabel.AutoSize = true;
            EndNumberLabel.Location = new Point(12, 44);
            EndNumberLabel.Name = "EndNumberLabel";
            EndNumberLabel.Size = new Size(74, 15);
            EndNumberLabel.TabIndex = 1;
            EndNumberLabel.Text = "End Number";
            //
            // StartNumberNumericUpDown
            //
            StartNumberNumericUpDown.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            StartNumberNumericUpDown.Location = new Point(120, 13);
            StartNumberNumericUpDown.Name = "StartNumberNumericUpDown";
            StartNumberNumericUpDown.Size = new Size(200, 23);
            StartNumberNumericUpDown.TabIndex = 2;
            StartNumberNumericUpDown.ValueChanged += StartNumberNumericUpDown_ValueChanged;
            //
            // EndNumberNumericUpDown
            //
            EndNumberNumericUpDown.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            EndNumberNumericUpDown.Location = new Point(120, 42);
            EndNumberNumericUpDown.Name = "EndNumberNumericUpDown";
            EndNumberNumericUpDown.Size = new Size(200, 23);
            EndNumberNumericUpDown.TabIndex = 3;
            //
            // OkButton
            //
            OkButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            OkButton.Location = new Point(164, 82);
            OkButton.Name = "OkButton";
            OkButton.Size = new Size(75, 23);
            OkButton.TabIndex = 4;
            OkButton.Text = "OK";
            OkButton.UseVisualStyleBackColor = true;
            OkButton.Click += OkButton_Click;
            //
            // CancelBtn
            //
            CancelBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            CancelBtn.DialogResult = DialogResult.Cancel;
            CancelBtn.Location = new Point(245, 82);
            CancelBtn.Name = "CancelBtn";
            CancelBtn.Size = new Size(75, 23);
            CancelBtn.TabIndex = 5;
            CancelBtn.Text = "Cancel";
            CancelBtn.UseVisualStyleBackColor = true;
            //
            // ExportRangeDialog
            //
            AcceptButton = OkButton;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = CancelBtn;
            ClientSize = new Size(332, 117);
            Controls.Add(CancelBtn);
            Controls.Add(OkButton);
            Controls.Add(EndNumberNumericUpDown);
            Controls.Add(StartNumberNumericUpDown);
            Controls.Add(EndNumberLabel);
            Controls.Add(StartNumberLabel);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ExportRangeDialog";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Export range";
            ((System.ComponentModel.ISupportInitialize)StartNumberNumericUpDown).EndInit();
            ((System.ComponentModel.ISupportInitialize)EndNumberNumericUpDown).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label StartNumberLabel;
        private Label EndNumberLabel;
        private NumericUpDown StartNumberNumericUpDown;
        private NumericUpDown EndNumberNumericUpDown;
        private Button OkButton;
        private Button CancelBtn;
    }
}
