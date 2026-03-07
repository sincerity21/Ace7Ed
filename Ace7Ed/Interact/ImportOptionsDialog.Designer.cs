namespace Ace7Ed.Interact
{
    partial class ImportOptionsDialog
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
            OverwriteExistingStringsCheckBox = new CheckBox();
            OkButton = new Button();
            CancelBtn = new Button();
            SuspendLayout();
            //
            // OverwriteExistingStringsCheckBox
            //
            OverwriteExistingStringsCheckBox.AutoSize = true;
            OverwriteExistingStringsCheckBox.FlatStyle = FlatStyle.Flat;
            OverwriteExistingStringsCheckBox.Location = new Point(12, 12);
            OverwriteExistingStringsCheckBox.Name = "OverwriteExistingStringsCheckBox";
            OverwriteExistingStringsCheckBox.Size = new Size(180, 19);
            OverwriteExistingStringsCheckBox.TabIndex = 0;
            OverwriteExistingStringsCheckBox.Text = "Overwrite existing strings";
            OverwriteExistingStringsCheckBox.UseVisualStyleBackColor = true;
            //
            // OkButton
            //
            OkButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            OkButton.Location = new Point(124, 48);
            OkButton.Name = "OkButton";
            OkButton.Size = new Size(75, 23);
            OkButton.TabIndex = 1;
            OkButton.Text = "OK";
            OkButton.UseVisualStyleBackColor = true;
            OkButton.Click += OkButton_Click;
            //
            // CancelBtn
            //
            CancelBtn.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            CancelBtn.DialogResult = DialogResult.Cancel;
            CancelBtn.Location = new Point(205, 48);
            CancelBtn.Name = "CancelBtn";
            CancelBtn.Size = new Size(75, 23);
            CancelBtn.TabIndex = 2;
            CancelBtn.Text = "Cancel";
            CancelBtn.UseVisualStyleBackColor = true;
            //
            // ImportOptionsDialog
            //
            AcceptButton = OkButton;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            CancelButton = CancelBtn;
            ClientSize = new Size(292, 83);
            Controls.Add(CancelBtn);
            Controls.Add(OkButton);
            Controls.Add(OverwriteExistingStringsCheckBox);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ImportOptionsDialog";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Import options";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private CheckBox OverwriteExistingStringsCheckBox;
        private Button OkButton;
        private Button CancelBtn;
    }
}
