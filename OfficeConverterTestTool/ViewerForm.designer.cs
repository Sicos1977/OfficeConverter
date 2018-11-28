namespace OfficeConverterTestTool
{
    partial class ViewerForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SelectButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.OutputTextBox = new System.Windows.Forms.TextBox();
            this.LibreOfficeCheckBox = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // SelectButton
            // 
            this.SelectButton.Location = new System.Drawing.Point(12, 11);
            this.SelectButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.SelectButton.Name = "SelectButton";
            this.SelectButton.Size = new System.Drawing.Size(92, 23);
            this.SelectButton.TabIndex = 6;
            this.SelectButton.Text = "Select file";
            this.SelectButton.UseVisualStyleBackColor = true;
            this.SelectButton.Click += new System.EventHandler(this.SelectButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 44);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(42, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "Output:";
            // 
            // OutputTextBox
            // 
            this.OutputTextBox.Location = new System.Drawing.Point(10, 63);
            this.OutputTextBox.Multiline = true;
            this.OutputTextBox.Name = "OutputTextBox";
            this.OutputTextBox.Size = new System.Drawing.Size(451, 347);
            this.OutputTextBox.TabIndex = 9;
            // 
            // LibreOfficeCheckBox
            // 
            this.LibreOfficeCheckBox.AutoSize = true;
            this.LibreOfficeCheckBox.Checked = global::OfficeConverterTestTool.Properties.Settings.Default.UseLibreOffice;
            this.LibreOfficeCheckBox.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::OfficeConverterTestTool.Properties.Settings.Default, "UseLibreOffice", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.LibreOfficeCheckBox.Location = new System.Drawing.Point(124, 16);
            this.LibreOfficeCheckBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.LibreOfficeCheckBox.Name = "LibreOfficeCheckBox";
            this.LibreOfficeCheckBox.Size = new System.Drawing.Size(102, 17);
            this.LibreOfficeCheckBox.TabIndex = 10;
            this.LibreOfficeCheckBox.Text = "Use Libre Office";
            this.LibreOfficeCheckBox.UseVisualStyleBackColor = true;
            // 
            // ViewerForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(473, 422);
            this.Controls.Add(this.LibreOfficeCheckBox);
            this.Controls.Add(this.OutputTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.SelectButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.MaximizeBox = false;
            this.Name = "ViewerForm";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "OfficeConverter test tool v1.2";
            this.Load += new System.EventHandler(this.ViewerForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SelectButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox OutputTextBox;
        private System.Windows.Forms.CheckBox LibreOfficeCheckBox;
    }
}

