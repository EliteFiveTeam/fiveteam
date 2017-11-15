namespace RPD
{
    partial class FormExcel
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
            this.bt_selct_excel = new System.Windows.Forms.Button();
            this.PB_Excel = new System.Windows.Forms.ProgressBar();
            this.RTB_ExcelLog = new System.Windows.Forms.RichTextBox();
            this.openFileExcel = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // bt_selct_excel
            // 
            this.bt_selct_excel.Location = new System.Drawing.Point(514, 379);
            this.bt_selct_excel.Name = "bt_selct_excel";
            this.bt_selct_excel.Size = new System.Drawing.Size(110, 50);
            this.bt_selct_excel.TabIndex = 0;
            this.bt_selct_excel.Text = "Выберите \"Учебный план\"";
            this.bt_selct_excel.UseVisualStyleBackColor = true;
            this.bt_selct_excel.Click += new System.EventHandler(this.bt_selct_excel_Click);
            // 
            // PB_Excel
            // 
            this.PB_Excel.Location = new System.Drawing.Point(24, 379);
            this.PB_Excel.Name = "PB_Excel";
            this.PB_Excel.Size = new System.Drawing.Size(452, 50);
            this.PB_Excel.TabIndex = 1;
            this.PB_Excel.Click += new System.EventHandler(this.PB_Excel_Click);
            // 
            // RTB_ExcelLog
            // 
            this.RTB_ExcelLog.Location = new System.Drawing.Point(24, 23);
            this.RTB_ExcelLog.Name = "RTB_ExcelLog";
            this.RTB_ExcelLog.Size = new System.Drawing.Size(556, 330);
            this.RTB_ExcelLog.TabIndex = 2;
            this.RTB_ExcelLog.Text = "";
            this.RTB_ExcelLog.TextChanged += new System.EventHandler(this.RTB_ExcelLog_TextChanged);
            // 
            // openFileExcel
            // 
            this.openFileExcel.FileName = "openFileExcel";
            // 
            // FormExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(687, 466);
            this.Controls.Add(this.RTB_ExcelLog);
            this.Controls.Add(this.PB_Excel);
            this.Controls.Add(this.bt_selct_excel);
            this.Name = "FormExcel";
            this.Text = "FormExcel";
            this.Load += new System.EventHandler(this.FormExcel_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button bt_selct_excel;
        private System.Windows.Forms.ProgressBar PB_Excel;
        private System.Windows.Forms.RichTextBox RTB_ExcelLog;
        private System.Windows.Forms.OpenFileDialog openFileExcel;
    }
}