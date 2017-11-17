namespace RPD
{
    partial class FormWord
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
            this.components = new System.ComponentModel.Container();
            this.rtb_Add_Litera = new System.Windows.Forms.RichTextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.rtb_LiteraBasic = new System.Windows.Forms.RichTextBox();
            this.rtb_Tems = new System.Windows.Forms.RichTextBox();
            this.rtb_Log = new System.Windows.Forms.RichTextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.rtb_ForExam = new System.Windows.Forms.RichTextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.btn_OpenWp = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.bt_create_newrp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // rtb_Add_Litera
            // 
            this.rtb_Add_Litera.Location = new System.Drawing.Point(12, 510);
            this.rtb_Add_Litera.Name = "rtb_Add_Litera";
            this.rtb_Add_Litera.Size = new System.Drawing.Size(505, 116);
            this.rtb_Add_Litera.TabIndex = 23;
            this.rtb_Add_Litera.Text = "";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(532, 578);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(388, 39);
            this.progressBar1.TabIndex = 22;
            // 
            // rtb_LiteraBasic
            // 
            this.rtb_LiteraBasic.Location = new System.Drawing.Point(12, 377);
            this.rtb_LiteraBasic.Name = "rtb_LiteraBasic";
            this.rtb_LiteraBasic.Size = new System.Drawing.Size(505, 127);
            this.rtb_LiteraBasic.TabIndex = 21;
            this.rtb_LiteraBasic.Text = "";
            // 
            // rtb_Tems
            // 
            this.rtb_Tems.Location = new System.Drawing.Point(12, 194);
            this.rtb_Tems.Name = "rtb_Tems";
            this.rtb_Tems.Size = new System.Drawing.Size(504, 177);
            this.rtb_Tems.TabIndex = 20;
            this.rtb_Tems.Text = "";
            // 
            // rtb_Log
            // 
            this.rtb_Log.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.rtb_Log.BackColor = System.Drawing.SystemColors.Window;
            this.rtb_Log.Location = new System.Drawing.Point(734, 2);
            this.rtb_Log.Name = "rtb_Log";
            this.rtb_Log.ReadOnly = true;
            this.rtb_Log.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical;
            this.rtb_Log.Size = new System.Drawing.Size(382, 460);
            this.rtb_Log.TabIndex = 19;
            this.rtb_Log.Text = "";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(940, 564);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(191, 23);
            this.textBox4.TabIndex = 18;
            this.textBox4.Visible = false;
            // 
            // rtb_ForExam
            // 
            this.rtb_ForExam.Location = new System.Drawing.Point(12, 2);
            this.rtb_ForExam.Name = "rtb_ForExam";
            this.rtb_ForExam.Size = new System.Drawing.Size(504, 186);
            this.rtb_ForExam.TabIndex = 17;
            this.rtb_ForExam.Text = "";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(940, 593);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(191, 24);
            this.textBox2.TabIndex = 16;
            this.textBox2.Visible = false;
            // 
            // btn_OpenWp
            // 
            this.btn_OpenWp.Location = new System.Drawing.Point(532, 525);
            this.btn_OpenWp.Name = "btn_OpenWp";
            this.btn_OpenWp.Size = new System.Drawing.Size(154, 47);
            this.btn_OpenWp.TabIndex = 15;
            this.btn_OpenWp.Text = "Открыть старую РП и проанализировать";
            this.btn_OpenWp.UseVisualStyleBackColor = true;
            this.btn_OpenWp.Click += new System.EventHandler(this.btn_OpenWp_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // bt_create_newrp
            // 
            this.bt_create_newrp.Location = new System.Drawing.Point(940, 510);
            this.bt_create_newrp.Name = "bt_create_newrp";
            this.bt_create_newrp.Size = new System.Drawing.Size(123, 38);
            this.bt_create_newrp.TabIndex = 24;
            this.bt_create_newrp.Text = "Создать новую РП";
            this.bt_create_newrp.UseVisualStyleBackColor = true;
            this.bt_create_newrp.Click += new System.EventHandler(this.bt_create_newrp_Click);
            // 
            // FormWord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1149, 653);
            this.Controls.Add(this.bt_create_newrp);
            this.Controls.Add(this.rtb_Add_Litera);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.rtb_LiteraBasic);
            this.Controls.Add(this.rtb_Tems);
            this.Controls.Add(this.rtb_Log);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.rtb_ForExam);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.btn_OpenWp);
            this.Name = "FormWord";
            this.Text = "FormWord";
            this.Load += new System.EventHandler(this.FormWord_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox rtb_Add_Litera;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.RichTextBox rtb_LiteraBasic;
        private System.Windows.Forms.RichTextBox rtb_Tems;
        private System.Windows.Forms.RichTextBox rtb_Log;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.RichTextBox rtb_ForExam;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btn_OpenWp;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button bt_create_newrp;
    }
}