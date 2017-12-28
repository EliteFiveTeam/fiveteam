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
            this.tab_Analysis_Pr = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btn_Clear = new System.Windows.Forms.Button();
            this.Create_Ticket = new System.Windows.Forms.Button();
            this.bt_create_newfos = new System.Windows.Forms.Button();
            this.Create_ANOT = new System.Windows.Forms.Button();
            this.tab_Analysis_Pr.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.SuspendLayout();
            // 
            // rtb_Add_Litera
            // 
            this.rtb_Add_Litera.Location = new System.Drawing.Point(0, 0);
            this.rtb_Add_Litera.Name = "rtb_Add_Litera";
            this.rtb_Add_Litera.Size = new System.Drawing.Size(900, 421);
            this.rtb_Add_Litera.TabIndex = 23;
            this.rtb_Add_Litera.Text = "";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1025, 602);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(388, 39);
            this.progressBar1.TabIndex = 22;
            this.progressBar1.Visible = false;
            // 
            // rtb_LiteraBasic
            // 
            this.rtb_LiteraBasic.Location = new System.Drawing.Point(0, 0);
            this.rtb_LiteraBasic.Name = "rtb_LiteraBasic";
            this.rtb_LiteraBasic.Size = new System.Drawing.Size(900, 421);
            this.rtb_LiteraBasic.TabIndex = 21;
            this.rtb_LiteraBasic.Text = "";
            // 
            // rtb_Tems
            // 
            this.rtb_Tems.Location = new System.Drawing.Point(0, 0);
            this.rtb_Tems.Name = "rtb_Tems";
            this.rtb_Tems.Size = new System.Drawing.Size(900, 421);
            this.rtb_Tems.TabIndex = 20;
            this.rtb_Tems.Text = "";
            this.rtb_Tems.TextChanged += new System.EventHandler(this.rtb_Tems_TextChanged);
            // 
            // rtb_Log
            // 
            this.rtb_Log.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.rtb_Log.BackColor = System.Drawing.SystemColors.Window;
            this.rtb_Log.Location = new System.Drawing.Point(939, 37);
            this.rtb_Log.Name = "rtb_Log";
            this.rtb_Log.ReadOnly = true;
            this.rtb_Log.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical;
            this.rtb_Log.Size = new System.Drawing.Size(403, 421);
            this.rtb_Log.TabIndex = 19;
            this.rtb_Log.Text = "";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(1222, 526);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(191, 23);
            this.textBox4.TabIndex = 18;
            this.textBox4.Visible = false;
            // 
            // rtb_ForExam
            // 
            this.rtb_ForExam.Location = new System.Drawing.Point(0, 0);
            this.rtb_ForExam.Name = "rtb_ForExam";
            this.rtb_ForExam.Size = new System.Drawing.Size(900, 421);
            this.rtb_ForExam.TabIndex = 0;
            this.rtb_ForExam.Text = "";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(1222, 555);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(191, 24);
            this.textBox2.TabIndex = 16;
            this.textBox2.Visible = false;
            // 
            // btn_OpenWp
            // 
            this.btn_OpenWp.Location = new System.Drawing.Point(16, 468);
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
            this.bt_create_newrp.Enabled = false;
            this.bt_create_newrp.Location = new System.Drawing.Point(196, 468);
            this.bt_create_newrp.Name = "bt_create_newrp";
            this.bt_create_newrp.Size = new System.Drawing.Size(130, 47);
            this.bt_create_newrp.TabIndex = 24;
            this.bt_create_newrp.Text = "Создать новую РП";
            this.bt_create_newrp.UseVisualStyleBackColor = true;
            this.bt_create_newrp.Click += new System.EventHandler(this.bt_create_newrp_Click);
            // 
            // tab_Analysis_Pr
            // 
            this.tab_Analysis_Pr.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tab_Analysis_Pr.Controls.Add(this.tabPage1);
            this.tab_Analysis_Pr.Controls.Add(this.tabPage2);
            this.tab_Analysis_Pr.Controls.Add(this.tabPage3);
            this.tab_Analysis_Pr.Controls.Add(this.tabPage4);
            this.tab_Analysis_Pr.HotTrack = true;
            this.tab_Analysis_Pr.Location = new System.Drawing.Point(12, 12);
            this.tab_Analysis_Pr.Multiline = true;
            this.tab_Analysis_Pr.Name = "tab_Analysis_Pr";
            this.tab_Analysis_Pr.SelectedIndex = 0;
            this.tab_Analysis_Pr.Size = new System.Drawing.Size(908, 450);
            this.tab_Analysis_Pr.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
            this.tab_Analysis_Pr.TabIndex = 25;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.rtb_ForExam);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(900, 421);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Вопросы к зачёту/экзамену";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.rtb_Tems);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(900, 421);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Перечень УМО";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.rtb_LiteraBasic);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(900, 421);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Основная литература";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.rtb_Add_Litera);
            this.tabPage4.Location = new System.Drawing.Point(4, 25);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(900, 421);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Дополнительная литература";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btn_Clear
            // 
            this.btn_Clear.Enabled = false;
            this.btn_Clear.Location = new System.Drawing.Point(16, 520);
            this.btn_Clear.Name = "btn_Clear";
            this.btn_Clear.Size = new System.Drawing.Size(154, 37);
            this.btn_Clear.TabIndex = 26;
            this.btn_Clear.Text = "Очистить";
            this.btn_Clear.UseVisualStyleBackColor = true;
            this.btn_Clear.Click += new System.EventHandler(this.btn_Clear_Click);
            // 
            // Create_Ticket
            // 
            this.Create_Ticket.Enabled = false;
            this.Create_Ticket.Location = new System.Drawing.Point(196, 520);
            this.Create_Ticket.Name = "Create_Ticket";
            this.Create_Ticket.Size = new System.Drawing.Size(130, 37);
            this.Create_Ticket.TabIndex = 27;
            this.Create_Ticket.Text = "Создать билеты";
            this.Create_Ticket.UseVisualStyleBackColor = true;
            this.Create_Ticket.Click += new System.EventHandler(this.Create_Ticket_Click);
            // 
            // bt_create_newfos
            // 
            this.bt_create_newfos.Location = new System.Drawing.Point(333, 469);
            this.bt_create_newfos.Name = "bt_create_newfos";
            this.bt_create_newfos.Size = new System.Drawing.Size(131, 46);
            this.bt_create_newfos.TabIndex = 28;
            this.bt_create_newfos.Text = "Создать новую ФОС";
            this.bt_create_newfos.UseVisualStyleBackColor = true;
            this.bt_create_newfos.Click += new System.EventHandler(this.bt_create_newfos_Click);
            // 
            // Create_ANOT
            // 
            this.Create_ANOT.Location = new System.Drawing.Point(332, 520);
            this.Create_ANOT.Name = "Create_ANOT";
            this.Create_ANOT.Size = new System.Drawing.Size(132, 37);
            this.Create_ANOT.TabIndex = 29;
            this.Create_ANOT.Text = "Создать новую Аннотацию";
            this.Create_ANOT.UseVisualStyleBackColor = true;
            this.Create_ANOT.Click += new System.EventHandler(this.Create_ANOT_Click);
            // 
            // FormWord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1425, 653);
            this.Controls.Add(this.Create_ANOT);
            this.Controls.Add(this.bt_create_newfos);
            this.Controls.Add(this.Create_Ticket);
            this.Controls.Add(this.btn_Clear);
            this.Controls.Add(this.tab_Analysis_Pr);
            this.Controls.Add(this.bt_create_newrp);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.rtb_Log);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.btn_OpenWp);
            this.Name = "FormWord";
            this.Text = "FormWord";
            this.Load += new System.EventHandler(this.FormWord_Load);
            this.tab_Analysis_Pr.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
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
        private System.Windows.Forms.TabControl tab_Analysis_Pr;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button btn_Clear;
        private System.Windows.Forms.Button Create_Ticket;
        private System.Windows.Forms.Button bt_create_newfos;
        private System.Windows.Forms.Button Create_ANOT;
    }
}