namespace RPD
{
    partial class FormMain
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.lb_prof = new System.Windows.Forms.Label();
            this.lb_disc = new System.Windows.Forms.Label();
            this.lst_prof = new System.Windows.Forms.ListBox();
            this.clst_disc = new System.Windows.Forms.CheckedListBox();
            this.bt_addprof = new System.Windows.Forms.Button();
            this.bt_select = new System.Windows.Forms.Button();
            this.bt_del_bd = new System.Windows.Forms.Button();
            this.lb_path_fos = new System.Windows.Forms.Label();
            this.lb_path_anat = new System.Windows.Forms.Label();
            this.lb_path_rp = new System.Windows.Forms.Label();
            this.bt_select_anat = new System.Windows.Forms.Button();
            this.bt_select_fos = new System.Windows.Forms.Button();
            this.lb_anat = new System.Windows.Forms.Label();
            this.lb_fos = new System.Windows.Forms.Label();
            this.bt_select_rp = new System.Windows.Forms.Button();
            this.lb_rp = new System.Windows.Forms.Label();
            this.openFileWord = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lb_prof
            // 
            this.lb_prof.AutoSize = true;
            this.lb_prof.Location = new System.Drawing.Point(43, 21);
            this.lb_prof.Name = "lb_prof";
            this.lb_prof.Size = new System.Drawing.Size(53, 13);
            this.lb_prof.TabIndex = 0;
            this.lb_prof.Text = "Профиль";
            // 
            // lb_disc
            // 
            this.lb_disc.AutoSize = true;
            this.lb_disc.Location = new System.Drawing.Point(336, 21);
            this.lb_disc.Name = "lb_disc";
            this.lb_disc.Size = new System.Drawing.Size(72, 13);
            this.lb_disc.TabIndex = 1;
            this.lb_disc.Text = "Дисциплины";
            // 
            // lst_prof
            // 
            this.lst_prof.FormattingEnabled = true;
            this.lst_prof.Location = new System.Drawing.Point(21, 46);
            this.lst_prof.Name = "lst_prof";
            this.lst_prof.Size = new System.Drawing.Size(281, 329);
            this.lst_prof.TabIndex = 2;
            this.lst_prof.SelectedIndexChanged += new System.EventHandler(this.lst_prof_SelectedIndexChanged);
            // 
            // clst_disc
            // 
            this.clst_disc.FormattingEnabled = true;
            this.clst_disc.Location = new System.Drawing.Point(308, 46);
            this.clst_disc.Name = "clst_disc";
            this.clst_disc.Size = new System.Drawing.Size(328, 334);
            this.clst_disc.TabIndex = 3;
            this.clst_disc.SelectedIndexChanged += new System.EventHandler(this.clst_disc_SelectedIndexChanged);
            // 
            // bt_addprof
            // 
            this.bt_addprof.Location = new System.Drawing.Point(21, 381);
            this.bt_addprof.Name = "bt_addprof";
            this.bt_addprof.Size = new System.Drawing.Size(134, 33);
            this.bt_addprof.TabIndex = 4;
            this.bt_addprof.Text = "Добавить профиль";
            this.bt_addprof.UseVisualStyleBackColor = true;
            this.bt_addprof.Click += new System.EventHandler(this.bt_addprof_Click);
            // 
            // bt_select
            // 
            this.bt_select.Location = new System.Drawing.Point(542, 392);
            this.bt_select.Name = "bt_select";
            this.bt_select.Size = new System.Drawing.Size(94, 35);
            this.bt_select.TabIndex = 6;
            this.bt_select.Text = "Выбрать";
            this.bt_select.UseVisualStyleBackColor = true;
            this.bt_select.Click += new System.EventHandler(this.bt_select_Click);
            // 
            // bt_del_bd
            // 
            this.bt_del_bd.Location = new System.Drawing.Point(161, 380);
            this.bt_del_bd.Name = "bt_del_bd";
            this.bt_del_bd.Size = new System.Drawing.Size(86, 47);
            this.bt_del_bd.TabIndex = 7;
            this.bt_del_bd.Text = "Удалить профиль";
            this.bt_del_bd.UseVisualStyleBackColor = true;
            this.bt_del_bd.Click += new System.EventHandler(this.bt_del_bd_Click);
            // 
            // lb_path_fos
            // 
            this.lb_path_fos.AutoSize = true;
            this.lb_path_fos.Location = new System.Drawing.Point(285, 485);
            this.lb_path_fos.Name = "lb_path_fos";
            this.lb_path_fos.Size = new System.Drawing.Size(0, 13);
            this.lb_path_fos.TabIndex = 24;
            // 
            // lb_path_anat
            // 
            this.lb_path_anat.AutoSize = true;
            this.lb_path_anat.Location = new System.Drawing.Point(285, 511);
            this.lb_path_anat.Name = "lb_path_anat";
            this.lb_path_anat.Size = new System.Drawing.Size(0, 13);
            this.lb_path_anat.TabIndex = 23;
            // 
            // lb_path_rp
            // 
            this.lb_path_rp.AutoSize = true;
            this.lb_path_rp.Location = new System.Drawing.Point(285, 450);
            this.lb_path_rp.Name = "lb_path_rp";
            this.lb_path_rp.Size = new System.Drawing.Size(0, 13);
            this.lb_path_rp.TabIndex = 22;
            // 
            // bt_select_anat
            // 
            this.bt_select_anat.Location = new System.Drawing.Point(187, 511);
            this.bt_select_anat.Name = "bt_select_anat";
            this.bt_select_anat.Size = new System.Drawing.Size(75, 23);
            this.bt_select_anat.TabIndex = 21;
            this.bt_select_anat.Text = "Обзор";
            this.bt_select_anat.UseVisualStyleBackColor = true;
            this.bt_select_anat.Click += new System.EventHandler(this.bt_select_anat_Click);
            // 
            // bt_select_fos
            // 
            this.bt_select_fos.Location = new System.Drawing.Point(187, 475);
            this.bt_select_fos.Name = "bt_select_fos";
            this.bt_select_fos.Size = new System.Drawing.Size(75, 23);
            this.bt_select_fos.TabIndex = 20;
            this.bt_select_fos.Text = "Обзор";
            this.bt_select_fos.UseVisualStyleBackColor = true;
            this.bt_select_fos.Click += new System.EventHandler(this.bt_select_fos_Click);
            // 
            // lb_anat
            // 
            this.lb_anat.AutoSize = true;
            this.lb_anat.Location = new System.Drawing.Point(11, 511);
            this.lb_anat.Name = "lb_anat";
            this.lb_anat.Size = new System.Drawing.Size(96, 13);
            this.lb_anat.TabIndex = 19;
            this.lb_anat.Text = "Шаблон анатации";
            // 
            // lb_fos
            // 
            this.lb_fos.AutoSize = true;
            this.lb_fos.Location = new System.Drawing.Point(11, 480);
            this.lb_fos.Name = "lb_fos";
            this.lb_fos.Size = new System.Drawing.Size(75, 13);
            this.lb_fos.TabIndex = 18;
            this.lb_fos.Text = "Шаблон ФОС";
            // 
            // bt_select_rp
            // 
            this.bt_select_rp.Location = new System.Drawing.Point(187, 445);
            this.bt_select_rp.Name = "bt_select_rp";
            this.bt_select_rp.Size = new System.Drawing.Size(75, 23);
            this.bt_select_rp.TabIndex = 17;
            this.bt_select_rp.Text = "Обзор";
            this.bt_select_rp.UseVisualStyleBackColor = true;
            this.bt_select_rp.Click += new System.EventHandler(this.bt_select_rp_Click);
            // 
            // lb_rp
            // 
            this.lb_rp.AutoSize = true;
            this.lb_rp.Location = new System.Drawing.Point(11, 450);
            this.lb_rp.Name = "lb_rp";
            this.lb_rp.Size = new System.Drawing.Size(152, 13);
            this.lb_rp.TabIndex = 16;
            this.lb_rp.Text = "Шаблон рабочей программы";
            // 
            // openFileWord
            // 
            this.openFileWord.FileName = "openFileWord";
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(779, 554);
            this.Controls.Add(this.lb_path_fos);
            this.Controls.Add(this.lb_path_anat);
            this.Controls.Add(this.lb_path_rp);
            this.Controls.Add(this.bt_select_anat);
            this.Controls.Add(this.bt_select_fos);
            this.Controls.Add(this.lb_anat);
            this.Controls.Add(this.lb_fos);
            this.Controls.Add(this.bt_select_rp);
            this.Controls.Add(this.lb_rp);
            this.Controls.Add(this.bt_del_bd);
            this.Controls.Add(this.bt_select);
            this.Controls.Add(this.bt_addprof);
            this.Controls.Add(this.clst_disc);
            this.Controls.Add(this.lst_prof);
            this.Controls.Add(this.lb_disc);
            this.Controls.Add(this.lb_prof);
            this.Name = "FormMain";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lb_prof;
        private System.Windows.Forms.Label lb_disc;
        private System.Windows.Forms.Button bt_addprof;
        private System.Windows.Forms.Button bt_select;
        private System.Windows.Forms.Button bt_del_bd;
        private System.Windows.Forms.Label lb_path_fos;
        private System.Windows.Forms.Label lb_path_anat;
        private System.Windows.Forms.Label lb_path_rp;
        private System.Windows.Forms.Button bt_select_anat;
        private System.Windows.Forms.Button bt_select_fos;
        private System.Windows.Forms.Label lb_anat;
        private System.Windows.Forms.Label lb_fos;
        private System.Windows.Forms.Button bt_select_rp;
        private System.Windows.Forms.Label lb_rp;
        public System.Windows.Forms.ListBox lst_prof;
        public System.Windows.Forms.CheckedListBox clst_disc;
        private System.Windows.Forms.OpenFileDialog openFileWord;
    }
}

