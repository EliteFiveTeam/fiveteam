using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Threading;
using System.Diagnostics;

namespace RPD
{
    public partial class FormMain : Form
    {
        connection_to_bd BD = new connection_to_bd();
        
       
       
        public FormMain()
        {
            InitializeComponent();
        }
        public void DataBase() // Добавление в ListBox1
        {
       
           
        }

        private void bt_createRP_Click(object sender, EventArgs e)
        {
            FormWord fm = new FormWord();
            fm.Show();
        }

        private void bt_addprof_Click(object sender, EventArgs e)
        {
            FormExcel fm = new FormExcel();
            fm.Show();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {

        }
    }
}
