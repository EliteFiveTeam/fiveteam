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
        Plan PL; // Переменная структуры "Титул"
       
       
        public FormMain()
        {
            InitializeComponent();
            DataBase();
        }
        public void DataBase() // Добавление в ListBox1
        {
            BD.Connect();
            BD.command.CommandText = "SELECT * FROM Профиль ;";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                lst_prof.Items.Add(BD.reader["Название_профиля"].ToString() + " " + BD.reader["Год_профиля"].ToString());
            }
        }

        private void bt_createRP_Click(object sender, EventArgs e)
        {
            FormWord fm = new FormWord();
            fm.Show();
        }

        private void bt_addprof_Click(object sender, EventArgs e)
        {
            FormExcel fm = new FormExcel();
            fm.Owner = this;
            fm.ShowDialog();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {

        }

        private void bt_del_bd_Click(object sender, EventArgs e)
        {
            BD.Connect();
            BD.command.CommandText = "DELETE Профиль.Код, Профиль.Название_профиля, Профиль.Год_профиля FROM Профиль WHERE (((Профиль.Код)=" + PL.ID + "));";
            BD.reader = BD.command.ExecuteReader();
            BD.reader.Close();
        }

        private void lst_prof_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Nazv = lst_prof.Text.Substring(0, lst_prof.Text.Length - 5).Trim();
            string god = lst_prof.Text.Substring(lst_prof.Text.Length - 5).Trim();
            BD.Connect();
            BD.command.CommandText = "SELECT Профиль.Название_профиля, Профиль.Год_профиля,Профиль.Код FROM Профиль WHERE (((Профиль.Название_профиля)='" + Nazv + "') AND ((Профиль.Год_профиля)='" + god + "'));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                PL.ID = Convert.ToInt32(BD.reader["Код"]);
            }
        }
    }
}
