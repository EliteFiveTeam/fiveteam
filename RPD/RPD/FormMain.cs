using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using excel = Microsoft.Office.Interop.Excel; // подключение библиотеки excel и создание псевдонима "Alias"
using word = Microsoft.Office.Interop.Word; // подключение библиотеки word и создание псевдонима "Alias"
using System.Threading;
using System.Diagnostics;

namespace RPD
{
    public partial class FormMain : Form
    {
        connection_to_bd BD = new connection_to_bd();
        Plan PL; // Переменная структуры "Титул"
        DataAccess DA;
        word.Application WordApp;
        FormWord FW = new FormWord();
       
       
        
       
        public FormMain()
        {
            InitializeComponent();
            DataBase();
        }
        public void DataBase() // Добавление в ListBox1
        {
            lst_prof.Items.Clear();
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
            
            FW.Show();
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
            lst_prof.Items.Clear();
            clst_disc.Items.Clear();
            Thread.Sleep(500);
            DataBase();
        }

        

        private void lst_prof_SelectedIndexChanged(object sender, EventArgs e)
        {
            clst_disc.Items.Clear();
            string Nazv = lst_prof.Text.Substring(0, lst_prof.Text.Length - 5).Trim();
            string god = lst_prof.Text.Substring(lst_prof.Text.Length - 5).Trim();
            BD.Connect();
            BD.command.CommandText = "SELECT Профиль.Название_профиля, Профиль.Год_профиля,Профиль.Код FROM Профиль WHERE (((Профиль.Название_профиля)='" + Nazv + "') AND ((Профиль.Год_профиля)='" + god + "'));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                PL.ID = Convert.ToInt32(BD.reader["Код"]);
            }
            BD.reader.Close();
            BD.command.CommandText = "SELECT Дисциплины_профиля.Дисциплины, Дисциплины_профиля.Код_профиля FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код_профиля)=" + PL.ID + "));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                clst_disc.Items.Add(BD.reader["Дисциплины"]);
            }
            BD.reader.Close();
        }

        private void clst_disc_SelectedIndexChanged(object sender, EventArgs e)
        {
            string id_disp = clst_disc.Text;
            BD.Connect();
            BD.command.CommandText = "SELECT Дисциплины_профиля.Код FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код_профиля)=" + PL.ID + ") AND ((Дисциплины_профиля.Дисциплины)='" + id_disp + "'));";
            BD.reader = BD.command.ExecuteReader();
            // берем id дисциплины выброной из clst_disc
            while (BD.reader.Read())
            {
                FW.ID = Convert.ToInt32(BD.reader["Код"]);
               
            }
            BD.reader.Close();
            
        }

        private void bt_select_Click(object sender, EventArgs e)
        {
           
            FW.fillingMainData(); // добавление в структуру DataAccess из БД 
        }

        
        private void bt_select_rp_Click(object sender, EventArgs e)
        {
           
            openFileWord.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx";
            openFileWord.ShowDialog();
            FW.FileNaim = openFileWord.FileName; // открытие шаблона Новой РП
             
        }
  
    }
}
