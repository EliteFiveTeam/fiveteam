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
        FormWord fm = new FormWord();
       
        
       
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
            lst_prof.Items.Clear();
            clst_disc.Items.Clear();
            Thread.Sleep(500);
            DataBase();
        }

        public void fillingMainData()
        {
            BD.Connect();
            BD.command.CommandText = "SELECT Дисциплины_профиля.Дисциплины, Дисциплины_профиля.Индекс, Дисциплины_профиля.Факт_по_зет, Дисциплины_профиля.По_плану, Дисциплины_профиля.Контакт_часы, Дисциплины_профиля.Аудиторные, Дисциплины_профиля.Самостоятельная_работа, Дисциплины_профиля.Контроль, Дисциплины_профиля.Элект_часы, Дисциплины_профиля.Интер_часы, Дисциплины_профиля.Закрепленная_кафедра, Дисциплины_профиля.Код_профиля FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код)=" + DA.Id_disp + "));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.Naim = BD.reader["Дисциплины"].ToString();
                DA.Napr = BD.reader["Код_направления_подготовки"].ToString();
                DA.Index = BD.reader["Индекс"].ToString();
                DA.Fact = Convert.ToInt32(BD.reader["Факт_по_зет"]);
                DA.AtPlan = Convert.ToInt32(BD.reader["По_плану"]);
                DA.ContactHours = Convert.ToInt32(BD.reader["Контакт_часы"]);
                DA.Aud = Convert.ToInt32(BD.reader["Аудиторные"]);
                DA.SR = Convert.ToInt32(BD.reader["Самостоятельная_работа"]);
                DA.Contr = Convert.ToInt32(BD.reader["Контроль"]);
                DA.ElectHours = Convert.ToInt32(BD.reader["Элект_часы"]);
                DA.InterHours = Convert.ToInt32(BD.reader["Интер_часы"]);
                DA.Kafedra = BD.reader["Закрепленная_кафедра"].ToString();
                DA.ID = Convert.ToInt32(BD.reader["Код_профиля"]);
            }
            BD.reader.Close();

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
            BD.command.CommandText = "SELECT Дисциплины_профиля.Дисциплины, Дисциплины_профиля.Код_профиля FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код_профиля)=" + PL.ID +"));";
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
            BD.command.CommandText = "SELECT Дисциплины_профиля.Код FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код_профиля)="+PL.ID+") AND ((Дисциплины_профиля.Дисциплины)='"+id_disp +"'));";
            BD.reader = BD.command.ExecuteReader();
            // берем id дисциплины выброной из clst_disc
            while (BD.reader.Read())
            {
               DA.Id_disp = Convert.ToInt32(BD.reader["Код"]);
               
            }
            BD.reader.Close();
            
        }

        private void bt_select_Click(object sender, EventArgs e)
        {
            
            fillingMainData();
        }

        
        private void bt_select_rp_Click(object sender, EventArgs e)
        {
           
            openFileWord.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx";
            openFileWord.ShowDialog();
            fm.FileNaim = openFileWord.FileName; 
             
        }
  
    }
}
