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
using System.IO;

namespace RPD
{
    public partial class FormMain : Form
    {
        connection_to_bd BD = new connection_to_bd();
        Plan PL; // Переменная структуры "Титул"
        FormWord FW = new FormWord();
        word.Application WordApp;
       
       
        
       
        public FormMain()
        {
            InitializeComponent();
            DataBase();
            if (lst_prof.Items.Count > 0)
            {
                lst_prof.SetSelected(0, true);
            }
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
            Thread.Sleep(500); // остановка потока для обращения к БД
            DataBase();
        }

        public bool AnalysisPattern(bool Flag)
        {
            Microsoft.Office.Interop.Word.Range r;
            r = WordApp.ActiveDocument.Range();
            r.Find.ClearFormatting(); //Сброс форматирований из предыдущих операций поиска 
            r.Find.Forward = true;
            r.Find.Format = true;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            Flag = false;
            r.Find.Text = "#Индекс";
            string SearhWord1 = r.Find.Text;
            if (r.Find.Execute(SearhWord1) == true)
            {
                r.Find.Text = "#Дисциплина";
                string SearhWord2 = r.Find.Text;
                if (r.Find.Execute(SearhWord2) == true)
                {

                    r.Find.Text = "#Направление";
                    string SearhWord3 = r.Find.Text;
                    if (r.Find.Execute(SearhWord3) == true)
                    {
                        r.Find.Text = "#ФГОС";
                        string SearhWord4 = r.Find.Text;
                        if (r.Find.Execute(SearhWord4) == true)
                        {

                                r.Find.Text = "#Цели";
                                string SearhWord6 = r.Find.Text;
                                if (r.Find.Execute(SearhWord6) == true)
                                {
                                    r.Find.Text = "#Задачи";
                                    string SearhWord7 = r.Find.Text;
                                    if (r.Find.Execute(SearhWord7) == true)
                                    {
                                        r.Find.Text = "#Част?"; 
                                        string SearhWord8 = r.Find.Text;
                                        if (r.Find.Execute(SearhWord8) == true)
                                        {
                                            r.Find.Text = "#ДисциплиныДО";
                                            string SearhWord9 = r.Find.Text;
                                            if (r.Find.Execute(SearhWord9) == true)
                                            {
                                                r.Find.Text = "#ЗнатьДО";
                                                string SearhWord10 = r.Find.Text;
                                                if (r.Find.Execute(SearhWord10) == true)
                                                {
                                                    r.Find.Text = "#УметьДО";
                                                    string SearhWord11 = r.Find.Text;
                                                    if (r.Find.Execute(SearhWord11) == true)
                                                    {
                                                        r.Find.Text = "#ВладетьДО";
                                                        string SearhWord12 = r.Find.Text;
                                                        if (r.Find.Execute(SearhWord12) == true)
                                                        {
                                                            r.Find.Text = "#ДисциплиныПосле";
                                                            string SearhWord13 = r.Find.Text;
                                                            if (r.Find.Execute(SearhWord13) == true)
                                                            {
                                                                r.Find.Text = "#зе";
                                                                string SearhWord14 = r.Find.Text;
                                                                if (r.Find.Execute(SearhWord14) == true)
                                                                {
                                                                    r.Find.Text = "#че";
                                                                    string SearhWord15 = r.Find.Text;
                                                                    if (r.Find.Execute(SearhWord15) == true)
                                                                    {
                                                                        r.Find.Text = "#конт";
                                                                        string SearhWord16 = r.Find.Text;
                                                                        if (r.Find.Execute(SearhWord16) == true)
                                                                        {
                                                                            r.Find.Text = "#аудит";
                                                                            string SearhWord17 = r.Find.Text;
                                                                            if (r.Find.Execute(SearhWord17) == true)
                                                                            {
                                                                                r.Find.Text = "#лек";
                                                                                string SearhWord18 = r.Find.Text;
                                                                                if (r.Find.Execute(SearhWord18) == true)
                                                                                {
                                                                                    r.Find.Text = "#лаб";
                                                                                    string SearhWord19 = r.Find.Text;
                                                                                    if (r.Find.Execute(SearhWord19) == true)
                                                                                    {
                                                                                        r.Find.Text = "#пр";
                                                                                        string SearhWord20 = r.Find.Text;
                                                                                        if (r.Find.Execute(SearhWord20) == true)
                                                                                        {
                                                                                            r.Find.Text = "#инт";
                                                                                            string SearhWord21 = r.Find.Text;
                                                                                            if (r.Find.Execute(SearhWord21) == true)
                                                                                            {
                                                                                                r.Find.Text = "#эл";
                                                                                                string SearhWord22 = r.Find.Text;
                                                                                                if (r.Find.Execute(SearhWord22) == true)
                                                                                                {
                                                                                                    r.Find.Text = "#срс";
                                                                                                    string SearhWord23 = r.Find.Text;
                                                                                                    if (r.Find.Execute(SearhWord23) == true)
                                                                                                    {
                                                                                                        r.Find.Text = "#конт";
                                                                                                        string SearhWord24 = r.Find.Text;
                                                                                                        if (r.Find.Execute(SearhWord24) == true)
                                                                                                        {
                                                                                                            r.Find.Text = "#кконтр";
                                                                                                        string SearhWord25 = r.Find.Text;
                                                                                                        if (r.Find.Execute(SearhWord24) == true)
                                                                                                        {
                                                                                                            r.Find.Text = "#Основная_л";
                                                                                                            string SearhWord26 = r.Find.Text;
                                                                                                            if (r.Find.Execute(SearhWord25) == true)
                                                                                                            {
                                                                                                                r.Find.Text = "#Дополнит_л";
                                                                                                                string SearhWord27 = r.Find.Text;
                                                                                                                if (r.Find.Execute(SearhWord26) == true)
                                                                                                                {
                                                                                                                    r.Find.Text = "#Посещение балла";
                                                                                                                    string SearhWord28 = r.Find.Text;
                                                                                                                    if (r.Find.Execute(SearhWord27) == true)
                                                                                                                    {
                                                                                                                        lb_path_rp.Text = "Шаблон корректен\n";
                                                                                                                        return Flag = true;
                                                                                                                    }
                                                                                                                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                                                }
                                                                                                                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                                            }
                                                                                                            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                                        }
                                                                                                        else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                                    }
                                                                                                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                                }
                                                                                                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                            }
                                                                                            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                        }
                                                                                        else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                    }
                                                                                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                                }
                                                                                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                            }
                                                                            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                        }
                                                                        else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                    }
                                                                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                                }
                                                                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                            }
                                                            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                        }
                                                        else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                    }
                                                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                                }
                                                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                            }
                                            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                        }
                                        else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                    }
                                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                                }
                                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                            }
                            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                        }
                        else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                    }
                    else { lb_path_rp.Text = "Шаблон не корректен\n"; }
                }
                else { lb_path_rp.Text = "Шаблон не корректен\n"; }
            }
            else { lb_path_rp.Text = "Шаблон не корректен\n"; }
            return Flag = false;





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
                FW.ID_Prof = PL.ID;
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
            FW.ShowDialog();

        }

        
        private void bt_select_rp_Click(object sender, EventArgs e)
        {
            WordApp = new word.Application(); // создаем объект word;
            openFileWord.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx";
            if (openFileWord.ShowDialog() == DialogResult.OK)
            { FW.FileNaim = openFileWord.FileName;
            lb_path_rp.Text = Path.GetFileNameWithoutExtension(FW.FileNaim) + " загружен";

            } // открытие шаблона Новой РП 
            else
            {
                return;
            }
            
            WordApp.Documents.Add(FW.FileNaim);
            //if (AnalysisPattern(true))
            //{
            //    /*Если шаблон вернёт значение true, то он корректен и мы можем приступить к замене слов(для замены создан специальный метод выше)*/
            //}
             
        }

        private void bt_select_fos_Click(object sender, EventArgs e)
        {
            WordApp = new word.Application(); // создаем объект word;
            WordApp.Visible = true;
            openFileWord.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx";
            if (openFileWord.ShowDialog() == DialogResult.OK)
            { FW.FileNaim_FOS = openFileWord.FileName;
            lb_path_fos.Text = Path.GetFileNameWithoutExtension(FW.FileNaim_FOS) + " загружен";
            } // Шаблон ФОС
            if (FW.FileNaim_FOS == null)
            {
                return;
            }
            
            WordApp.Documents.Add(FW.FileNaim_FOS);
        }

        private void bt_select_anat_Click(object sender, EventArgs e)
        {
            WordApp = new word.Application(); // создаем объект word;
            openFileWord.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx";
            if (openFileWord.ShowDialog() == DialogResult.OK)
            { FW.FileNaim_ANAT = openFileWord.FileName;
            lb_path_anat.Text = Path.GetFileNameWithoutExtension(FW.FileNaim_ANAT) + " загружен";
            } // Шаблон АНАТ
            if (FW.FileNaim_ANAT == null)
            {
                return;
            }
            WordApp.Documents.Add(FW.FileNaim_ANAT);
        }
  
    }
}
