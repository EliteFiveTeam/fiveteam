﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using excel = Microsoft.Office.Interop.Excel; // подключение библиотеки excel и создание псевдонима "Alias"
using word = Microsoft.Office.Interop.Word; // подключение библиотеки word и создание псевдонима "Alias"
using System.Diagnostics;

namespace RPD
{
    public partial class FormWord : Form
    {
        private string _FileNaim; 
        public string FileNaim // путь к шаблону НРП
        {
            get { return _FileNaim; }
            set { _FileNaim = value; }
        }

        public static bool btn1;
        Tema tems;
        Discipline dis;
        public Dis D = new Dis(); /*Класс*/
        char[] MyChar = { '\f', '\n', '\r', '\t', '\v', '\0', ' ', '2', '3', '.', ')', ';' };
        int CountKFind;  //' счетчик найденных фрагментов, n-сколько надо отсчитать нахождений до нужного
        word.Application WordApp;
        private int sec; // переменная, содержащая значение времени
        public FormWord()
        {
            InitializeComponent();
            sec = 0;
        }
        public string SearchText(string wordText1, string wordText2, int nf) // Поиск между двумя фрагментами - метод поиска 
        {
            Microsoft.Office.Interop.Word.Range r;//Range
            string st;
            st = "";
            r = WordApp.ActiveDocument.Range();
            bool f;
            f = false;
            int firstOccurence;
            firstOccurence = 0;
            CountKFind = 0;
            r.Find.ClearFormatting(); //Сброс форматирований из предыдущих операций поиска
            r.Find.Text = wordText1 + "*" + wordText2;
            r.Find.Forward = true;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.Format = false;
            r.Find.MatchCase = false;
            r.Find.MatchWholeWord = false;
            r.Find.MatchAllWordForms = false;
            r.Find.MatchSoundsLike = false;
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            while (r.Find.Execute() == true) // Проверка поиска, если нашёл фрагменты, то...
            {
                CountKFind = CountKFind + 1;// то счётчик найденных фрагментоd увеличивается на 1
                if (f)
                {
                    if (r.Start == firstOccurence)
                    { }
                    else
                    {
                        firstOccurence = r.Start;
                        f = true;
                    }
                }
                st = WordApp.ActiveDocument.Range(r.Start + wordText1.Length, r.End - wordText2.Length).Text; //убираем кл.
                r.Start = r.Start + wordText1.Length;
                r.End = r.End - wordText2.Length;
                if (CountKFind >= nf) // если нужный по счету фрагмент найден
                {
                    // r = WordApp.ActiveDocument.Range(r.Start, r.End);
                    break;
                }
            }

            CountKFind = 0;

            if (r.Text != "")
            {
                if (st != "")
                {
                    r.Copy();
                }
                else //' если текст не найден очистим буфер обмена
                {
                    Clipboard.Clear();
                }
            }
            else
            {
                {
                    Clipboard.Clear();
                }
            }

            return st;
        }

        private void FormWord_Load(object sender, EventArgs e)
        {

        }
        private void AnalysisOldProgramm()
        {
            string Filename_;
            WordApp = new word.Application(); // создаем объект word;
            WordApp.Visible = true; // показывает или скрывает файл word;
            openFileDialog1.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx";
            Action action = () => { openFileDialog1.ShowDialog(); }; Invoke(action);
            // фильтрует, оставляя только ворд файлы
            Filename_ = openFileDialog1.FileName;
            WordApp.Documents.Add(Filename_);// загружаем в word файл с рабочей книгой 
            Action action1 = () => { btn_OpenWp.Enabled = false; }; Invoke(action1);


            SearchText(textBox2.Text, textBox4.Text, CountKFind);
            int N = 0;
            int i = 0;
            int j = 0;
            progressBar1.Value = j;
            Microsoft.Office.Interop.Word.Range r;//Range
            Microsoft.Office.Interop.Word.ListParagraphs p;
            D.CreateLitera();
            string ss;
            ss = "";
            r = WordApp.ActiveDocument.Range();
            p = WordApp.ActiveDocument.ListParagraphs;
            word.Document document = WordApp.ActiveDocument;
            int NnN = document.ListParagraphs.Count;

            //Поиск литературы
            string str1 = "Основная литература";
            string str2 = "Дополнительная литература";
            string str3 = "Перечень";
            string gg1; string gg2;

            // Поиск основной литературы
            r.Find.Text = str1 + "*" + str2;
            r.Find.Forward = true;
            string f1 = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ

            if (r.Find.Execute(f1))// Проверка поиска, если нашёл фрагменты, то...
            {
                gg1 = WordApp.ActiveDocument.Range(r.Start + str1.Length, r.End - str2.Length).Text; //убираем кл.
                r.Start = r.Start + str1.Length;
                r.End = r.End - str2.Length;
                int m21 = r.ListParagraphs.Count;
                if (m21 != 0)
                {
                    object Start = r.ListParagraphs[1].Range.Start;
                    object End = r.ListParagraphs[m21].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    rtb_LiteraBasic.Paste();
                    for (int y = 1; y <= r.ListParagraphs.Count; y++)
                    {
                        string dfs = r.ListParagraphs[y].Range.Text;
                        D.MyListAdd(dfs, false);
                    }
                    Action Progress = () => { rtb_Log.AppendText("Основная литература считана\n", Color.Green); }; Invoke(Progress);

                }
                else
                {
                    Action Progress = () => { rtb_Log.AppendText("Основная литература не найдена\n", Color.Red); }; Invoke(Progress);
                }

            }
            // поиск дополнительной литературы
            r.Find.Text = str2 + "*" + str3;
            r.Find.Forward = true;
            string f2 = r.Find.Text;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
            if (r.Find.Execute(f2))// Проверка поиска, если нашёл фрагменты, то...
            {
                gg2 = WordApp.ActiveDocument.Range(r.Start + str2.Length, r.End - str3.Length).Text; //убираем кл.
                r.Start = r.Start + str2.Length;
                r.End = r.End - str3.Length;
                int m12 = r.ListParagraphs.Count;
                if (m12 != 0)
                {
                    object Start = r.ListParagraphs[1].Range.Start;
                    object End = r.ListParagraphs[m12].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    rtb_Add_Litera.Paste();
                    for (int x = 1; x <= r.ListParagraphs.Count; x++)
                    {
                        string dsf = r.ListParagraphs[x].Range.Text;
                        D.MyListAdd(dsf, true);
                    }
                    Action Progress = () => { rtb_Log.AppendText("Дополнительная литература считана\n", Color.Green); }; Invoke(Progress);
                }
                else
                {
                    Action Progress = () => { rtb_Log.AppendText("Основная литература не найдена\n", Color.Red); }; Invoke(Progress);
                }

            } // поиск закончился, литература записана в массив


            //находим цели дисциплины
            ss = SearchText("явля?????", "Учебные задачи дисциплины", 2);
            if (ss == "") //' Если цели не попали в оглавление
            {
                ss = SearchText("явля?????", "Учебные задачи дисциплины", 1); // искомый текст после оглавления
                if (ss == "")
                {
                    Action Progress = () => { rtb_Log.AppendText("Цели дисциплины не найдены\n", Color.Red); }; Invoke(Progress);
                }

            }
            else { Action Progress = () => { rtb_Log.AppendText("Цели дисциплины найдены\n", Color.Green); }; Invoke(Progress); }

            ss = ss.TrimEnd(MyChar);
            N = ss.IndexOf("явля");
            if (N > 0 && N < ss.Length - 9)
            {
                D.Cel = ss.Remove(1, N + 9);
            }
            else
            {
                D.Cel = ss;// записали переменную цель
            }



            //' Находим задачи и оставляем все после слова "является" или "являются:"
            ss = SearchText("Учебные задачи дисциплины", "Место дисциплины", 2);
            if (ss == "")// ' Если задачи не попали в оглавление
            {
                ss = SearchText("Учебные задачи дисциплины", "Место дисциплины", 1);
                if (ss == "")
                {
                    Action Progress = () => { rtb_Log.AppendText("Задачи дисциплины не найдены\n", Color.Red); }; Invoke(Progress);
                }

            }
            else { Action Progress = () => { rtb_Log.AppendText("Задачи дисциплины найдены\n", Color.Green); }; Invoke(Progress); }

            ss = ss.TrimEnd(MyChar);
            N = ss.IndexOf("явля");

            if (N > 0 && N < ss.Length - 9)
            {
                D.Tasks = ss.Remove(1, N + 9);
            }
            else
            {
                D.Tasks = ss; // записали цели
            }

            //Находим знания, умения и владения и оставляем все до знаков препинания и символов перевода, или цифр 2, 3.
            ss = SearchText("Знать:", "Уметь:", 1);
            D.Zn_before = ss.TrimEnd(MyChar);
            ss = SearchText("Уметь:", "Владеть:", 1);
            D.Um_before = ss.TrimEnd(MyChar);
            ss = SearchText("Владеть:", ".", 1);
            D.Vl_before = ss.TrimEnd(MyChar);
            ss = SearchText("Знать:", "Уметь:", 2);
            D.Zn_after = ss.TrimEnd(MyChar);
            ss = SearchText("Уметь:", "Владеть:", 2);
            D.Um_after = ss.TrimEnd(MyChar);
            ss = SearchText("Владеть:", ".", 2);
            D.Vl_after = ss.TrimEnd(MyChar);
            if (ss == "")
            {
                Action Progress = () => { rtb_Log.AppendText("Знания, умения, навыки до не найдены\n", Color.Red); }; Invoke(Progress);
            }

            byte razd = 1;  //'номер раздела
            int CountTems = 0;
            r.Find.Text = "Наименование";
            string texttable = r.Find.Text;
            if (WordApp.ActiveDocument.Tables.Count != 0)
            {
                try
                {
                    for (i = 1; i <= WordApp.ActiveDocument.Tables.Count; i++)
                    {
                        if (WordApp.ActiveDocument.Tables[i].Cell(1, 2).Range.Find.Execute(texttable))
                        {

                            Action Progress = () => { rtb_Log.AppendText("Таблица с темой " + i + " считана\n", Color.Green); }; Invoke(Progress);
                            for (int n = 2; n <= WordApp.ActiveDocument.Tables[i].Rows.Count; n++)
                            {
                                if (WordApp.ActiveDocument.Tables[i].Rows[n].Cells.Count >= 5)
                                {
                                    D.tems[i - 2].Name = WordApp.ActiveDocument.Tables[i].Cell(n, 2).Range.Text;
                                    D.tems[i - 2].Text = WordApp.ActiveDocument.Tables[i].Cell(n, 3).Range.Text;
                                    D.tems[i - 2].Rez = WordApp.ActiveDocument.Tables[i].Cell(n, 5).Range.Text;
                                    D.tems[i - 2].FormZ = WordApp.ActiveDocument.Tables[i].Cell(n, 6).Range.Text;
                                    CountTems++;
                                }
                            }
                            break;
                        }
                        else
                        {
                            Action Progress = () => { rtb_Log.AppendText("Таблица с темой " + i + " не найдена\n", Color.Red); }; Invoke(Progress);
                            if (i != 2)
                            {
                                razd += razd;  //' счетчик разделов срабатывает если их больше одного
                            }
                        }
                    }
                }
                catch { Action Progress = () => { rtb_Log.AppendText("Таблица с темой не найдена\n", Color.Red); }; Invoke(Progress); }

                D.Nt = CountTems; //Записали количество тем в дисциплине
            }
            else
            {
                Action Progress = () => { rtb_Log.AppendText("Таблица с темой не найдена\n", Color.Red); }; Invoke(Progress);
            }

            Clipboard.Clear();


            // считываются темы и их литература, вопросы для самопроверки


            ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ", 2);

            if (ss.Contains("Тема 1.") & ss.Contains("Литература") & ss.Contains("Вопросы для самопроверки"))
            {

                rtb_Tems.Paste();
                rtb_Log.AppendText("Перечень УМО считаны\n", Color.Green);
            }
            else
            {
                ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ", 1);
                if (ss.Contains("Тема 1.") & ss.Contains("Литература") & ss.Contains("Вопросы для самопроверки"))
                {
                    rtb_Tems.Paste();
                    rtb_Log.AppendText("Перечень УМО считаны\n", Color.Green);
                }
                else
                {
                    ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Рекомендуемые обучающие", 2);

                    if (ss.Contains("Тема 1.") & ss.Contains("Литература") & ss.Contains("Вопросы для самопроверки"))
                    {
                        rtb_Tems.Paste();
                        rtb_Log.AppendText("Перечень УМО считаны\n", Color.Green);
                    }
                    else
                    {
                        ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Рекомендуемые обучающие", 1);
                        if (ss.Contains("Тема 1.") & ss.Contains("Литература") & ss.Contains("Вопросы для самопроверки"))
                        {
                            rtb_Tems.Paste();
                            rtb_Log.AppendText("Перечень УМО считаны\n", Color.Green);
                        }
                        else
                        {
                            ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 2);
                            if (ss.Contains("Тема 1.") & ss.Contains("Литература") & ss.Contains("Вопросы для самопроверки"))
                            {
                                rtb_Tems.Paste();
                                rtb_Log.AppendText("Перечень УМО считаны\n", Color.Green);
                            }
                            else
                            {
                                ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 1);
                                if (ss.Contains("Тема 1.") & ss.Contains("Литература") & ss.Contains("Вопросы для самопроверки"))
                                {
                                    rtb_Tems.Paste();
                                    rtb_Log.AppendText("Перечень УМО считаны\n", Color.Green);
                                }
                                else
                                {
                                    rtb_Log.AppendText("Перечень УМО не считаны\n", Color.Red);
                                }
                            }
                        }
                    }
                }
            }
            

            Clipboard.Clear();


            //Поиск вопросов к экзамену/зачёту с учётом итогового контроля
            string exstr1 = "Вопросы к";
            string exstr2 = "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ";
            string exstr3 = "Итоговый контроль";
            string exgg1;
            ss = SearchText("Вопросы к", "Итоговый контроль", 1);

            if (ss != "")
            {
                // Поиск 
                r.Find.Text = exstr1 + "*" + exstr3;
                r.Find.Forward = true;
                string exf1 = r.Find.Text;
                r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
                r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ

                if (r.Find.Execute(exf1))// Проверка поиска, если нашёл фрагменты, то...
                {
                    exgg1 = WordApp.ActiveDocument.Range(r.Start + exstr1.Length, r.End - exstr3.Length).Text; //убираем кл.
                    r.Start = r.Start + exstr1.Length;
                    r.End = r.End - exstr3.Length;
                    int exm21 = r.ListParagraphs.Count;
                    object Start = r.ListParagraphs[1].Range.Start;
                    object End = r.ListParagraphs[exm21].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    rtb_ForExam.Paste();
                    Action Progress = () => { rtb_Log.AppendText("Вопросы к экзамену считаны\n", Color.Green); }; Invoke(Progress);
                    for (int y = 1; y <= r.ListParagraphs.Count; y++)
                    {
                        string dfs = r.ListParagraphs[y].Range.Text;
                        D.MyForExamAdd(dfs);
                    }

                }
            }
            else
            {
                r.Find.Text = exstr1 + "*" + exstr2;
                r.Find.Forward = true;
                string exf1 = r.Find.Text;
                r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
                r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ
                ss = SearchText("Вопросы к", "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ", 1);
                if (ss == "")
                {
                    Action Progress = () => { rtb_Log.AppendText("Вопросы для зачёта/экзамена не найдены\n", Color.Red); }; Invoke(Progress);
                }
                if (r.Find.Execute(exf1))// Проверка поиска, если нашёл фрагменты, то...
                {
                    exgg1 = WordApp.ActiveDocument.Range(r.Start + exstr1.Length, r.End - exstr2.Length).Text; //убираем кл.
                    r.Start = r.Start + exstr1.Length;
                    r.End = r.End - exstr2.Length;
                    int exm21 = r.ListParagraphs.Count;
                    object Start = r.ListParagraphs[1].Range.Start;
                    object End = r.ListParagraphs[exm21].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    rtb_ForExam.Paste();
                    Action Progress = () => { rtb_Log.AppendText("Вопросы к экзамену считаны\n", Color.Green); }; Invoke(Progress);
                    for (int y = 1; y <= r.ListParagraphs.Count; y++)
                    {
                        string dfs = r.ListParagraphs[y].Range.Text;
                        D.MyForExamAdd(dfs);
                    }
                }
            }

            ss = SearchText("Итоговый контроль", "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ", 1);
            if (ss != "")
            {
                // Поиск 
                r.Find.Text = exstr3 + "*" + exstr2;
                r.Find.Forward = true;
                string exf1 = r.Find.Text;
                r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
                r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ

                if (r.Find.Execute(exf1))// Проверка поиска, если нашёл фрагменты, то...
                {
                    exgg1 = WordApp.ActiveDocument.Range(r.Start, r.End - exstr2.Length).Text; //убираем кл.
                    r.Start = r.Start;
                    r.End = r.End - exstr2.Length;
                    int exm21 = r.Paragraphs.Count;
                    object Start = r.Paragraphs[1].Range.Start;
                    object End = r.Paragraphs[exm21].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    rtb_ForExam.Paste();
                    Action Progress = () => { rtb_Log.AppendText("Итоговый контроль найден\n", Color.Green); }; Invoke(Progress);
                }
            }
            else
            {
                Action Progress = () => { rtb_Log.AppendText("Итоговый контроль не найден\n", Color.Red); }; Invoke(Progress);
            }


        }
      
        private void CreateNewProgram()
        {
            WordApp = new word.Application(); // создаем объект word;
            FormMain FM = new FormMain();
            string Check = FileNaim;
            
        }
        
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (sec == 2)
            {
                sec = 0;
                btn_OpenWp.Enabled = true;
                timer1.Stop();
            }
            else
                sec++;

        }

        private void btn_OpenWp_Click(object sender, EventArgs e)
        {
            AnalysisOldProgramm();
            WordApp.Quit();
        }

        private void bt_create_newrp_Click(object sender, EventArgs e)
        {
            CreateNewProgram();
        }
    }
}
    
