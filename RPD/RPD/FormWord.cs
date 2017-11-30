using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Data.OleDb;
using excel = Microsoft.Office.Interop.Excel; // подключение библиотеки excel и создание псевдонима "Alias"
using word = Microsoft.Office.Interop.Word; // подключение библиотеки word и создание псевдонима "Alias"
using System.Diagnostics;

namespace RPD
{
    public partial class FormWord : Form
    {
        connection_to_bd BD = new connection_to_bd();
        DataAccess DA;
        private string _FileNaim;
        private string _FileNaim_FOS;
        private string _FileNaim_ANAT;
        private int _ID_Prof;
        public int ID_Prof // id профиля
        {
            get { return _ID_Prof; }
            set { _ID_Prof = value; }
        }
        private int _ID;
        public string FileNaim // путь к шаблону НРП
        {
            get { return _FileNaim; }
            set { _FileNaim = value; }
        }
        public string FileNaim_FOS // путь к шаблону НРП
        {
            get { return _FileNaim_FOS; }
            set { _FileNaim_FOS = value; }
        }
        public string FileNaim_ANAT // путь к шаблону НРП
        {
            get { return _FileNaim_ANAT; }
            set { _FileNaim_ANAT = value; }
        }
        public int ID
        {
            get { return _ID; }
            set { _ID = value; }
        } // id Дисциплины
        private int ID_Napr; // id Направление Подготовки

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
        public void Replace_Words_in_Pattern()
        {
            Microsoft.Office.Interop.Word.Range r;
            r = WordApp.ActiveDocument.Range();
            r.Find.ClearFormatting(); //Сброс форматирований из предыдущих операций поиска 
            r.Find.Forward = true;
            r.Find.Format = true;
            r.Find.Wrap = word.WdFindWrap.wdFindContinue; //при достижении конца документа поиск будет продолжаться с начала пока не будет достигнуто положение начала поиска
            r.Find.MatchWildcards = true;//подстановочные знаки ВКЛ

            ///*Здесь и далее замена ключевых слов в копии шаблона(РП) на нужные значения их excel и word*/

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
                        r.Find.Text = "#ДатаФГОС";
                        string SearhWord4 = r.Find.Text;
                        if (r.Find.Execute(SearhWord4) == true)
                        {
                            r.Find.Text = "#НомерФГОС";
                            string SearhWord5 = r.Find.Text;
                            if (r.Find.Execute(SearhWord5) == true)
                            {
                                r.Find.Text = "#Цели";
                                string SearhWord6 = r.Find.Text;
                                if (r.Find.Execute(SearhWord6) == true)
                                {
                                    r.Find.Text = "#Задачи";
                                    string SearhWord7 = r.Find.Text;
                                    if (r.Find.Execute(SearhWord7) == true)
                                    {
                                        r.Find.Text = "#Часть";
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
                                                                                                        r.Find.Text = "#контр";
                                                                                                        string SearhWord24 = r.Find.Text;
                                                                                                        if (r.Find.Execute(SearhWord24) == true)
                                                                                                        {
                                                                                                            r.Find.Text = "#Основная_л";
                                                                                                            string SearhWord25 = r.Find.Text;
                                                                                                            if (r.Find.Execute(SearhWord25) == true)
                                                                                                            {
                                                                                                                r.Find.Text = "#Дополнит_л";
                                                                                                                string SearhWord26 = r.Find.Text;
                                                                                                                if (r.Find.Execute(SearhWord26) == true)
                                                                                                                {
                                                                                                                    r.Find.Text = "#Посещение балла";
                                                                                                                    string SearhWord27 = r.Find.Text;
                                                                                                                    if (r.Find.Execute(SearhWord27) == true)
                                                                                                                    {
                                                                                                                        rtb_Log.AppendText("Шаблон корректен\n", Color.Green);
                                                                                                                        return Flag = true;
                                                                                                                    }
                                                                                                                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                                                }
                                                                                                                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                                            }
                                                                                                            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                                        }
                                                                                                        else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                                    }
                                                                                                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                                }
                                                                                                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                            }
                                                                                            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                        }
                                                                                        else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                    }
                                                                                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                                }
                                                                                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                            }
                                                                            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                        }
                                                                        else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                    }
                                                                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                                }
                                                                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                            }
                                                            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                        }
                                                        else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                    }
                                                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                                }
                                                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                            }
                                            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                        }
                                        else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                    }
                                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                                }
                                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                            }
                            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                        }
                        else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                    }
                    else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
                }
                else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
            }
            else { rtb_Log.AppendText("Шаблон некорректен\n", Color.Red); }
            return Flag = false;





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
                            
                            int k = 0; // счетчик кол-во тем
                            Action Progress = () => { rtb_Log.AppendText("Таблица с темой " + i + " считана\n", Color.Green); }; Invoke(Progress);
                            for (int n = 2; n <= WordApp.ActiveDocument.Tables[i].Rows.Count; n++)
                            {
                                
                                if (WordApp.ActiveDocument.Tables[i].Rows[n].Cells.Count >= 5)
                                {
                                    if (WordApp.ActiveDocument.Tables[i].Rows[n].Cells[2].Range.Text.Length >3) // проверка пустых значений названий тем
                                    {
                                        D.tems[k].Name = WordApp.ActiveDocument.Tables[i].Cell(n, 2).Range.Text;
                                        D.tems[k].Text = WordApp.ActiveDocument.Tables[i].Cell(n, 3).Range.Text;
                                        D.tems[k].Rez = WordApp.ActiveDocument.Tables[i].Cell(n, 5).Range.Text;
                                        D.tems[k].FormZ = WordApp.ActiveDocument.Tables[i].Cell(n, 6).Range.Text;
                                        CountTems++;
                                        k++; // кол-во тем
                                    }
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
                    rtb_Log.AppendText("Вопросы к экзамену считаны\n", Color.Green); 
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
                    rtb_Log.AppendText("Вопросы для зачёта/экзамена не найдены\n", Color.Red);
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
                    rtb_Log.AppendText("Вопросы к экзамену считаны\n", Color.Green);
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
                    rtb_Log.AppendText("Итоговый контроль найден\n", Color.Green); 
                }
            }
            else
            {
                rtb_Log.AppendText("Итоговый контроль не найден\n", Color.Red);
            }


        }

        private void FindReplace(string str_old, string str_new) // Замена фрагментов текста длинными кусками(больше 246 символ)
        {
            Microsoft.Office.Interop.Word.Range r;//Range
            r = WordApp.ActiveDocument.Range();
            r.Find.Text = str_old; // Находим слово которое нужно заменить
            if (str_new.Length > 246) // Проверка если длинна слова больше 246 символов 
            {
                string Str_long = str_new; // новая переменная для работы с кусками текста
                while (Str_long.Length > 0) // разьбиение строки на фрагменты и добавление в НРП
                {
                    if (Str_long.Length > 246) 
                    {
                        r.Find.Replacement.Text = Str_long.Substring(0, 245) + "<Text>"; 
                        Str_long = Str_long.Substring(245, Str_long.Length - 245);
                        r.Find.Execute(r.Find.Text, Replace: word.WdReplace.wdReplaceAll);
                        r.Find.Text = "<Text>"; // хештег для поиска замены
                    }
                    else // если осталось меньше 246, добавляем последний кусок текста
                    {
                        r.Find.Replacement.Text = Str_long.Substring(0, Str_long.Length);
                        r.Find.Execute(r.Find.Text, Replace: word.WdReplace.wdReplaceAll);
                        break;
                    }
                }
            }
            else
            {
                r.Find.Replacement.Text = str_new;
                r.Find.Execute(r.Find.Text, Replace: word.WdReplace.wdReplaceAll);
            }
            
        }

        private void CreateNewProgram() // работа с Новой РП
        {
            WordApp = new word.Application(); // создаем объект word;
            FormMain FM = new FormMain();
            string NRP = FileNaim;
            string FOS = FileNaim_FOS;
            string ANAT = FileNaim_ANAT;
            WordApp.Documents.Add(FileNaim);
            if (FileNaim_FOS != null)
            {
                WordApp.Documents.Add(FileNaim_FOS);
            }
            if (FileNaim_ANAT != null)
            {
                WordApp.Documents.Add(FileNaim_ANAT);
            }
            string Name_NRP = DA.Index + "_" +DA.Naim + "_" + DA.Profile + ".docx";
            WordApp.ActiveDocument.SaveAs2(Name_NRP);
            WordApp.Visible = true;
            FindReplace("#Направление", DA.Napr);
            FindReplace("#Индекс", DA.Index);
            FindReplace("#Дисциплина", DA.Naim);
            FindReplace("#Профиль", DA.Profile);
            FindReplace("#Цели", D.Cel);
            FindReplace("#Задачи", D.Tasks);
            foreach(string s in DA.PreDis)
            {
            FindReplace("#ДисциплиныДО", s);
            }
            FindReplace("#ЗнатьДО", D.Zn_before);
            FindReplace("#УметьДО", D.Um_before);
            FindReplace("#ВладетьДО", D.Vl_before);
            foreach (string s in DA.AfterDis)
            {
                FindReplace("#ДисциплиныПосле", s);
            }
            string FindTable = "Наименование темы дисциплины";
            for (int i = 1; i <= WordApp.ActiveDocument.Tables.Count; i++)
            {
                if (WordApp.ActiveDocument.Tables[i].Cell(1, 2).Range.Find.Execute(FindTable))
                {
                    for (int z = 2; z <= D.Nt+1; z++) // z - номер строки в таблице с темами
                    {
                        
                        WordApp.ActiveDocument.Tables[i].Cell(z, 2).Range.Text = D.tems[z-2].Name;
                        WordApp.ActiveDocument.Tables[i].Cell(z, 3).Range.Text = D.tems[z-2].Text;
                        WordApp.ActiveDocument.Tables[i].Cell(z, 5).Range.Text = D.tems[z-2].Rez;
                        WordApp.ActiveDocument.Tables[i].Cell(z, 4).Range.Text = D.tems[z-2].Comp;
                        // WordApp.ActiveDocument.Tables[i].Cell(z,2).Range.Text = D.tems[z].Name;
                        if (z != D.Nt + 1) WordApp.ActiveDocument.Tables[i].Rows.Add();
                    }
                }
            }


           
            WordApp.ActiveDocument.Save();
            
            

            
            
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
          
            //if (AnalysisPattern(true))
            //{
            //    /*Если шаблон вернёт значение true, то он корректен и мы можем приступить к замене слов(для замены создан специальный метод выше)*/
            //}
        }

        private void rtb_Tems_TextChanged(object sender, EventArgs e)
        {

        }

        private void rtb_ForExam_TextChanged(object sender, EventArgs e)
        {

        }

        public void fillingMainData() // загрузка информации из БД
        {
            // Обьявление массивов
            DA.CreateList();
            DA.initStruct();


            BD.Connect();
            BD.command.CommandText = "SELECT Дисциплины_профиля.Дисциплины, Дисциплины_профиля.Индекс,Дисциплины_профиля.Код_направления_подготовки, Дисциплины_профиля.Факт_по_зет, Дисциплины_профиля.По_плану, Дисциплины_профиля.Контакт_часы, Дисциплины_профиля.Аудиторные, Дисциплины_профиля.Самостоятельная_работа, Дисциплины_профиля.Контроль, Дисциплины_профиля.Элект_часы, Дисциплины_профиля.Интер_часы, Дисциплины_профиля.Закрепленная_кафедра, Дисциплины_профиля.Код_профиля, Дисциплины_профиля.Код FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код)=" + ID + "));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.Id_disp = Convert.ToInt32(BD.reader["Код"]);
                DA.Naim = BD.reader["Дисциплины"].ToString();
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
                ID_Napr = Convert.ToInt32(BD.reader["Код_направления_подготовки"]);
            }
            BD.reader.Close();
            // Запись направление подготовки и стандарт
            BD.command.CommandText = "SELECT Направление_подготовки.Код, Направление_подготовки.Направление_подготовки, Направление_подготовки.Станд FROM Направление_подготовки WHERE (((Направление_подготовки.Код)="+ID_Napr+"));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.Napr = BD.reader["Направление_подготовки"].ToString();
                DA.Standart = BD.reader["Станд"].ToString();
            }
            BD.reader.Close();
            // Запись профиль и год
            BD.command.CommandText = "SELECT Профиль.Название_профиля, Профиль.Год_профиля FROM Профиль WHERE (((Профиль.Код)="+ID_Prof+"));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.Profile = BD.reader["Название_профиля"].ToString();
                DA.Year = BD.reader["Год_профиля"].ToString();
            }
            BD.reader.Close();
            // Запись "Виды деятельности"
            BD.command.CommandText = "SELECT Виды_дейтельности.Список_дейтельности FROM Виды_дейтельности WHERE (((Виды_дейтельности.Код_направления_подготовки)=" + ID_Napr + "));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.MyList(BD.reader["Список_дейтельности"].ToString());
            }
            BD.reader.Close();
            // Запись часов по СЕМЕСТРАМ
            BD.command.CommandText = "SELECT Семестр.Номер_семестра, Семестр.ZET, Семестр.Итого, Семестр.Лек, Семестр.Лек_инт, Семестр.Лаб, Семестр.Лаб_инт, Семестр.ПР, Семестр.ПР_инт, Семестр.Элек, Семестр.СР, Семестр.Часы_конт, Семестр.Часы_конт_электр, Семестр.Экзамен, Семестр.Зачет, Семестр.Зачет_с_оценкой, Семестр.Курсовая FROM Семестр WHERE (((Семестр.Код_дисциплины)="+DA.Id_disp+"));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.LS = Convert.ToInt32(BD.reader["Номер_семестра"]);
                DA._ZET(DA.LS,Convert.ToInt32(BD.reader["ZET"]));
                DA._Itogo(DA.LS, Convert.ToInt32(BD.reader["Итого"]));
                DA._Lekc(DA.LS, Convert.ToInt32(BD.reader["Лек"]));
                DA._LekcInter(DA.LS, Convert.ToInt32(BD.reader["Лек_инт"]));
                DA._Lab(DA.LS, Convert.ToInt32(BD.reader["Лаб"]));
                DA._LabInter(DA.LS, Convert.ToInt32(BD.reader["Лаб_инт"]));
                DA._Practice(DA.LS, Convert.ToInt32(BD.reader["ПР"]));
                DA._PractInter(DA.LS, Convert.ToInt32(BD.reader["ПР_инт"]));
                DA._Elect(DA.LS, Convert.ToInt32(BD.reader["Элек"]));
                DA._SR1(DA.LS, Convert.ToInt32(BD.reader["СР"]));
                DA._HoursCont(DA.LS, Convert.ToInt32(BD.reader["Часы_конт"]));
                DA._HoursContElect(DA.LS, Convert.ToInt32(BD.reader["Часы_конт_электр"]));

                if (Convert.ToBoolean(BD.reader["Экзамен"]) == true)
                {
                    DA._Examen(DA.LS);
                }
                if (Convert.ToBoolean(BD.reader["Зачет"]) == true)
                {
                    DA._Zachet(DA.LS);
                }
                if (Convert.ToBoolean(BD.reader["Зачет_с_оценкой"]) == true)
                {
                    DA._Dif_Zachet(DA.LS);
                }

                DA.KR = Convert.ToInt32(BD.reader["Курсовая"]);

            }
            BD.reader.Close();
            // Запись компетенций дисцп
            BD.command.CommandText = "SELECT Компетенции_дисциплины.Код_дисциплины, Компетенции.Компетенция, Компетенции.Содержание FROM Компетенции INNER JOIN Компетенции_дисциплины ON Компетенции.Код = Компетенции_дисциплины.Код_компетенции WHERE (((Компетенции_дисциплины.Код_дисциплины)="+DA.Id_disp+"));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.AddCompet(BD.reader["Компетенция"].ToString());
                DA._InfoCompet(BD.reader["Содержание"].ToString());
            }
            BD.reader.Close();
            // Запись Дисцп ДО
            BD.command.CommandText = "SELECT Дисциплина_до.Дисциплина_до FROM Дисциплина_до WHERE (((Дисциплина_до.Код_дисциплины)=" + DA.Id_disp + "));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.AddPreDis(BD.reader["Дисциплина_до"].ToString());            
            }
            BD.reader.Close();
            // Запись Дисцп ПОСЛЕ
            BD.command.CommandText = "SELECT Дисциплина_после.Дисциплина_после FROM Дисциплина_после WHERE (((Дисциплина_после.Код_дисциплины)=" + DA.Id_disp + "));";
            BD.reader = BD.command.ExecuteReader();
            while (BD.reader.Read())
            {
                DA.AddAfterDis(BD.reader["Дисциплина_после"].ToString());
            }
            BD.reader.Close();
            

        }

    }
}
    
