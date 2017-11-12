using System;
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

namespace WindowsFormsApplication1
{
    public struct Tema
    {
        public string Name;// ' Название темы
        public string Text; // ' Содержание темы
        public string Rez;// As String ' Результат темы
        public string Comp;// As String ' Компетенции, развиваемые темой
        public string FormZ;// As String ' Формы занятий
        public int N_Sem; // As Integer  ' Номер семестра
    }

    public struct Discipline
    {
        public string Index;// 'Индекс (номер дисциплины в плане)
        public string Name;// 'Наименование
        public string Exam;// 'Экзамены
        public string Zach;// 'Зачеты
        public string Zach_E;// 'Зачеты с оценкой
        public string Section;// 'Раздел плана
        public string Curs_R;// ' Курсовые работы
        public string Cafedra;// 'Закрепленная кафедра
        public byte First_Sem;// 'Первый семестр изучения дисциплины
        public byte Last_Sem;//'Последний семестр изучения дисциплины
        public string List_Comp;// 'Список компетенций
    }
    
    public partial class Form1 : Form
    {
        
        public static bool btn1; 
        Tema tems;
        Discipline dis;
        public Dis D = new Dis(); /*Класс*/
        char[] MyChar = { '\f', '\n', '\r', '\t', '\v', '\0', ' ', '2', '3', '.', ')', ';' };
        int CountKFind;  //' счетчик найденных фрагментов, n-сколько надо отсчитать нахождений до нужного
        word.Application WordApp;
        private int sec; // переменная, содержащая значение времени
        public Form1()
        {
            InitializeComponent();
            sec = 0; // начальное значение времени
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
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void AnalysisOldProgramm()
        {
            string Filename_;
            WordApp = new word.Application(); // создаем объект word;
            WordApp.Visible = true; // показывает или скрывает файл word;
            Action action = () => { openFileDialog1.ShowDialog(); }; Invoke(action);
            //openFileDialog1.ShowDialog();
            openFileDialog1.Filter = "Файлы Word(*.doc)|*.doc|Word(*.docx)|*.docx"; // фильтрует, оставляя только ворд файлы
            Filename_ = openFileDialog1.FileName;
            WordApp.Documents.Add(Filename_);// загружаем в word файл с рабочей книгой 
            button1.Enabled = false; // отключает кнопку на 2 секунды, согласно таймеру
            timer1.Start();

            Action action1 = () => { MessageBox.Show("Complete"); }; Invoke(action1); // Запуск главного потока 

            SearchText(textBox2.Text, textBox4.Text, CountKFind);
            int N = 0;
            int i = 0;
            //int j = 0;
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
                if (m21 == 0)
                {
                    richTextBox2.Text = "Основная литература не найдена!";
                }
                else
                {
                    object Start = r.ListParagraphs[1].Range.Start;
                    object End = r.ListParagraphs[m21].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    richTextBox4.Paste();
                    for (int y = 1; y <= r.ListParagraphs.Count; y++)
                    {
                        string dfs = r.ListParagraphs[y].Range.Text;
                        D.MyListAdd(dfs, false);
                    }
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
                if (m12 == 0)
                {
                    richTextBox2.Text = "Дополнительная литература не найдена!";
                }
                else
                {
                    object Start = r.ListParagraphs[1].Range.Start;
                    object End = r.ListParagraphs[m12].Range.End;
                    word.Range myRange = WordApp.ActiveDocument.Range(Start, End);
                    myRange.Copy();
                    richTextBox5.Paste();
                    for (int x = 1; x <= r.ListParagraphs.Count; x++)
                    {
                        string dsf = r.ListParagraphs[x].Range.Text;
                        D.MyListAdd(dsf, true);
                    }
                }
            } // поиск закончился, литература записана в массив


            //находим цели дисциплины
            if (ss == "") //' Если цели не попали в оглавление
            {
                ss = SearchText("явля?????", "Учебные задачи дисциплины", 2); // искомый текст после оглавления
            }

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
            }

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
            byte razd = 1;  //'номер раздела
            int CountTems = 0;
            for (i = 2; i <= WordApp.ActiveDocument.Tables[2].Rows.Count; i++)
            {
                if (WordApp.ActiveDocument.Tables[2].Rows[i].Cells.Count >= 5)
                {
                    D.tems[i - 2].Name = WordApp.ActiveDocument.Tables[2].Cell(i, 2).Range.Text;
                    D.tems[i - 2].Text = WordApp.ActiveDocument.Tables[2].Cell(i, 3).Range.Text;
                    D.tems[i - 2].Rez = WordApp.ActiveDocument.Tables[2].Cell(i, 5).Range.Text;
                    D.tems[i - 2].FormZ = WordApp.ActiveDocument.Tables[2].Cell(i, 6).Range.Text;
                    CountTems++;

                    //richTextBox2.Text = richTextBox2.Text + D.tems[i - 2].Name + D.tems[i - 2].Text + D.tems[i - 2].Rez + D.tems[i - 2].FormZ;
                    //Clipboard.SetText(richTextBox2.Text + D.tems[i - 2].Name + D.tems[i - 2].Text + D.tems[i - 2].Rez + D.tems[i - 2].FormZ);


                }
                else
                {
                    if (i != 2)
                    {
                        razd += razd;  //' счетчик разделов срабатывает если их больше одного
                    }
                }
            }
            D.Nt = CountTems; //Записали количество тем в дисциплине

            Clipboard.Clear();

            // считываются темы и их литература, вопросы для самопроверки
            ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 2);
            //Clipboard.SetText(ss);

            int n1, n2, n3, n4;
            n1 = ss.IndexOf("Тема");
            n2 = ss.IndexOf("Литература");
            n3 = ss.IndexOf("Вопросы для");
            n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
            if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10)  //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
            {
                richTextBox3.Text = "";
                richTextBox3.Paste();
                if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) // ' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                {
                    richTextBox3.Text = "";
                    richTextBox3.Paste();
                    if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                    {
                        richTextBox3.Text = "";
                        richTextBox3.Paste();
                        if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                        {
                            richTextBox3.Text = "";
                            richTextBox3.Paste();
                            if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > 0 && n4 < 10) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                            {
                                richTextBox3.Text = "";
                                richTextBox3.Paste();
                                if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > n2) //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                                {
                                    richTextBox3.Text = "";
                                    richTextBox3.Paste();
                                    if ((n1 > 0 && n1 < 100) && (n2 > n1 && n2 < 300) && n3 > n2)  //' попытка определить что искомый фрагмент найден (он начинается со слов "Тема"
                                    {
                                        richTextBox3.Text = "";
                                        richTextBox3.Paste();
                                    }
                                    else
                                    {
                                        richTextBox2.Text = "Ошибка: Перечень УМО не найден";
                                    }
                                }
                                else
                                {
                                    ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ", 1);
                                    n1 = ss.IndexOf("Тема");
                                    n2 = ss.IndexOf("Литература");
                                    n3 = ss.IndexOf("Вопросы для");
                                }
                            }
                            else //' Это для РП образца 2015г.
                            {
                                ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ", 2);
                                n1 = ss.IndexOf("Тема");
                                n2 = ss.IndexOf("Литература");
                                n3 = ss.IndexOf("Вопросы для");
                            }
                        }
                        else
                        {
                            ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Рекомендуемые обучающие", 1);
                            n1 = ss.IndexOf("Тема");
                            n2 = ss.IndexOf("Литература");
                            n3 = ss.IndexOf("Вопросы для");
                            n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
                        }
                    }
                    else
                    {
                        ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Рекомендуемые обучающие", 2);
                        n1 = ss.IndexOf("Тема");
                        n2 = ss.IndexOf("Литература");
                        n3 = ss.IndexOf("Вопросы для");
                        n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
                    }
                }
                else //' это если в конце файла есть еще раз этот раздел, то надо искать третье вхождение
                {
                    ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 3);
                    n1 = ss.IndexOf("Тема");
                    n2 = ss.IndexOf("Литература");
                    n3 = ss.IndexOf("Вопросы для");
                    n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
                }
            }
            else //' это если в содержании нет этого раздела, а в тексте есть
            {
                ss = SearchText("Перечень учебно-методического обеспечения для самостоятельной работы обучающихся по дисциплине", "Материально-техническое обеспечение дисциплины", 1);
                n1 = ss.IndexOf("Тема");
                n2 = ss.IndexOf("Литература");
                n3 = ss.IndexOf("Вопросы для");
                n4 = ss.IndexOf("III. ОБРАЗОВАТЕЛЬНЫЕ ТЕХНОЛОГИИ");
            }


            Clipboard.Clear();
            // поиск вопросов к экзаменам
            int k; // ' метки для найденных символов
            k = richTextBox2.Find("Тема 1");
            richTextBox2.SelectAll();
            if (k > 0)
            {
                richTextBox2.SelectedText.Remove(0, k - 1);
            }
            ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ ДЛЯ ОБУЧАЮЩИХСЯ", 2);
            //Clipboard.SetText(ss);
            if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
            {
                richTextBox3.Text = "";
                richTextBox3.Paste();
                if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
                {
                    richTextBox3.Text = "";
                    richTextBox3.Paste();

                    if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                    {
                        richTextBox3.Text = "";
                        richTextBox3.Paste();
                        if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                        {
                            richTextBox3.Text = "";
                            richTextBox3.Paste();
                            if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                            {
                                richTextBox3.Text = "";
                                richTextBox3.Paste();
                                if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
                                {
                                    richTextBox3.Text = "";
                                    richTextBox3.Paste();
                                    if (ss.Length > 500) //  ' Вставка в RTB3 заданий и вопросов к экзамену
                                    {
                                        richTextBox3.Text = "";
                                        richTextBox3.Paste();
                                        if (ss.Length > 500) // ' Вставка в RTB3 заданий и вопросов к экзамену
                                        {
                                            richTextBox3.Text = "";
                                            richTextBox3.Paste();
                                        }
                                        else
                                        {
                                            richTextBox3.Text = "Ошибка:Перечень Заданий не найден!";
                                        }
                                    }
                                    else
                                    {
                                        ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "Тематический план", 1);
                                    }
                                }
                                else
                                {
                                    ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "Тематический план", 2);
                                }
                            }
                            else
                            {
                                ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "ТЕМАТИЧЕСКИЙ ПЛАН", 1);
                            }
                        }
                        else //' это для РПД образца 2015г.
                        {
                            ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "ТЕМАТИЧЕСКИЙ ПЛАН", 2);
                        }
                    }
                    else
                    {
                        ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "МЕТОДИЧЕСКИЕ УКАЗАНИЯ", 1);
                    }
                }
                else
                {
                    ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "МЕТОДИЧЕСКИЕ УКАЗАНИЯ", 2);
                }
            }
            else
            {
                ss = SearchText("характеризующих этапы формирования компетенций в процессе освоения образовательной программы", "VII.  МЕТОДИЧЕСКИЕ УКАЗАНИЯ ДЛЯ ОБУЧАЮЩИХСЯ", 1);
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
                    richTextBox1.Paste();
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
                    richTextBox1.Paste();
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
                    richTextBox1.Paste();
                }
            }
            
        }
        private void button1_Click(object sender, EventArgs e) // Метод, открывающий ворд документ
        {
           AnalysisOldProgramm();
           WordApp.Quit();
        } 
            
            
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e) // Кнопка создания новой рабочей программы
        {
        } 

        private void button3_Click(object sender, EventArgs e)
        { 
        } // бесполезная кнопка.

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e) // Основная кнопка поиска
        {
        }

        public void Process()
        {
            Application.Run();
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
                if (sec == 2)
            {
                sec = 0;
                button1.Enabled = true;
                timer1.Stop();
            }
            else
                sec++;
        
        } 
   }
}

