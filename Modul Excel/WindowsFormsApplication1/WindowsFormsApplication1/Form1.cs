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
using WindowsFormsApplication1;
using System.Threading;

namespace WindowsFormsApplication1
{
    
    struct Plan // Хранение данных из листа "Титул"
    {
        public string Napr; // Направление подготовки 
        public int LS; // Считывания номера семестра
        public int DistCount; // Количество дисциплин 
        public string Profile; // Профиль дисциплины 
        public string Standart; // Стандарт дисциплины
        public List<string> VidActive; // Список видов деятельности
        public List<string> OriginalCompet;
        public List<string> InfoCompet;



        public void CreateList() 
        {
            VidActive = new List<string>();
            OriginalCompet = new List<string>();
            InfoCompet = new List<string>();
        } // Объявления списка (ВД)
        public void MyList(string Val)
        {
            VidActive.Add(Val); 
        } // Add to List (ВД)
        public void _OriginalCompet(string Val)
        {
            OriginalCompet.Add(Val);
        } // Add to List (ВД)
        public void _InfoCompet(string Val)
        {
            InfoCompet.Add(Val);
        } // Add to List (ВД)
        
    }
    struct PlanTime // Хранение данных из листа "План"
    {
        public string Naim {get; set;} // Наименование предмета 
        public string Index {get; set;} // Индекс предмета 
        public int Fact {get; set;} // Факт по ЗЕТ  
        public int AtPlan {get; set;} // По плану 
        public int ContactHours {get; set;} // Контакт часы 
        public int Aud {get; set;} // Ауд.
        public int SR {get; set;} // СР
        public int Contr {get; set;} // Контроль
        public int ElectHours {get; set;} // Элект часы
        public int InterHours { get; set; } // Интер часы
        public int StartDis; // Начало дисциплины
        public int EndDis; // Конец дисциплины
        public string Kafedra; // Наименование кафедры
        public List<string> Compet; // Список компетенций
        public List<string> PreDis; // Дисц ДО
        public List<string> AfterDis; // Дисц ПОСЛЕ

        public void AddAfterDis(string Val)
        {
            AfterDis.Add(Val);
        } // Метод для добавления в список (Дисц ПОСЛЕ)
        public void AddPreDis(string Val)
        {
            PreDis.Add(Val);
        }   // Метод для добавления в список (Дисц ДО)
        public void AddCompet(string Val)
        {
            Compet.Add(Val);
        }   // Метод для добавления в список (Список компетенций)

            /* Хранение данных в семестрах */
            public int [] ZET; // № семестр | ЗЕТ
            public void _ZET(int Var, int Val)
            {
                ZET[Var - 1] = Val; 
            }
            public int[] Itogo; // № семестр | Итого
            public void _Itogo(int Var, int Val)
            {
                Itogo[Var - 1] = Val;
            }
            public int[] Lekc; // № семестр | Лекции
            public void _Lekc(int Var, int Val)
            {
                Lekc[Var - 1] = Val;
            }
            public int[] LekcInter; // № семестр | Интеракт лекции
            public void _LekcInter(int Var, int Val)
            {
                LekcInter[Var - 1] = Val;
            }
            public int[] Lab; // № семестр | Лаборот
            public void _Lab(int Var, int Val)
            {
                Lab[Var - 1] = Val;
            }
            public int[] LabInter; // № семестр | Интеракт лаборот
            public void _LabInter(int Var, int Val)
            {
                LabInter[Var - 1] = Val;
            }
            public int[] Practice; // № семестр | Практика
            public void _Practice(int Var, int Val)
            {
                Practice[Var - 1] = Val;
            }
            public int[] PractInter; // № семестр | Интеракт практика
            public void _PractInter(int Var, int Val)
            {
                PractInter[Var - 1] = Val;
            }
            public int[] Elect; // № семестр | Электив
            public void _Elect(int Var, int Val)
            {
                Elect[Var - 1] = Val;
            }
            public int[] _SR; // № семестр | СР
            public void _SR1(int Var, int Val)
            {
                _SR[Var - 1] = Val;
            }
            public int[] HoursCont; // № семестр | Контакт часы
            public void _HoursCont(int Var, int Val)
            {
                HoursCont[Var - 1] = Val;
            }
            public int[] HoursContElect; // № семестр | Элект контакт часы
            public void _HoursContElect(int Var, int Val)
            {
                HoursContElect[Var - 1] = Val;
            }
    
            /* ФОРМА КОНТРОЛЯ */
            public bool[] Examen; // Форм. контр | Экзамен
            public bool[] Zachet; // Форм. контр | Зачет
            public bool[] Dif_Zachet; // Форм. контр | Диф зачет
            public int KR; // Форм. контр | Курс раб

            public void _Examen (int Var)
            {
                
                Examen[Var - 1] = true;
            } // add to array
            public void _Zachet (int Var)
            {
                
                Zachet[Var - 1] = true;
            } // add to array
            public void _Dif_Zachet (int Var)
            {
                
                Dif_Zachet[Var - 1] = true;
            } // add to array

            public void initStruct()
            {
                Examen = new bool[10];
                Zachet = new bool[10];
                Dif_Zachet = new bool[10];
                ZET = new int[10];
                Itogo = new int[10];
                Lekc = new int[10];
                LekcInter = new int[10];
                Lab = new int[10];
                LabInter = new int[10];
                Practice = new int[10];
                PractInter = new int[10];
                Elect = new int[10];
                _SR = new int[10];
                HoursCont = new int[10];
                HoursContElect = new int[10];
                Compet = new List<string>();
                PreDis = new List<string>();
                AfterDis = new List<string>();
            } // Метод для объявление массивов (в структуре объявление методов недоступен)   
            
           
            }

   
    
    public partial class Form1 : Form
    {
        Plan PL; // Переменная структуры "Титул"
        PlanTime[] PLtime = new PlanTime[150]; // Переменная структуры "План"
        
        
        public Form1()
        {
            InitializeComponent();  
        }

        private void StartEndDist() // метод для определения начало и конца дисц
        {
            List<int> ListDisc = new List<int>();  // Список семестров дисц
            for (int j = 0; j <= PL.DistCount-1; j++)
            {
                for (int i = 0; i <= 9; i++)
                {
                    if (PLtime[j].Examen[i] == true || PLtime[j].Dif_Zachet[i] == true || PLtime[j].Zachet[i] == true)
                    {
                        int value = i+1;
                        ListDisc.Add(value); // Добавление в список
                    }
                }
                PLtime[j].StartDis = ListDisc.Min(); // Минимальное значение в списке (Начало дисц)
                PLtime[j].EndDis = ListDisc.Max(); // Максимальное значение в списке (Конец дисц)
                ListDisc.Clear(); // Очищаем список
            }
           

           
        }

        private void BeforeAndAfterDis () // Дисципл ДО и Дисциплин ПОСЛЕ
        {
            for (int i = 0; i <= PL.DistCount-1; i++) // первый список дисц
            {
                for (int j = 0; j <= PL.DistCount-1; j++) // второй список дисц
                {
                    bool flag = true;
                    if (i == j) // если одинаковые дисцип, переходим к другой
                    {
                        flag = false;
                    }
                    if (flag == true)
                    {
                        if (inlist(i, j) == true) // после проверки inlist, определяем дисц ДО и ПОСЛЕ
                        {
                            if (PLtime[i].StartDis > PLtime[j].EndDis)
                            {
                                PLtime[i].AddPreDis(PLtime[j].Naim); // доб. дисц ДО
                            }
                            if (PLtime[i].EndDis < PLtime[j].StartDis)
                            {
                                PLtime[i].AddAfterDis(PLtime[j].Naim); // доб. дисц ПОСЛЕ
                            }
                        }
                    }
                }
                   
            }
        }

        private bool inlist(int a, int b) // Проверка компетенций 
        {
            bool flag = false;
            for (int i = 0; i <= PLtime[a].Compet.Count - 1; i++)
            {
                for (int j = 0; j <= PLtime[b].Compet.Count-1; j++)
                {
                    if (PLtime[a].Compet[i] == PLtime[b].Compet[j])
                    {
                        flag = true;
                        return flag;
                    }
                }
                
            }
            return flag;
        }

        //private void Print() // вывод на экран для проверки 
        //{
        //    for (int i = 0; i <= PL.DistCount - 1; i++)
        //    {
        //        Action action5 = () => { textBox2.Text += PLtime[i].Naim + Environment.NewLine; }; Invoke(action5);
                
        //        for (int k = 0; k <= PLtime[i].PreDis.Count - 1; k++)
        //            {  Action action7 = () => { textBox2.Text += " | Дисциплина до |" + " " + PLtime[i].PreDis[k] + " "  + Environment.NewLine; }; Invoke(action7); }
        //        for (int j = 0; j <= PLtime[i].AfterDis.Count - 1; j++)
        //        {
        //            Action action8 = () => { textBox2.Text += " | Дисциплина после |" + " " + PLtime[i].AfterDis[j] + Environment.NewLine; }; Invoke(action8); 
        //        }  
                    
                
        //    }
        //}

        private void AnalysisDataExcel()
        {
            /* Открываем файл Excel и считываем информацию с первого листа "Титул" */

            string Fname;
            int NS;
            excel.Application ExcelApp = new excel.Application(); // создаем объект excel;
            ExcelApp.Visible = false; // показывает или скрывает файл Excel;
            Action action = () => { openFileDialog1.ShowDialog(); }; Invoke(action);  // Запуск главного потока 
            Fname = openFileDialog1.FileName;
            ExcelApp.Workbooks.Add(Fname); // загружаем в excel файл с рабочей книгой
            excel.Sheets excelsheets; // объявление переменных хранящих листы книги
            excel.Worksheet excelworksheet;
            excelsheets = ExcelApp.Worksheets;
            excelworksheet = (excel.Worksheet)excelsheets.get_Item("Титул"); // обращение к листу по названию
            string Open1Sheet = excelworksheet.Cells[11, 3].Text; // обращение к ячейкам книги
            for (int i = 20; i <= 50; i++)
            {
                string ST = excelworksheet.Cells[i, 13].Text;
                if (ST.IndexOf("стандарт") > 0)
                {
                string Open2Sheet = excelworksheet.Cells[i, 18].Text;
                PL.Standart = Open2Sheet.Trim();
                if (PL.Standart != null && PL.Standart != "")
                {
                    Action Progress = () => { richTextBox1.Text += "Стандарт профиля загружен" + Environment.NewLine; }; Invoke(Progress);
                }
                else
                {
                    Action Progress = () =>
                    {

                        richTextBox1.Text += "Стандарт профиля не найден!" + Environment.NewLine;
                        richTextBox1.Select(richTextBox1.Text.IndexOf("Стандарт профиля не найден!"), "Стандарт профиля не найден!".Length);
                        richTextBox1.SelectionColor = Color.Red;
                    }; Invoke(Progress);
                }
                }
                
            }
            PL.CreateList();

            NS = 3;
            int Flag = 1;
            for (int i = 1; i <= 5; i++)
            {
                string STR = excelworksheet.Cells[11, i].Text;
                if (STR.IndexOf("Направленность") > 0)
                {
                    NS = i;
                    Flag = 0;
                    break;
                }

            }
            if (Flag == 0)
            {
                Open1Sheet = excelworksheet.Cells[11, NS].Text;
            }
            else
            {
                for (int i = 1; i <= 5; i++)
                {
                    string STR = excelworksheet.Cells[18, i].Text;
                    if (STR.IndexOf("Направленность") > 0)
                    {
                        NS = i;
                        Flag = 0;
                        break;
                    }
                }
                if (Flag == 0)
                {
                    Open1Sheet = excelworksheet.Cells[18, NS].Text;
                }
                else
                {
                    //MessageBox.Show("Направление не найдено", "Ошибка!", MessageBoxButtons.OK);
                }
            }
            if (Flag == 0)
            {


                int i1 = Open1Sheet.IndexOf("Направленность");


                string STRNapr = Open1Sheet.Substring(22, i1 - 24);
                int i2 = Open1Sheet.IndexOf("\"");
                i1 = Open1Sheet.LastIndexOf("\"");
                string STRProf = Open1Sheet.Substring(i2 + 1, i1 - i2 - 1);
                ExcelApp.Visible = false;
                
                PL.Napr = STRNapr.Trim();
                if (PL.Napr != null && PL.Napr != "")
                {
                    Action Progress = () =>
                    {
                        richTextBox1.Text += "Направление профиля загружено" + Environment.NewLine;
                        richTextBox1.Select(richTextBox1.Text.IndexOf("Направление профиля загружено"), "Направление профиля загружено".Length);
                        richTextBox1.SelectionColor = Color.Black;
                    }; Invoke(Progress);
                }
                else 
                {
                    Action Progress = () => { richTextBox1.Text += "Направление профиля не найдено!"; }; Invoke(Progress);
                }
                PL.Profile = STRProf.Trim();

            }
            int J; // переменная номера столбца
            int SN = 1; // переменная номера ячейки со словом "Виды"
            int FlagVids = 1; // переменная признак нахождения "Виды деятельности"
            for (J = 2; J <= 3; J++)
            {
                for (int i = 15; i <= 40; i++)
                {
                    string STR = excelworksheet.Cells[i, J].Text;
                    if (STR.IndexOf("Виды") >= 0)
                    {

                        SN = i;
                        FlagVids = 0;
                        break;
                    }


                }
                if (FlagVids == 0)
                { break; }
            }
            if (FlagVids == 0)
            {
                for (int i = SN + 1; i <= SN + 10; i++)
                {
                    string STR = excelworksheet.Cells[i, J].Text;
                    string STR1 = excelworksheet.Cells[i, J - 1].Text;
                    if (STR1.IndexOf("+") >= 0)
                    {
                        PL.MyList(STR.Trim());


                        //textBox1.Text = textBox1.Text  +  Environment.NewLine  +  PL.VidActive[PL.VidActive.Count - 1];
                    }

                }
            }
            //textBox2.Text = textBox2.Text + PL.Napr + Environment.NewLine + PL.Profile + Environment.NewLine + PL.Standart + Environment.NewLine;



            /* Считывания информации с листа "Компетенции" */

            excelworksheet = (excel.Worksheet)excelsheets.get_Item("Компетенции");
            for (int a = 3; a <= 400; a++)
            {
                if (excelworksheet.Cells[a, 2].Text != "")
                {
                    string Compet = excelworksheet.Cells[a, 2].Text;
                    string Info = excelworksheet.Cells[a, 4].Text;
                    PL._OriginalCompet(Compet.Trim());
                    PL._InfoCompet(Info.Trim());
                }
            }
          
            /* Считывания информации с листа "План" */



            excelworksheet = (excel.Worksheet)excelsheets.get_Item("План");
            string PlanSheet1 = excelworksheet.Cells[6, 3].Text; // обращение к ячейкам книги "Список дисциплин"
            int ND = 0;

            /////////////////////////////////////////////////////////////////////
            for (int d = 6; d <= 140; d++)
            {
                string stroch = excelworksheet.Cells[d, 1].Text; // j - строчка ; i - столбец


                if (excelworksheet.Cells[d, 3].Font.Bold != true && stroch.IndexOf("+") >= 0 || excelworksheet.Cells[d, 3].Font.Bold != true && stroch.IndexOf("-") >= 0)
                {
                    PL.DistCount++;
                }
            }
            ////////////////////////////////////////////////////////////////

            for (int j = 6; j <= 150; j++)
            {
                string STR1 = excelworksheet.Cells[j, 1].Text; // j - строчка ; i - столбец


                if (excelworksheet.Cells[j, 3].Font.Bold != true && STR1.IndexOf("+") >= 0 || excelworksheet.Cells[j, 3].Font.Bold != true && STR1.IndexOf("-") >= 0)
                {
                    PLtime[ND].initStruct(); // объявление массива

                    for (int i = 4; i <= 125; i++)
                    {
                        string STR = excelworksheet.Cells[j, 3].Text;
                        string _index = excelworksheet.Cells[j, 2].Text;
                        PLtime[ND].Naim = STR; // наименование
                        PLtime[ND].Index = _index; // индекс дисциплины


                        string PlanSheet2 = excelworksheet.Cells[3, i].Text; // читаем название шапки
                        PlanSheet2 = PlanSheet2.Replace(" ", "");
                        PlanSheet2 = PlanSheet2.Replace(".", "");// удаляем все пробелы
                        PlanSheet2 = PlanSheet2.ToLower(); // переводим в нижний регистор
                        int Sem;

                        switch (PlanSheet2) // запись в структуру "Форма контроля"
                        {
                            case "экзамен":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    Sem = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                    if (Sem > 9)
                                    {
                                        string CheckSem = Sem.ToString();
                                        char[] NumSem = new char[CheckSem.Length];
                                        for (int z = 0; z < CheckSem.Length; z++)
                                        {
                                            NumSem[z] = CheckSem[z];
                                            string _CheckSem = NumSem[z].ToString();
                                            int N = Int32.Parse(_CheckSem);
                                            PLtime[ND]._Examen(N);


                                        }

                                    }
                                    else { PLtime[ND]._Examen(Sem); }


                                }
                                break;
                            case "зачет":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    Sem = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                    if (Sem > 9)
                                    {
                                        string CheckSem = Sem.ToString();
                                        char[] NumSem = new char[CheckSem.Length];
                                        for (int z = 0; z < CheckSem.Length; z++)
                                        {
                                            NumSem[z] = CheckSem[z];
                                            string _CheckSem = NumSem[z].ToString();
                                            int N = Int32.Parse(_CheckSem);
                                            PLtime[ND]._Examen(N);


                                        }

                                    }
                                    else { PLtime[ND]._Zachet(Sem); }
                                }
                                break;
                            case "зачетсоц":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    Sem = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                    if (Sem > 9)
                                    {
                                        string CheckSem = Sem.ToString();
                                        char[] NumSem = new char[CheckSem.Length];
                                        for (int z = 0; z < CheckSem.Length; z++)
                                        {
                                            NumSem[z] = CheckSem[z];
                                            string _CheckSem = NumSem[z].ToString();
                                            int N = Int32.Parse(_CheckSem);
                                            PLtime[ND]._Examen(N);


                                        }

                                    }
                                    else { PLtime[ND]._Dif_Zachet(Sem); }
                                }
                                break;
                            case "кр":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    Sem = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                    if (Sem > 9)
                                    {
                                        string CheckSem = Sem.ToString();
                                        char[] NumSem = new char[CheckSem.Length];
                                        for (int z = 0; z < CheckSem.Length; z++)
                                        {
                                            NumSem[z] = CheckSem[z];
                                            string _CheckSem = NumSem[z].ToString();
                                            int N = Int32.Parse(_CheckSem);
                                            PLtime[ND]._Examen(N);


                                        }

                                    }
                                    else { PLtime[ND].KR = Sem; }
                                }
                                break;
                        }

                        switch (PlanSheet2) // запись "Итого часов"
                        {
                            case "факт":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].Fact = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                            case "поплану":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].AtPlan = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                            case "контактчасы":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].ContactHours = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                            case "ауд":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].Aud = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                            case "ср":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].SR = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                            case "контроль":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].Contr = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                            case "интерчасы":
                                if (excelworksheet.Cells[j, i].Text != "")
                                {
                                    PLtime[ND].InterHours = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                }
                                break;
                        }
                        string NomerSemestra = excelworksheet.Cells[2, i].Text;

                        NomerSemestra.Trim();


                        if (NomerSemestra.IndexOf("Сем") >= 0)
                        {
                            string LastSymbol = NomerSemestra.Substring(NomerSemestra.Length - 1); // номер семестра в шапке
                            PL.LS = Int32.Parse(LastSymbol);
                        }


                        if (PL.LS > 0)
                        {


                            switch (PlanSheet2)
                            {
                                case "зет":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._ZET(PL.LS, Kek);
                                    }
                                    break;
                                case "итого":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._Itogo(PL.LS, Kek);
                                    }
                                    break;
                                case "лек":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._Lekc(PL.LS, Kek);
                                    }
                                    break;
                                case "лекинтер":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._LekcInter(PL.LS, Kek);
                                    }
                                    break;
                                case "лаб":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._Lab(PL.LS, Kek);
                                    }
                                    break;
                                case "лабинтер":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._LabInter(PL.LS, Kek);
                                    }
                                    break;
                                case "пр":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._Practice(PL.LS, Kek);
                                    }
                                    break;
                                case "принтер":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._PractInter(PL.LS, Kek);
                                    }
                                    break;
                                case "элект":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._Elect(PL.LS, Kek);
                                    }
                                    break;
                                case "ср":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._SR1(PL.LS, Kek);
                                    }
                                    break;
                                case "часыконт":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._HoursCont(PL.LS, Kek);
                                    }
                                    break;
                                case "часыконтэлектр":
                                    if (excelworksheet.Cells[j, i].Text != "")
                                    {
                                        int Kek = Int32.Parse(excelworksheet.Cells[j, i].Text);
                                        PLtime[ND]._HoursContElect(PL.LS, Kek);
                                    }
                                    break;

                            }


                        }

                        if (PlanSheet2.IndexOf("компетенции") >= 0) // Код компетенции
                        {
                            string Compet = excelworksheet.Cells[j, i].Text;
                            string[] DivComp = Compet.Split(new char[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string s in DivComp)
                            {
                                PLtime[ND].AddCompet(s);
                            }
                        }
                        if (PlanSheet2.LastIndexOf("наименование") >= 0) // Кафедра 
                        {
                            string KF = excelworksheet.Cells[j, i].Text;
                            PLtime[ND].Kafedra = KF;
                        }
                    }

                    ND++;

                }
                Action action2 = () => { progressBar1.Maximum = PL.DistCount; progressBar1.Value = ND; }; Invoke(action2);
               
            }
            ExcelApp.Quit();
            StartEndDist(); // определения начало и конца дисцип
            BeforeAndAfterDis(); // анализ дисц ПОСЛЕ и ДО
            //Print(); вывод дисциплин ДО и ПОСЛЕ
            PL.DistCount = 0;

            
            /* Заполнение инфррмации в БАЗУ ДАННЫХ */ 

            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + Application.StartupPath + "/baza_dan_proekt_kh.accdb");
            OleDbCommand command = new OleDbCommand("INSERT INTO Направление_подготовки (Индекс, Название, Станд) VALUES ('" + PL.Profile + "','" + PL.Napr + "','" + PL.Standart + "');", con);
            con.Open();
            OleDbDataReader reader;
            // запись в таблицу профиль
            command.CommandText = "INSERT INTO Профиль (Название_профиля) VALUES ('" + PL.Profile + "');";
            reader = command.ExecuteReader();
            reader.Close();
            // берем ID из профиля
            command.CommandText = "SELECT Профиль.Код FROM Профиль WHERE (((Профиль.[Название_профиля])='" + PL.Profile + "')); ";
            var code_profile = command.ExecuteScalar();
            reader.Close();
            command.CommandText = "INSERT INTO Направление_подготовки (Код_профиля, Направление_подготовки, Станд) VALUES ('" + code_profile + "','" + PL.Napr + "','" + PL.Standart + "');";
            reader = command.ExecuteReader();
            reader.Close();
            // получаем id из Направление_подготовки для записи в Дисциплины_профиля
            command.CommandText = "SELECT Направление_подготовки.Код FROM Направление_подготовки WHERE (((Направление_подготовки.[Направление_подготовки])='" + PL.Napr + "') AND ((Направление_подготовки.[Код_профиля])=" + code_profile + ") AND ((Направление_подготовки.[Станд])='" + PL.Standart + "')); ";
            var code = command.ExecuteScalar();
            reader.Close();
            //компетенции с листа компетенции 
            for (int y = 0; y <= PL.OriginalCompet.Count - 1; y++)
            {
                command.CommandText = "INSERT INTO Компетенции (Код_направления,Содержание,Компетенция) VALUES ('" + code + "','" + PL.InfoCompet[y] + "','" + PL.OriginalCompet[y] + "');";
                reader = command.ExecuteReader();
                reader.Close();
            }


            for (int i = 0; i <= PLtime.Length - 1; i++)
            {
                if (PLtime[i].Naim != null)
                {
                    command.CommandText = "INSERT INTO Дисциплины_профиля (Код_направления_подготовки,Дисциплины,Индекс,Факт_по_зет,По_плану,Контакт_часы,Аудиторные,Самостоятельная_работа,Контроль,Элект_часы,Интер_часы,Код_профиля) VALUES ('" + code + "','" + PLtime[i].Naim + "','" + PLtime[i].Index + "','" + PLtime[i].Fact + "','" + PLtime[i].AtPlan + "','" + PLtime[i].ContactHours + "','" + PLtime[i].Aud + "','" + PLtime[i].SR + "','" + PLtime[i].Contr + "','" + PLtime[i].ElectHours + "','" + PLtime[i].InterHours + "'," + code_profile + ");";
                    reader = command.ExecuteReader();
                    reader.Close();
                    //получаем ID дисциплины которую записали
                    command.CommandText = "SELECT Дисциплины_профиля.Код FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код_направления_подготовки)=" + code + ") AND ((Дисциплины_профиля.Дисциплины)='" + PLtime[i].Naim + "'));";
                    var code_distip = command.ExecuteScalar();
                    reader.Close();
                    //компетенции
                    //for (int y = 0; y <= PLtime[i].Compet.Count - 1 ; y++)
                    //{
                    //    // 
                    //    command.CommandText = "INSERT INTO Компетенция (Код_направления,Код_дисциплины,Компетенция) VALUES ('" + code + "','" + code_distip + "','" + PLtime[i].Compet[y] + "');";
                    //    reader = command.ExecuteReader();
                    //    reader.Close(); 
                    //}
                    //дисциплины до
                    for (int y1 = 0; y1 <= PLtime[i].PreDis.Count - 1; y1++)
                    {
                        // 
                        command.CommandText = "INSERT INTO Дисциплина_до (Код_дисциплины,Дисциплина_до) VALUES ('" + code_distip + "','" + PLtime[i].PreDis[y1] + "');";
                        reader = command.ExecuteReader();
                        reader.Close();
                    }
                    //дисциплины после
                    for (int y2 = 0; y2 <= PLtime[i].AfterDis.Count - 1; y2++)
                    {
                        // 
                        command.CommandText = "INSERT INTO Дисциплина_после (Код_дисциплины,Дисциплина_после) VALUES ('" + code_distip + "','" + PLtime[i].AfterDis[y2] + "');";
                        reader = command.ExecuteReader();
                        reader.Close();
                    }
                    int t; // прохождение по симестрам
                    for (t = 0; t <= 9; t++)
                    {
                        if (PLtime[i].Dif_Zachet[t] == true || PLtime[i].Zachet[t] == true || PLtime[i].Examen[t] == true)
                        {
                            int nomer_sem = t + 1;
                            command.CommandText = "INSERT INTO Семестр (Номер_семестра,ZET,Лек,Лек_инт,ПР,Лаб,Лаб_инт,ПР_инт,Элек,СР,Часы_конт,Часы_конт_электр,Курсовая,Итого,Код_дисциплины,Экзамен,Зачет,Зачет_с_оценкой) VALUES ('" + nomer_sem + "','" + PLtime[i].ZET[t] + "','" + PLtime[i].Lekc[t] + "','" + PLtime[i].LekcInter[t] + "','" + PLtime[i].Practice[t] + "','" + PLtime[i].Lab[t] + "','" + PLtime[i].LabInter[t] + "','" + PLtime[i].PractInter[t] + "','" + PLtime[i].Elect[t] + "','" + PLtime[i]._SR[t] + "','" + PLtime[i].HoursCont[t] + "','" + PLtime[i].HoursContElect[t] + "','" + PLtime[i].KR + "','" + PLtime[i].Itogo[t] + "','" + code_distip + "'," + PLtime[i].Examen[t] + "," + PLtime[i].Zachet[t] + "," + PLtime[i].Dif_Zachet[t] + ");";
                            reader = command.ExecuteReader();
                            reader.Close();
                        }
                    }
                }
            }

            //           Action action3 = () => { textBox3.Text += PL.VidActive[1]; }; Invoke(action3);
            for (int i = 0; i <= PL.VidActive.Count - 1; i++)
            {
                command.CommandText = "INSERT INTO Виды_дейтельности (Список_дейтельности,Код_направления_подготовки) VALUES ('" + PL.VidActive[i] + "','" + code + "');";
                reader = command.ExecuteReader();
                reader.Close();

            }
            Action action1 = () => { MessageBox.Show("Complete"); }; Invoke(action1); // Запуск главного потока

        }

       

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Thread theard = new Thread(AnalysisDataExcel); //второй поток для 
            theard.Start();
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

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
        }
    }  
}
