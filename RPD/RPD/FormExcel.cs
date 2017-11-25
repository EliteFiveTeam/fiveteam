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
    public partial class FormExcel : Form
    {
        Plan PL; // Переменная структуры "Титул"
        PlanTime[] PLtime = new PlanTime[150]; // Переменная структуры "План"
        connection_to_bd BD = new connection_to_bd();
        
        public FormExcel()
        {
            InitializeComponent();
        }

        private void FormExcel_Load(object sender, EventArgs e)
        {

        }
        public void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }

        private void StartEndDist() // метод для определения начало и конца дисц
        {
            List<int> ListDisc = new List<int>();  // Список семестров дисц
            for (int j = 0; j <= PL.DistCount - 1; j++)
            {
                for (int i = 0; i <= 9; i++)
                {
                    if (PLtime[j].Examen[i] == true || PLtime[j].Dif_Zachet[i] == true || PLtime[j].Zachet[i] == true)
                    {
                        int value = i + 1;
                        ListDisc.Add(value); // Добавление в список
                    }
                }
                PLtime[j].StartDis = ListDisc.Min(); // Минимальное значение в списке (Начало дисц)
                PLtime[j].EndDis = ListDisc.Max(); // Максимальное значение в списке (Конец дисц)
                ListDisc.Clear(); // Очищаем список
            }



        }

        private void BeforeAndAfterDis() // Дисципл ДО и Дисциплин ПОСЛЕ
        {
            for (int i = 0; i <= PL.DistCount - 1; i++) // первый список дисц
            {
                for (int j = 0; j <= PL.DistCount - 1; j++) // второй список дисц
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
                for (int j = 0; j <= PLtime[b].Compet.Count - 1; j++)
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


        private void AnalysisDataExcel()
        {
            
                /* Открываем файл Excel и считываем информацию с первого листа "Титул" */
                bool Prov_on_double = true;
                string Fname;
                int NS;
                excel.Application ExcelApp = new excel.Application(); // создаем объект excel;
                ExcelApp.Visible = false; // показывает или скрывает файл Excel;
                openFileExcel.Filter = "Файлы Excel(*.xls)|*.xls|Excel(*.xls)|*.xls"; 
                Action action = () => { openFileExcel.ShowDialog(); }; Invoke(action);  // Запуск главного потока 
                Fname = openFileExcel.FileName;
                ExcelApp.Workbooks.Add(Fname); // загружаем в excel файл с рабочей книгой
                Action action1 = () => { bt_selct_excel.Enabled = false; }; Invoke(action1);
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
                            Action Progress = () => { RTB_ExcelLog.AppendText("Стандарт профиля загружен\n", Color.Green); }; Invoke(Progress);

                        }
                        else
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Стандарт профиля не найден!\n", Color.Red); }; Invoke(Progress);
                        }
                    }
                    string YR = excelworksheet.Cells[i, 13].Text;
                    if (YR.IndexOf("подготовки") > 0)
                    {
                        string Open2Sheet = excelworksheet.Cells[i, 18].Text;
                        PL.Year = Open2Sheet.Trim();
                        if (PL.Year == "")
                        {
                            string repeat = excelworksheet.Cells[i, 20].Text;
                            PL.Year = repeat;
                            PL.Year = repeat.Trim();
                        }
                        if (PL.Year != null && PL.Year != "")
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Год профиля загружен\n", Color.Green); }; Invoke(Progress);
                        }
                        else
                        {
                            Action Progress = () =>
                            { RTB_ExcelLog.AppendText("Год профиля не найден!\n", Color.Red); }; Invoke(Progress);
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
                        Action Progress = () => { RTB_ExcelLog.AppendText("Направление профиля загружено\n", Color.Green); }; Invoke(Progress);
                    }
                    else
                    {
                        Action Progress = () => { RTB_ExcelLog.AppendText("Направление профиля не найдено!\n", Color.Red); }; Invoke(Progress);
                    }
                    PL.Profile = STRProf.Trim();
                    if (PL.Profile != null && PL.Profile != "")
                    {
                        Action Progress = () => { RTB_ExcelLog.AppendText("Профиль загружен\n", Color.Green); }; Invoke(Progress);
                    }
                    else
                    {
                        Action Progress = () => { RTB_ExcelLog.AppendText("Профиль не найден!\n", Color.Red); }; Invoke(Progress);
                    }
                    //BD.Connect(); //проверка на дюблирование
                    //BD.command.CommandText = "SELECT Профиль.Код FROM Профиль WHERE (((Профиль.Название_профиля)='"+PL.Profile+"') AND ((Профиль.Год_профиля)="+PL.Year+"));";
                    //int jdfbier = BD.command.ExecuteNonQuery();
                    //if (jdfbier!=0)
                    //{
                    //    Action Progress = () => { MessageBox.Show("данные есть"); }; Invoke(Progress);
                    //    Prov_on_double = false;
                    //}
                    //else
                    //{
                    //    Action Progress = () => { MessageBox.Show("данных нет"); }; Invoke(Progress);
                    //}

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
                            Action Progress = () => { RTB_ExcelLog.AppendText("Виды деятельности загружены\n", Color.Green); }; Invoke(Progress);
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
                        }

                    }
                }




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
                if (PL.OriginalCompet.Count != 0)
                {
                    Action Progress = () => { RTB_ExcelLog.AppendText("Информация о компетенциях загружена\n", Color.Green); }; Invoke(Progress);
                }
                else
                {
                    Action Progress = () => { RTB_ExcelLog.AppendText("Информация о компетенциях не найдена!\n", Color.Red); }; Invoke(Progress);
                }


                /* Считывания информации с листа "План" */

                excelworksheet = (excel.Worksheet)excelsheets.get_Item("План");
                string PlanSheet1 = excelworksheet.Cells[6, 3].Text; // обращение к ячейкам книги "Список дисциплин"
                int ND = 0;
                PL.DistCount = 0;
                /////////////////////////////////////////////////////////////////////
                for (int d = 6; d <= 150; d++)
                {
                    string stroch = excelworksheet.Cells[d, 1].Text; // j - строчка ; i - столбец


                    if (excelworksheet.Cells[d, 3].Font.Bold != true && stroch.IndexOf("+") >= 0 || excelworksheet.Cells[d, 3].Font.Bold != true && stroch.IndexOf("-") >= 0)
                    {
                        PL.DistCount++; // количество дисциплин
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

                        // Процесс загрузки
                        if (PL.DistCount > 0)
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Загрузка дисциплин прогресс " + PLtime[ND].Naim.ToString() + " загружено\n", Color.Green); }; Invoke(Progress);
                        }
                        /* Обработка возможных ошибок*/
                        if (PLtime[ND].Naim == "")
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Наименование дисциплины не найдено!\n", Color.Red); }; Invoke(Progress);
                        }
                        if (PLtime[ND].Index == "")
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Индекс не найден!\n", Color.Red); }; Invoke(Progress);
                        }
                        if (PLtime[ND].Compet.Count == 0)
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Компетенции не найдены!\n", Color.Red); }; Invoke(Progress);
                        }
                        if (PLtime[ND].Kafedra == "")
                        {
                            Action Progress = () => { RTB_ExcelLog.AppendText("Кафедра не найдена!\n", Color.Red); }; Invoke(Progress);
                        }


                        // счетчик дисциплин
                        ND++;

                    }
                    Action action2 = () => { PB_Excel.Maximum = PL.DistCount; PB_Excel.Value = ND; }; Invoke(action2);

                }

                // Если дисциплины не найдены, появляется информация об ошибке
                if (PL.DistCount == 0)
                {
                    Action Progress = () => { RTB_ExcelLog.AppendText("Дисциплины не найдены!\n", Color.Red); }; Invoke(Progress);
                }


                StartEndDist(); // определения начало и конца дисцип
                BeforeAndAfterDis(); // анализ дисц ПОСЛЕ и ДО
                PL.DistCount = 0;
                CloseProcess();

                /* Заполнение инфррмации в БАЗУ ДАННЫХ */
                Action AddBD = () => { RTB_ExcelLog.AppendText("Заполняем Базу Данных \n", Color.Blue); }; Invoke(AddBD);
                OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + Application.StartupPath + "/baza_dan_proekt_kh.accdb");
                OleDbCommand command = new OleDbCommand("INSERT INTO Направление_подготовки (Индекс, Название, Станд) VALUES ('" + PL.Profile + "','" + PL.Napr + "','" + PL.Standart + "');", con);
                con.Open();
                OleDbDataReader reader;
                // запись в таблицу профиль
                command.CommandText = "INSERT INTO Профиль (Название_профиля,Год_профиля) VALUES ('" + PL.Profile + "','" + PL.Year + "');";
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
                        command.CommandText = "INSERT INTO Дисциплины_профиля (Код_направления_подготовки,Дисциплины,Индекс,Факт_по_зет,По_плану,Контакт_часы,Аудиторные,Самостоятельная_работа,Контроль,Элект_часы,Интер_часы,Код_профиля,Закрепленная_кафедра) VALUES ('" + code + "','" + PLtime[i].Naim + "','" + PLtime[i].Index + "','" + PLtime[i].Fact + "','" + PLtime[i].AtPlan + "','" + PLtime[i].ContactHours + "','" + PLtime[i].Aud + "','" + PLtime[i].SR + "','" + PLtime[i].Contr + "','" + PLtime[i].ElectHours + "','" + PLtime[i].InterHours + "'," + code_profile + ",'" + PLtime[i].Kafedra + "');";
                        reader = command.ExecuteReader();
                        reader.Close();
                        //получаем ID дисциплины которую записали
                        command.CommandText = "SELECT Дисциплины_профиля.Код FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код_направления_подготовки)=" + code + ") AND ((Дисциплины_профиля.Дисциплины)='" + PLtime[i].Naim + "'));";
                        var code_distip = command.ExecuteScalar();
                        reader.Close();
                        //  подготовка к записи в таблицу компетенции_дисциплины 
                        for (int y = 0; y <= PLtime[i].Compet.Count - 1; y++)
                        {  //берем ID из таблицы компетенции для помещения в таблицу  компетенции_дисциплины
                            command.CommandText = "SELECT Компетенции.Код, Компетенции.Компетенция FROM Компетенции WHERE (((Компетенции.Компетенция)='" + PLtime[i].Compet[y] + "')); ";
                            var code_komped = command.ExecuteScalar();
                            reader.Close();
                            command.CommandText = "INSERT INTO Компетенции_дисциплины (Код_компетенции,Код_дисциплины) VALUES (" + code_komped + "," + code_distip + ");";
                            reader = command.ExecuteReader();
                            reader.Close();

                        }
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


                for (int i = 0; i <= PL.VidActive.Count - 1; i++)
                {
                    command.CommandText = "INSERT INTO Виды_дейтельности (Список_дейтельности,Код_направления_подготовки) VALUES ('" + PL.VidActive[i] + "','" + code + "');";
                    reader = command.ExecuteReader();
                    reader.Close();

                }
                Action CompleteBD = () => { RTB_ExcelLog.AppendText("Информация в Базу Данных загружена  \n", Color.Green); }; Invoke(CompleteBD);

                Action BT = () => { bt_selct_excel.Enabled = true; }; Invoke(BT);
                DataBase();
            
        }

        public void DataBase() // Добавление в ListBox1
        {
            FormMain main = this.Owner as FormMain;
            if (main != null)
            {

                Action BT = () =>
                {
                    main.lst_prof.Items.Clear();
                    BD.Connect();
                    BD.command.CommandText = "SELECT * FROM Профиль ;";
                    BD.reader = BD.command.ExecuteReader();
                    while (BD.reader.Read())
                    {
                        main.lst_prof.Items.Add(BD.reader.GetString(1) + " " + BD.reader.GetString(2));

                    }
                }; Invoke(BT);
            }
        }
        


        private void bt_selct_excel_Click(object sender, EventArgs e)
        {
            Thread theard = new Thread(AnalysisDataExcel); //второй поток для анализа EXCEL
            theard.Start();
        }

        private void RTB_ExcelLog_TextChanged(object sender, EventArgs e)
        {

        }

        private void PB_Excel_Click(object sender, EventArgs e)
        {

        }

    }
}
