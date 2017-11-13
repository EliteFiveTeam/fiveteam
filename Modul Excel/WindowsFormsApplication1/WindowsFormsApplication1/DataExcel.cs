using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using WindowsFormsApplication1;

namespace WindowsFormsApplication1
{
    
    struct Plan // Хранение данных из листа "Титул"
    {
        public int ID; // id профиля
        public string Napr; // Направление подготовки 
        public int LS; // Считывания номера семестра
        public int DistCount; // Количество дисциплин 
        public string Profile; // Профиль дисциплины 
        public string Standart; // Стандарт дисциплины
        public string Year; // Год 
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
        public string Naim { get; set; } // Наименование предмета 
        public string Index { get; set; } // Индекс предмета 
        public int Fact { get; set; } // Факт по ЗЕТ  
        public int AtPlan { get; set; } // По плану 
        public int ContactHours { get; set; } // Контакт часы 
        public int Aud { get; set; } // Ауд.
        public int SR { get; set; } // СР
        public int Contr { get; set; } // Контроль
        public int ElectHours { get; set; } // Элект часы
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
        public int[] ZET; // № семестр | ЗЕТ
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

        public void _Examen(int Var)
        {

            Examen[Var - 1] = true;
        } // add to array
        public void _Zachet(int Var)
        {

            Zachet[Var - 1] = true;
        } // add to array
        public void _Dif_Zachet(int Var)
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
}

