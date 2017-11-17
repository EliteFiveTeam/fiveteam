using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RPD
{
     public class Dis
        {
            public string Name;     //'Название
            public string Num;       //'Номер в уч плане
            public string Направл; //'Направление
            public string Проф;     //'Профиль
            public string Section;  //'Раздел плана (базовый или вариативный)
            public string Exam;     //'Экзамены/Зач/Зач с оценкой
            public string Curs_R;    //' Курсовые работы
            public string Period;    //' Семестры, в которых изучается дисциплина
            public int Zet_Exp;  //' ЗЕТ Экспертное
            public int Zet_Fact; //'ЗЕТ Факт 
            public int H_Zet;   //'Итого часов по ЗЕТ
            public int H_Plan;  //'Итого часов по плану
            public int H_Contact; //'Контакт часы 
            public int H_Aud;   //'Аудиторные часы 
            public int H_CPC;    //'Часов СРС
            public int H_Control; //'Часов Контроль
            public int H_Electr; //'Электронные часы 
            public string Cafedra;  //'Закрепленная кафедра
            public string List_Activities;  //'Список видов занятий

            public int H_Lec;// ' Часов лекций
            public int H_Lec_InterA;// ' Часов лекций интерактивн.
            public int H_Prac;// 'часов Практики
            public int H_Prac_InterA;// 'часов Практики интерактивн.
            public int H_Lab;//'часов Лабоаторн


            public string Cel = "";// As String 'Цель дисциплины
            public string Tasks;// As String ' Задачи
            public string Dis_before;// As String ' предшествующие дисциплины
            public string Dis_after;// As String ' последующие дисциплины
            public Tema[] tems = new Tema[30];
            //Tema tems;// As Tema ' Темы дисциплины
            public int Nc;// ' Колво компетенций
            public int Nt;// ' Колво тем в дисциплине
            public string Zn_before;// 'Знания до
            public string Zn_after;// 'Знания после
            public string Um_before;// 'Умения до
            public string Um_after;//'Знания после
            public string Vl_before;// 'Знания до
            public string Vl_after;// 'Знания после
            public string FGOS;// ' номер и дата утверждения приказа о введении ФГОС
            public byte flCurs;// As Byte ' флаг того, что есть курсовые
            public byte flExam;// As Byte ' флаг того, что есть экзамен

            public List<string> LiteraBasic = new List<string>();
            public List<string> LiteraAdditional = new List<string>();

            public List<string> ForExam = new List<string>();// массив для вопросов к экзамену
            public void CreateForExam()
            {
                ForExam = new List<string>();

            }
            public void MyForExamAdd(string Valu1)
            {
                ForExam.Add(Valu1);
            }


            public void CreateLitera()
            {
                LiteraBasic = new List<string>();
                LiteraAdditional = new List<string>();
            }
            public void MyListAdd(string Val, bool direct)
            {
                if (direct == false)
                {
                    LiteraBasic.Add(Val);
                }
                else
                {
                    LiteraAdditional.Add(Val);
                }
            }
        }
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
    
}
