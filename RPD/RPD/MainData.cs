using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RPD
{
    class MainData
    {
        DataAccess DA;
        connection_to_bd BD = new connection_to_bd();
        Plan pl;
        public void fillingMainData()
        {
            int rota = DA.Id_disp;
            BD.Connect();
            BD.command.CommandText = "SELECT Дисциплины_профиля.Дисциплины, Дисциплины_профиля.Индекс, Дисциплины_профиля.Факт_по_зет, Дисциплины_профиля.По_плану, Дисциплины_профиля.Контакт_часы, Дисциплины_профиля.Аудиторные, Дисциплины_профиля.Самостоятельная_работа, Дисциплины_профиля.Контроль, Дисциплины_профиля.Элект_часы, Дисциплины_профиля.Интер_часы, Дисциплины_профиля.Закрепленная_кафедра, Дисциплины_профиля.Код_профиля FROM Дисциплины_профиля WHERE (((Дисциплины_профиля.Код)="+ DA.Id_disp+"));";
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
        
    }
}
