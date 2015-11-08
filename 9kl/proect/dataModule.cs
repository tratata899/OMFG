using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;

namespace proect
{
    class dataModule
    {
        protected OleDbConnection gConnection;
        protected string gStr = 
            " SELECT "+'\n'+
            "   Головна.[№ п/п], "+'\n'+
            "   Головна.[№ реєстрації], " + '\n' +
            "   Головна.ПІБ, " + '\n' +
            "   Спеціальність.Спеціальність as [Обрана Спеціальність], " + '\n' +
            "   Головна.[Шкільна оцінка по алгебрі], " + '\n' +
            "   Головна.[Шкільна оцінка по геометрії], " + '\n' +
            "   Головна.[Шкільна оцінка по українській мові], " + '\n' +
            "   Головна.[Шкільний середній бал], " + '\n' +
            "   Головна.[Бал за результатами вступних екзаменів: українська мова], " + '\n' +
            "   Головна.[Бал за результатами вступних екзаменів: математика], " + '\n' +
            "   Головна.[Бал за результатами вступних екзаменів: рисунок], " + '\n' +
            "   Головна.[Бал за результатами вступних екзаменів: композиція], " + '\n' +
            "   Головна.[Бал за підготовчі курси], " + '\n' +
            "   Головна.[Сума балів (прохідний бал)], " + '\n' +
            "   Головна.Пільги, Головна.Відзнака, " + '\n' +
            "   [Підготовчі курси].Тривалість as [Проходив підготовчі курси], " + '\n' +
            "   Головна.[Документ про освіту від 7 до 12 балів], " + '\n' +
            "   Головна.Немісцевий, Головна.[Потребує гуртожитку], " + '\n' +
            "   Стать.Стать as Стать, Головна.[Неповна сім'я], " + '\n' +
            "   [Навчальний заклад].[Навчальний заклад] as Освіта, " + '\n' +
            "   Головна.[Документ про освіту: серія], " + '\n' +
            "   Головна.[Документ про освіту: номер], " + '\n' +
            "   Спеціальність_1.Спеціальність as [Друга спеціальність], " + '\n' +
            "   Головна.Зараховано " + '\n' +
            " FROM " + '\n' +
            "   Стать " + '\n' +
            "   Спеціальність  " + '\n' +
            "   [Підготовчі курси]  " + '\n' +
            "   [Навчальний заклад]  " + '\n' +
            "   Спеціальність AS Спеціальність_1  " + '\n' +
            "   Головна " + '\n' +
            " WHERE" + '\n' +
            "   Спеціальність_1.№ = Головна.[Друга спеціальність]) AND " + '\n' +
            "   [Навчальний заклад].№ = Головна.Освіта)  AND " + '\n' +
            "   [Підготовчі курси].№ = Головна.[Закінчили підготовчі курси]) AND " + '\n' +
            "   Спеціальність.№ = Головна.Спеціальність) AND " + '\n' +
            "   Стать.№ = Головна.Стать " + '\n' +
            " ";

        public void Connect(string aConnectedString)
        {
            gConnection = new OleDbConnection(aConnectedString);
        }

        public OleDbDataAdapter RefreshBenefit()
        {
            OleDbConnection connection = new OleDbConnection(proect.Properties.Settings.Default.probaConnectionString);
            connection.Close();
            string sql = "SELECT * FROM SPRAV_BENEFIT";

            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataSet ds = new DataSet();
            da.Fill(ds, "SPRAV_BENEFIT");
            //dataGridView1.DataSource = ds.Tables["Головна"].DefaultView;
            connection.Close();
            return da;
        }

        

    }
}
