﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;




namespace proect
{
    public partial class golovna : Form
    {

        private string filterGender, filterHostel, filterSpets, filterLenguage;
        private LookupComboBox lcbSpec, lcbLenguage1;

        public golovna()
        {
            InitializeComponent();

            lcbSpec = new LookupComboBox(
                "select [№], [Спеціальність] FROM [Спеціальність]",
                "Спеціальність",
                comboBox1);
            lcbLenguage1 = new LookupComboBox(
                " select LANGUAGE_RECID, LANGUAGE_NAME FROM SPRAV_LANGUAGE ",
                "SPRAV_LANGUAGE",
                comboBox8
              );
            lcbLenguage1.keyValue = 6;
            ind = "0";
        }

        private string GetMainSql()
        {
            string lShowDeleted = "";
            if (!cbShowDeleted.Checked)
            {
                lShowDeleted = " AND (Головна.[Забрав документи] = false) ";
            }
            return " SELECT " +
                "   Головна.[№ п/п], " + //0
                "   Головна.[№ реєстрації], " +
                "   Головна.ПІБ, " +
                "   Спеціальність_1.Спеціальність as [Обрана Спеціальність], " +
                "   Головна.[Шкільна оцінка по алгебрі], " +
                "   Головна.[Шкільна оцінка по геометрії], " +
                "   Головна.[Шкільна оцінка по українській мові], " +
                "   Головна.[Шкільний середній бал], " +
                "   Головна.[Бал за результатами вступних екзаменів: українська мова], " +
                "   Головна.[Бал за результатами вступних екзаменів: математика], " +
                "   Головна.[Бал за результатами вступних екзаменів: рисунок], " +//10 
                "   Головна.[Бал за результатами вступних екзаменів: композиція], " +
                "   Головна.[Бал за підготовчі курси], " +
                "   Головна.[Сума балів (прохідний бал)], " +
                "   Головна.Пільги, " +
                "   Головна.Відзнака, " +
                "   [Підготовчі курси].Тривалість as [Проходив підготовчі курси], " +
                "   Головна.[Документ про освіту від 7 до 12 балів], " +
                "   Головна.Немісцевий, Головна.[Потребує гуртожитку], " +
                "   Стать.Стать as Стать, " + //20
                "   Головна.[Неповна сім'я], " +
                "   [Навчальний заклад].[Навчальний заклад] as Освіта, " +
                "   Головна.[Документ про освіту: серія], " +
                "   Головна.[Документ про освіту: номер], " +
                "   Спеціальність_2.Спеціальність as [Друга спеціальність], " +
                "   Спеціальність_3.Спеціальність as [Третя спеціальність], " +
                "   Головна.[Закінчили підготовчі курси], " +
                "   Головна.Зараховано, " +
                "   Головна.[Забрав документи], " +
                "   Головна.[Лист від підприємства], " + // 30
                "   SPRAV_LANGUAGE_1.LANGUAGE_NAME as Мова1, " +
                "   SPRAV_LANGUAGE_2.LANGUAGE_NAME as Мова2, " +
                "   Головна.Форма_навчання, " +
                "   Головна.[Документи]" +
                " FROM " +
                "   Головна, " +
                 "   Стать, " +
                  "   Спеціальність AS Спеціальність_1, " +
                  "   Спеціальність AS Спеціальність_2,  " +
                  "   Спеціальність AS Спеціальність_3,  " +
                  "   [Підготовчі курси], " +
                  "   [SPRAV_LANGUAGE] as SPRAV_LANGUAGE_1,  " +
                  "   [SPRAV_LANGUAGE] as SPRAV_LANGUAGE_2,  " +
                  "   [Навчальний заклад]  " +
                " WHERE " +
                "   ( " +
                "     (Стать.[№] = Головна.Стать) OR " +
                "     (isnull(Головна.Стать)) " +
                "   )  " +
                "  AND ( " +
                "     (Спеціальність_1.№ = Головна.Спеціальність) OR " +
                "     (isnull(Головна.Спеціальність)  AND (Спеціальність_1.[№] = 11))" +
                "   )  " +
                 "  AND (" +
                 "     ([Підготовчі курси].[№] = Головна.[Закінчили підготовчі курси]) OR " +
                 "     (isnull(Головна.[Закінчили підготовчі курси]))" +
                 "   )  " +
                "  AND ( " +
                "     ([Навчальний заклад].[№] = Головна.Освіта) OR " +
                "     (isnull(Головна.Освіта))" +
                "   )  " +
                "   AND  ( " +
                "     (Спеціальність_2.[№] = Головна.[Друга спеціальність]) OR " +
                "     (isnull(Головна.[Друга спеціальність]) AND (Спеціальність_2.[№] = 11))" +
                "   )  " +
                "  AND ( " +
                "     (Спеціальність_3.[№] = Головна.[Третя спеціальність]) OR " +
                "     (isnull(Головна.[Третя спеціальність]) AND (Спеціальність_3.[№] = 11))" +
                "   )  " +
                "  AND (" +
                "    (SPRAV_LANGUAGE_1.LANGUAGE_RECID = Головна.[Мова1]) OR " +//LANGUAGE_NAME
                "    (isnull(Головна.[Мова1]) and SPRAV_LANGUAGE_1.LANGUAGE_RECID = 6)" +
                "  )  " +
                "  AND (" +
                "    (SPRAV_LANGUAGE_2.LANGUAGE_RECID = Головна.[Мова2]) OR " +
                "    (isnull(Головна.[Мова2]) and SPRAV_LANGUAGE_2.LANGUAGE_RECID = 6)" +
                "  )" +
                lShowDeleted +
                " AND (Головна.[Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#) ";
            //" ORDER BY Головна.[Сума балів (прохідний бал)] ";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //this.WindowState = FormWindowState.Maximized;

            // TODO: данная строка кода позволяет загрузить данные в таблицу "probaDataSet.Головна". При необходимости она может быть перемещена или удалена.
            //  string sql = "SELECT Головна.[№ п/п],Головна.[№ реєстрації], Головна.ПІБ, Спеціальність.Спеціальність, Головна.[Шкільна оцінка по алгебрі], Головна.[Шкільна оцінка по геометрії], Головна.[Шкільна оцінка по українській мові], Головна.[Шкільний середній бал], Головна.[Бал за результатами вступних екзаменів: українська мова], Головна.[Бал за результатами вступних екзаменів: математика], Головна.[Бал за результатами вступних екзаменів: рисунок], Головна.[Бал за результатами вступних екзаменів: композиція], Головна.[Бал за підготовчі курси], Головна.[Сума балів (прохідний бал)], Головна.Пільги, Головна.Відзнака, [Підготовчі курси].Тривалість, Головна.[Документ про освіту від 7 до 12 балів], Головна.Немісцевий, Головна.[Потребує гуртожитку], Стать.Стать, Головна.[Неповна сім'я], [Навчальний заклад].[Навчальний заклад], Головна.[Документ про освіту: серія], Головна.[Документ про освіту: номер], Спеціальність.Cпеціальність FROM Головна, Спеціальність, Стать, [Навчальний заклад], [Підготовчі курси]  WHERE Спеціальність.№=Головна.Спеціальність AND  Головна.Стать=Стать.№ AND  Головна.Освіта=[Навчальний заклад].№ AND  Головна.[Закінчили підготовчі курси]=[Підготовчі курси].№";

            InputLanguage.CurrentInputLanguage =
                InputLanguage.FromCulture(new System.Globalization.CultureInfo("uk-UA"));

            string sql = GetMainSql();
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Головна");
            dataGridView1.DataSource = ds.Tables["Головна"].DefaultView;
            connection.Close();

            //this.головнаTableAdapter.Fill(this.probaDataSet.Головна);
            поискToolStripMenuItem_Click(sender, e);
            toolStripMenuItem2_Click(sender, e);

            lcbSpec.keyValue = 11;

            var lDate = DateTime.Now;

            lDate = lDate.AddMonths(0 - 5 - DateTime.Now.Month);
            lDate = lDate.AddDays(1 - DateTime.Now.Day);

            //dtpCreatedFrom.Value = lDate;

            lDate = lDate.AddMonths(11);
            lDate = lDate.AddDays(30);

            dtpCreatedTill.Value = lDate;

            UpdateDataWithFilters();
        }
        OleDbConnection connection = new OleDbConnection(proect.Properties.Settings.Default.probaConnectionString);
        private void button1_Click(object sender, EventArgs e)
        {
            toolStripMenuItem1_Click(sender, e);
        }


        string ind = "";



        private void button3_Click(object sender, EventArgs e)
        {

            connection.Close();
            string sql = "Delete * from Головна  where [№ п/п] = " + ind;


            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open(); //открытие соединения
            command.ExecuteReader(); //выполнение запроса на удаление

            connection.Close();
            UpdateDataWithFilters();

        }




        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            //UpdateDataWithFilters();
            /*connection.Close();
            string sql = "SELECT Головна.[№ п/п], Головна.[№ реєстрації], Головна.ПІБ, Спеціальність.Спеціальність as [Обрана Спеціальність], Головна.[Шкільна оцінка по алгебрі], Головна.[Шкільна оцінка по геометрії], Головна.[Шкільна оцінка по українській мові], Головна.[Шкільний середній бал], Головна.[Бал за результатами вступних екзаменів: українська мова], Головна.[Бал за результатами вступних екзаменів: математика], Головна.[Бал за результатами вступних екзаменів: рисунок], Головна.[Бал за результатами вступних екзаменів: композиція], Головна.[Бал за підготовчі курси], Головна.[Сума балів (прохідний бал)], Головна.Пільги, Головна.Відзнака, [Підготовчі курси].Тривалість as [Проходив підготовчі курси], Головна.[Документ про освіту від 7 до 12 балів], Головна.Немісцевий, Головна.[Потребує гуртожитку], Стать.Стать as Стать, Головна.[Неповна сім'я], [Навчальний заклад].[Навчальний заклад] as Освіта, Головна.[Документ про освіту: серія], Головна.[Документ про освіту: номер], Спеціальність_1.Спеціальність as [Друга спеціальність], Головна.Зараховано FROM Стать INNER JOIN (Спеціальність INNER JOIN ([Підготовчі курси] INNER JOIN ([Навчальний заклад] INNER JOIN (Спеціальність AS Спеціальність_1 INNER JOIN Головна ON Спеціальність_1.№ = Головна.[Друга спеціальність]) ON [Навчальний заклад].№ = Головна.Освіта) ON [Підготовчі курси].№ = Головна.[Закінчили підготовчі курси]) ON Спеціальність.№ = Головна.Спеціальність) ON Стать.№ = Головна.Стать WHERE (Головна.ПІБ like '" + textBox16.Text + "%')";

            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataSet ds = new DataSet();
            da.Fill(ds, "Головна");
            dataGridView1.DataSource = ds.Tables["Головна"].DefaultView;
            connection.Close(); //выполнение запроса на удаление
            */

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            filterSpets = "";
            if (lcbSpec != null)
            {
                if ((lcbSpec.keyValue != null) && (lcbSpec.keyValue != 11))
                {
                    filterSpets = " AND ((Головна.Спеціальність = " + lcbSpec.keyValue + ") "
                        + " OR (Головна.[Друга спеціальність] = " + lcbSpec.keyValue + ")  "
                        + " OR (Головна.[Третя спеціальність] = " + lcbSpec.keyValue + "))";
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            filterGender = " AND (Головна.Стать = 1) ";
            UpdateDataWithFilters();
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            filterGender = " AND (Головна.Стать = 2)";
            UpdateDataWithFilters();
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            filterGender = " ";
            UpdateDataWithFilters();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            filterHostel = " AND  Головна.[Потребує гуртожитку] = true ";
            UpdateDataWithFilters();
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {

            filterHostel = " AND  Головна.[Потребує гуртожитку] = false ";
            UpdateDataWithFilters();
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            filterHostel = " ";
            UpdateDataWithFilters();
        }


        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

            ind = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ind = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            connection.Close();
            if (ind != "")
            {
                string sql = "SELECT Головна.[№ п/п], Головна.[№ реєстрації], Головна.ПІБ, Спеціальність.Спеціальність as [Обрана Спеціальність], Головна.[Шкільна оцінка по алгебрі], Головна.[Шкільна оцінка по геометрії], Головна.[Шкільна оцінка по українській мові], Головна.[Шкільний середній бал], Головна.[Бал за результатами вступних екзаменів: українська мова], Головна.[Бал за результатами вступних екзаменів: математика], Головна.[Бал за результатами вступних екзаменів: рисунок], Головна.[Бал за результатами вступних екзаменів: композиція], Головна.[Бал за підготовчі курси], Головна.[Сума балів (прохідний бал)], Головна.Пільги, Головна.Відзнака, [Підготовчі курси].Тривалість as [Проходив підготовчі курси], Головна.[Документ про освіту від 7 до 12 балів], Головна.Немісцевий, Головна.[Потребує гуртожитку], Стать.Стать as Стать, Головна.[Неповна сім'я], [Навчальний заклад].[Навчальний заклад] as Освіта, Головна.[Документ про освіту: серія], Головна.[Документ про освіту: номер], Спеціальність_1.Спеціальність as [Друга спеціальність], Головна.Зараховано FROM Стать INNER JOIN (Спеціальність INNER JOIN ([Підготовчі курси] INNER JOIN ([Навчальний заклад] INNER JOIN (Спеціальність AS Спеціальність_1 INNER JOIN Головна ON Спеціальність_1.№ = Головна.[Друга спеціальність]) ON [Навчальний заклад].№ = Головна.Освіта) ON [Підготовчі курси].№ = Головна.[Закінчили підготовчі курси]) ON Спеціальність.№ = Головна.Спеціальність) ON Стать.№ = Головна.Стать where [№ п/п] = " + ind;

                OleDbCommand dc = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataReader or = dc.ExecuteReader();
                or.Read();
                if ((or[3].ToString() == "МЕУ") || (or[3].ToString() == "ТОМ") || (or[3].ToString() == "РПЗ") || (or[3].ToString() == "ЕМА")
                   || (or[3].ToString() == "КВЕТ") || (or[3].ToString() == "ЕП 9") || (or[3].ToString() == "ЕП 11"))
                {
                    label10.Enabled = false;
                    label11.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    label9.Enabled = true;
                    textBox9.Enabled = true;
                }
                if (or[3].ToString() == "ДЗ")
                {
                    label10.Enabled = true;
                    label11.Enabled = true;
                    textBox10.Enabled = true;
                    textBox11.Enabled = true;
                    label9.Enabled = false;
                    textBox9.Enabled = false;
                }
                textBox5.Text = or[1].ToString();
                textBox6.Text = or[2].ToString();
                if (or[3].ToString() == "МЕУ")
                {
                    comboBox2.SelectedItem = "МЕУ";
                }
                if (or[3].ToString() == "ТОМ")
                {
                    comboBox2.SelectedItem = "ТОМ";
                }
                if (or[3].ToString() == "РПЗ")
                {
                    comboBox2.SelectedItem = "РПЗ";
                }
                if (or[3].ToString() == "ЕМА")
                {
                    comboBox2.SelectedItem = "ЕМА";
                }
                if (or[3].ToString() == "КВЕТ")
                {
                    comboBox2.SelectedItem = "КВЕТ";
                }
                if (or[3].ToString() == "ДЗ")
                {
                    comboBox2.SelectedItem = "ДЗ";
                }
                if (or[3].ToString() == "ЕП 9")
                {
                    comboBox2.SelectedItem = "ЕП 9";
                }
                if (or[3].ToString() == "ЕП 11")
                {
                    comboBox2.SelectedItem = "ЕП 11";
                }
                textBox1.Text = or[4].ToString();
                textBox2.Text = or[5].ToString();
                textBox3.Text = or[6].ToString();
                textBox4.Text = or[7].ToString();
                textBox8.Text = or[8].ToString();
                textBox9.Text = or[9].ToString();
                textBox10.Text = or[10].ToString();
                textBox11.Text = or[11].ToString();
                textBox12.Text = or[12].ToString();
                textBox13.Text = or[13].ToString();
                comboBox6.SelectedItem = or[14].ToString();

                if (or[14].ToString() == "-")
                {
                    comboBox6.SelectedItem = "-";
                }
                if (or[14].ToString() == "сирота")
                {
                    comboBox6.SelectedItem = "сирота";
                }
                if (or[14].ToString() == "Чорнобилець")
                {
                    comboBox6.SelectedItem = "Чорнобилець";
                }
                if (or[14].ToString() == "інвалід І-ІІ групи")
                {
                    comboBox6.SelectedItem = "інвалід І-ІІ групи";
                }
                if (or[14].ToString() == "багатодітна родина")
                {
                    comboBox6.SelectedItem = "багатодітна родина";
                }
                textBox15.Text = or[15].ToString();
                //textBox17.Text = or[17].ToString();
                textBox19.Text = or[23].ToString();
                textBox20.Text = or[24].ToString();
                if (or[16].ToString() == "-")
                {
                    comboBox5.SelectedItem = "-";
                }
                if (or[16].ToString() == "1 місяць")
                {
                    comboBox5.SelectedItem = "1 місяць";
                }
                if (or[16].ToString() == "3 місяці")
                {
                    comboBox5.SelectedItem = "3 місяці";
                }
                if (or[16].ToString() == "6 місяців")
                {
                    comboBox5.SelectedItem = "6 місяців";
                }
                if (or[20].ToString() == "Чоловіча")
                {
                    radioButton12.Checked = true;
                }

                if (or[20].ToString() == "Жіноча")
                {
                    radioButton11.Checked = true;
                }

                if (Convert.ToBoolean(or[18].ToString()) == true)
                {
                    radioButton7.Checked = true;
                    checkBox1.Enabled = true;
                }
                else
                {
                    radioButton8.Checked = true;
                    checkBox1.Enabled = false;
                }

                //radioButton7.Checked = Convert.ToBoolean(or[18].ToString());

                //radioButton8.Checked = !Convert.ToBoolean(or[18].ToString());
                checkBox2.Checked = Convert.ToBoolean(or[17].ToString());
                checkBox1.Checked = Convert.ToBoolean(or[19].ToString());
                checkBox3.Checked = Convert.ToBoolean(or[26].ToString());

                radioButton9.Checked = Convert.ToBoolean(or[21].ToString());
                radioButton10.Checked = !Convert.ToBoolean(or[21].ToString());
                if (or[22].ToString() == "9 кл. загальноосвітньої школи")
                {
                    comboBox3.SelectedItem = "9 кл. загальноосвітньої школи";
                }
                if (or[22].ToString() == "11 кл. загальноосвітньої школи")
                {
                    comboBox3.SelectedItem = "11 кл. загальноосвітньої школи";
                }
                if (or[22].ToString() == "ПТУ")
                {
                    comboBox3.SelectedItem = "ПТУ";
                }
                if (or[25].ToString() == "МЕУ")
                {
                    comboBox4.SelectedItem = "МЕУ";
                }
                if (or[25].ToString() == "ТОМ")
                {
                    comboBox4.SelectedItem = "ТОМ";
                }
                if (or[25].ToString() == "РПЗ")
                {
                    comboBox4.SelectedItem = "РПЗ";
                }
                if (or[25].ToString() == "ЕМА")
                {
                    comboBox4.SelectedItem = "ЕМА";
                }
                if (or[25].ToString() == "КВЕТ")
                {
                    comboBox4.SelectedItem = "КВЕТ";
                }
                if (or[25].ToString() == "ДЗ")
                {
                    comboBox4.SelectedItem = "ДЗ";
                }
                if (or[25].ToString() == "ЕП 9")
                {
                    comboBox4.SelectedItem = "ЕП 9";
                }
                if (or[25].ToString() == "ЕП 11")
                {
                    comboBox4.SelectedItem = "ЕП 11";
                }

                ind = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();

            }
        }
        int s1 = 0;

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox4.SelectedItem == comboBox2.SelectedItem)
            {
                MessageBox.Show("ПОМИЛКА! Обрана та друга спеціальності не можуть бути однаковими");

            }
            else
            {

                if (comboBox2.SelectedItem == "МЕУ")
                {
                    s1 = 1;
                }
                if (comboBox2.SelectedItem == "ТОМ")
                {
                    s1 = 2;
                }
                if (comboBox2.SelectedItem == "РПЗ")
                {
                    s1 = 3;
                }
                if (comboBox2.SelectedItem == "ЕМА")
                {
                    s1 = 4;
                }
                if (comboBox2.SelectedItem == "КВЕТ")
                {
                    s1 = 5;
                }
                if (comboBox2.SelectedItem == "ДЗ")
                {
                    s1 = 6;
                }
                if (comboBox2.SelectedItem == "ЕП 9")
                {
                    s1 = 7;
                }
                if (comboBox2.SelectedItem == "ЕП 11")
                {
                    s1 = 8;
                }
                if ((comboBox2.SelectedItem == "МЕУ") || (comboBox2.SelectedItem == "ТОМ") || (comboBox2.SelectedItem == "РПЗ") || (comboBox2.SelectedItem == "ЕМА")
                     || (comboBox2.SelectedItem == "КВЕТ") || (comboBox2.SelectedItem == "ЕП 9") || (comboBox2.SelectedItem == "ЕП 11"))
                {
                    label10.Enabled = false;
                    label11.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    label9.Enabled = true;
                    textBox9.Enabled = true;
                    textBox10.Text = "0";
                    textBox11.Text = "0";

                }

                if (comboBox2.SelectedItem == "ДЗ")
                {
                    label9.Enabled = false;
                    textBox9.Enabled = false;
                    label10.Enabled = true;
                    label11.Enabled = true;
                    textBox10.Enabled = true;
                    textBox11.Enabled = true;
                    textBox9.Text = "0";


                }

            }
        }



        private void button2_Click(object sender, EventArgs e)
        {

            connection.Close();
            if (s2 == 0) s2 = 9;
            string sql = "Update Головна Set  [№ реєстрації]='" + textBox5.Text + "', ПІБ='" + textBox6.Text + "', Спеціальність = " + s1 + ", [Шкільна оцінка по алгебрі] = " + textBox1.Text + ", [Шкільна оцінка по геометрії] = " + textBox2.Text + ", [Шкільна оцінка по українській мові] = " + textBox3.Text + ", [Шкільний середній бал] = '" + textBox4.Text + "', [Бал за результатами вступних екзаменів: українська мова] = " + textBox8.Text + ", [Бал за результатами вступних екзаменів: математика]= " + textBox9.Text + ", [Бал за результатами вступних екзаменів: рисунок]= " + textBox10.Text + ", [Бал за результатами вступних екзаменів: композиція]= " + textBox11.Text + ", [Бал за підготовчі курси]= '" + textBox12.Text + "', [Сума балів (прохідний бал)]= '" + textBox13.Text + "', Пільги= '" + comboBox6.SelectedItem + "', Відзнака= '" + textBox15.Text + "', [Закінчили підготовчі курси] = " + k + ", [Документ про освіту від 7 до 12 балів]= '" + d + "', Немісцевий = '" + m + "', [Потребує гуртожитку]= '" + g + "', Стать= " + i + ", [Неповна сім'я] = '" + n + "', Освіта = " + o + ", [Документ про освіту: серія]='" + textBox19.Text + "', [Документ про освіту: номер]='" + textBox20.Text + "', [Друга спеціальність] = " + s2 + ", Зараховано = '" + z + "'  where [№ п/п] = " + ind;

            OleDbCommand dc = new OleDbCommand(sql, connection);

            connection.Open();
            OleDbDataReader or = dc.ExecuteReader();

            connection.Close();
            UpdateDataWithFilters();
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }
        double suma = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            if ((comboBox2.SelectedItem == "МЕУ") || (comboBox2.SelectedItem == "ТОМ") || (comboBox2.SelectedItem == "РПЗ") || (comboBox2.SelectedItem == "ЕМА")
               || (comboBox2.SelectedItem == "КВЕТ") || (comboBox2.SelectedItem == "ЕП 9") || (comboBox2.SelectedItem == "ЕП 11"))
            {

                suma = Convert.ToDouble(textBox1.Text) + Convert.ToDouble(textBox2.Text) + Convert.ToDouble(textBox3.Text) + Convert.ToDouble(textBox4.Text) + (Convert.ToDouble(textBox8.Text) + Convert.ToDouble(textBox9.Text)) + Convert.ToDouble(textBox12.Text);
                textBox13.Text = Convert.ToString(suma);
            }

            if (comboBox2.SelectedItem == "ДЗ")
            {

                suma = Convert.ToDouble(textBox1.Text) + Convert.ToDouble(textBox2.Text) + Convert.ToDouble(textBox3.Text) + Convert.ToDouble(textBox4.Text) + Convert.ToDouble(textBox8.Text) + Convert.ToDouble(textBox10.Text) + Convert.ToDouble(textBox11.Text) + Convert.ToDouble(textBox12.Text);
                textBox13.Text = Convert.ToString(suma);

            }
        }
        int k = 0;
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedItem == "-")
            {

                k = 1;

            }
            if (comboBox5.SelectedItem == "1 місяць")
            {

                k = 2;

            }
            if (comboBox5.SelectedItem == "3 місяці")
            {

                k = 3;

            }
            if (comboBox5.SelectedItem == "6 місяців")
            {

                k = 4;

            }
        }
        int m = 0;
        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton8.Checked == true)
            {
                checkBox1.Enabled = false;
                m = 0;
            }
        }

        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton7.Checked == true)
            {
                checkBox1.Enabled = true;
                m = 1;
            }
        }
        int g = 0;
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                g = 1;
            }
            if (checkBox1.Checked == false)
            {
                g = 0;
            }
        }

        int d = 0;
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                d = 1;
            }
            if (checkBox2.Checked == false)
            {
                d = 0;
            }
        }


        int i = 0;
        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton12.Checked == true)
            {
                i = 1;
            }
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {

            if (radioButton11.Checked == true)
            {
                i = 2;
            }
        }
        int n = 0;
        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton10.Checked == true)
            {
                n = 0;
            }
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton9.Checked == true)
            {
                n = 1;
            }
        }
        int o = 0;
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox3.SelectedItem == "9 кл. загальноосвітньої школи")
            {

                o = 1;

            }
            if (comboBox3.SelectedItem == "11 кл. загальноосвітньої школи")
            {

                o = 2;

            }
            if (comboBox3.SelectedItem == "ПТУ")
            {

                o = 3;

            }
        }
        int s2 = 0;
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox4.SelectedItem == comboBox2.SelectedItem)
            {
                MessageBox.Show("ПОМИЛКА! Обрана та друга спеціальності не можуть бути однаковими");

            }
            else
            {
                if (comboBox4.SelectedItem == "МЕУ")
                {
                    s2 = 1;
                }
                if (comboBox4.SelectedItem == "-")
                {
                    s2 = 9;
                }
                if (comboBox4.SelectedItem == "ТОМ")
                {
                    s2 = 2;
                }
                if (comboBox4.SelectedItem == "РПЗ")
                {
                    s2 = 3;
                }
                if (comboBox4.SelectedItem == "ЕМА")
                {
                    s2 = 4;
                }
                if (comboBox4.SelectedItem == "КВЕТ")
                {
                    s2 = 5;
                }
                if (comboBox4.SelectedItem == "ДЗ")
                {
                    s2 = 6;
                }
                if (comboBox4.SelectedItem == "ЕП 9")
                {
                    s2 = 7;
                }
                if (comboBox4.SelectedItem == "ЕП 11")
                {
                    s2 = 8;
                }
                else
                {
                    s2 = 9;
                }
            }
        }

        int z = 0;
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                z = 1;
            }
            if (checkBox3.Checked == false)
            {
                z = 0;
            }
        }
        Excel.Application xlApp;

        private void button4_Click(object sender, EventArgs e)
        {
            string lZvit = CopyExcelDocFromTemplate();
            connection.Close();

            // Подано заяв
            string sql1 = "Select COUNT(*) FROM Головна where ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
               "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)";
            OleDbCommand dc1 = new OleDbCommand(sql1, connection);
            connection.Open();
            OleDbDataReader or1 = dc1.ExecuteReader();
            or1.Read();
            String s1 = or1[0].ToString();

            string sql2 = "Select COUNT(*) FROM Головна where (Спеціальність = 1) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc2 = new OleDbCommand(sql2, connection);

            OleDbDataReader or2 = dc2.ExecuteReader();
            or2.Read();
            String s2 = or2[0].ToString();

            string sql3 = "Select COUNT(*) FROM Головна where (Спеціальність = 5) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)";
            OleDbCommand dc3 = new OleDbCommand(sql3, connection);

            OleDbDataReader or3 = dc3.ExecuteReader();
            or3.Read();
            String s3 = or3[0].ToString();

            string sql4 = "Select COUNT(*) FROM Головна where (Спеціальність = 4) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc4 = new OleDbCommand(sql4, connection);

            OleDbDataReader or4 = dc4.ExecuteReader();
            or4.Read();
            String s4 = or4[0].ToString();

            string sql5 = "Select COUNT(*) FROM Головна where (Спеціальність = 2) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc5 = new OleDbCommand(sql5, connection);

            OleDbDataReader or5 = dc5.ExecuteReader();
            or5.Read();
            String s5 = or5[0].ToString();


            string sql6 = "Select COUNT(*) FROM Головна where (Спеціальність = 3) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc6 = new OleDbCommand(sql6, connection);

            OleDbDataReader or6 = dc6.ExecuteReader();
            or6.Read();
            String s6 = or6[0].ToString();


            string sql7 = "Select COUNT(*) FROM Головна where (Спеціальність = 6) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc7 = new OleDbCommand(sql7, connection);

            OleDbDataReader or7 = dc7.ExecuteReader();
            or7.Read();
            String s7 = or7[0].ToString();


            string sql8 = "Select COUNT(*) FROM Головна where (Спеціальність = 7) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc8 = new OleDbCommand(sql8, connection);

            OleDbDataReader or8 = dc8.ExecuteReader();
            or8.Read();
            String s8 = or8[0].ToString();


            string sql9 = "Select COUNT(*) FROM Головна where (Спеціальність = 8) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc9 = new OleDbCommand(sql9, connection);

            OleDbDataReader or9 = dc9.ExecuteReader();
            or9.Read();
            String s9 = or9[0].ToString();
            // подано заяв кінець



            // зараховано
            string sql11 = "Select COUNT(*) FROM Головна where ( Зараховано = true)" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc11 = new OleDbCommand(sql11, connection);

            OleDbDataReader or11 = dc11.ExecuteReader();
            or11.Read();
            String s11 = or11[0].ToString();

            string sql12 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc12 = new OleDbCommand(sql12, connection);

            OleDbDataReader or12 = dc12.ExecuteReader();
            or12.Read();
            String s12 = or12[0].ToString();

            string sql13 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc13 = new OleDbCommand(sql13, connection);

            OleDbDataReader or13 = dc13.ExecuteReader();
            or13.Read();
            String s13 = or13[0].ToString();

            string sql14 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc14 = new OleDbCommand(sql14, connection);

            OleDbDataReader or14 = dc14.ExecuteReader();
            or14.Read();
            String s14 = or14[0].ToString();

            string sql15 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc15 = new OleDbCommand(sql15, connection);

            OleDbDataReader or15 = dc15.ExecuteReader();
            or15.Read();
            String s15 = or15[0].ToString();


            string sql16 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc16 = new OleDbCommand(sql16, connection);

            OleDbDataReader or16 = dc16.ExecuteReader();
            or16.Read();
            String s16 = or16[0].ToString();

            string sql17 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc17 = new OleDbCommand(sql17, connection);

            OleDbDataReader or17 = dc17.ExecuteReader();
            or17.Read();
            String s17 = or17[0].ToString();


            string sql18 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc18 = new OleDbCommand(sql18, connection);

            OleDbDataReader or18 = dc18.ExecuteReader();
            or18.Read();
            String s18 = or18[0].ToString();

            string sql19 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc19 = new OleDbCommand(sql19, connection);

            OleDbDataReader or19 = dc19.ExecuteReader();
            or19.Read();
            String s19 = or19[0].ToString();
            // зараховано конец




            // відзнака
            string sql21 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc21 = new OleDbCommand(sql21, connection);

            OleDbDataReader or21 = dc21.ExecuteReader();
            or21.Read();
            String s21 = or21[0].ToString();

            string sql22 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc22 = new OleDbCommand(sql22, connection);

            OleDbDataReader or22 = dc22.ExecuteReader();
            or22.Read();
            String s22 = or22[0].ToString();

            string sql23 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc23 = new OleDbCommand(sql23, connection);

            OleDbDataReader or23 = dc23.ExecuteReader();
            or23.Read();
            String s23 = or23[0].ToString();

            string sql24 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc24 = new OleDbCommand(sql24, connection);

            OleDbDataReader or24 = dc24.ExecuteReader();
            or24.Read();
            String s24 = or24[0].ToString();

            string sql25 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc25 = new OleDbCommand(sql25, connection);

            OleDbDataReader or25 = dc25.ExecuteReader();
            or25.Read();
            String s25 = or25[0].ToString();


            string sql26 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc26 = new OleDbCommand(sql26, connection);

            OleDbDataReader or26 = dc26.ExecuteReader();
            or26.Read();
            String s26 = or26[0].ToString();

            string sql27 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc27 = new OleDbCommand(sql27, connection);

            OleDbDataReader or27 = dc27.ExecuteReader();
            or27.Read();
            String s27 = or27[0].ToString();


            string sql28 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc28 = new OleDbCommand(sql28, connection);

            OleDbDataReader or28 = dc28.ExecuteReader();
            or28.Read();
            String s28 = or28[0].ToString();

            string sql29 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Відзнака <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc29 = new OleDbCommand(sql29, connection);

            OleDbDataReader or29 = dc29.ExecuteReader();
            or29.Read();
            String s29 = or29[0].ToString();
            // відзнака конец





            // 7-12
            string sql31 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc31 = new OleDbCommand(sql31, connection);

            OleDbDataReader or31 = dc31.ExecuteReader();
            or31.Read();
            String s31 = or31[0].ToString();

            string sql32 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc32 = new OleDbCommand(sql32, connection);

            OleDbDataReader or32 = dc32.ExecuteReader();
            or32.Read();
            String s32 = or32[0].ToString();

            string sql33 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc33 = new OleDbCommand(sql33, connection);

            OleDbDataReader or33 = dc33.ExecuteReader();
            or33.Read();
            String s33 = or33[0].ToString();

            string sql34 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc34 = new OleDbCommand(sql34, connection);

            OleDbDataReader or34 = dc34.ExecuteReader();
            or34.Read();
            String s34 = or34[0].ToString();

            string sql35 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc35 = new OleDbCommand(sql35, connection);

            OleDbDataReader or35 = dc35.ExecuteReader();
            or35.Read();
            String s35 = or35[0].ToString();


            string sql36 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc36 = new OleDbCommand(sql36, connection);

            OleDbDataReader or36 = dc36.ExecuteReader();
            or36.Read();
            String s36 = or36[0].ToString();

            string sql37 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc37 = new OleDbCommand(sql37, connection);

            OleDbDataReader or37 = dc37.ExecuteReader();
            or37.Read();
            String s37 = or37[0].ToString();


            string sql38 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc38 = new OleDbCommand(sql38, connection);

            OleDbDataReader or38 = dc38.ExecuteReader();
            or38.Read();
            String s38 = or38[0].ToString();

            string sql39 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and ([Документ про освіту від 7 до 12 балів] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc39 = new OleDbCommand(sql39, connection);

            OleDbDataReader or39 = dc39.ExecuteReader();
            or39.Read();
            String s39 = or39[0].ToString();
            // 7-12 конец 






            // курси

            string sql41 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc41 = new OleDbCommand(sql41, connection);

            OleDbDataReader or41 = dc41.ExecuteReader();
            or41.Read();
            String s41 = or41[0].ToString();

            string sql42 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc42 = new OleDbCommand(sql42, connection);

            OleDbDataReader or42 = dc42.ExecuteReader();
            or42.Read();
            String s42 = or42[0].ToString();

            string sql43 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc43 = new OleDbCommand(sql43, connection);

            OleDbDataReader or43 = dc43.ExecuteReader();
            or43.Read();
            String s43 = or43[0].ToString();

            string sql44 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc44 = new OleDbCommand(sql44, connection);

            OleDbDataReader or44 = dc44.ExecuteReader();
            or44.Read();
            String s44 = or44[0].ToString();

            string sql45 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc45 = new OleDbCommand(sql45, connection);

            OleDbDataReader or45 = dc45.ExecuteReader();
            or45.Read();
            String s45 = or45[0].ToString();


            string sql46 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc46 = new OleDbCommand(sql46, connection);

            OleDbDataReader or46 = dc46.ExecuteReader();
            or46.Read();
            String s46 = or46[0].ToString();

            string sql47 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc47 = new OleDbCommand(sql47, connection);

            OleDbDataReader or47 = dc47.ExecuteReader();
            or47.Read();
            String s47 = or47[0].ToString();


            string sql48 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc48 = new OleDbCommand(sql48, connection);

            OleDbDataReader or48 = dc48.ExecuteReader();
            or48.Read();
            String s48 = or48[0].ToString();

            string sql49 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and ([Закінчили підготовчі курси] <> 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc49 = new OleDbCommand(sql49, connection);

            OleDbDataReader or49 = dc49.ExecuteReader();
            or49.Read();
            String s49 = or49[0].ToString();
            //курси конец






            // пільги
            string sql51 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc51 = new OleDbCommand(sql51, connection);

            OleDbDataReader or51 = dc51.ExecuteReader();
            or51.Read();
            String s51 = or51[0].ToString();

            string sql52 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc52 = new OleDbCommand(sql52, connection);

            OleDbDataReader or52 = dc52.ExecuteReader();
            or52.Read();
            String s52 = or52[0].ToString();

            string sql53 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc53 = new OleDbCommand(sql53, connection);

            OleDbDataReader or53 = dc53.ExecuteReader();
            or53.Read();
            String s53 = or53[0].ToString();

            string sql54 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc54 = new OleDbCommand(sql54, connection);

            OleDbDataReader or54 = dc54.ExecuteReader();
            or54.Read();
            String s54 = or54[0].ToString();

            string sql55 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc55 = new OleDbCommand(sql55, connection);

            OleDbDataReader or55 = dc55.ExecuteReader();
            or55.Read();
            String s55 = or55[0].ToString();


            string sql56 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc56 = new OleDbCommand(sql56, connection);

            OleDbDataReader or56 = dc56.ExecuteReader();
            or56.Read();
            String s56 = or56[0].ToString();

            string sql57 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc57 = new OleDbCommand(sql57, connection);

            OleDbDataReader or57 = dc57.ExecuteReader();
            or57.Read();
            String s57 = or57[0].ToString();


            string sql58 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc58 = new OleDbCommand(sql58, connection);

            OleDbDataReader or58 = dc58.ExecuteReader();
            or58.Read();
            String s58 = or58[0].ToString();

            string sql59 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Пільги <> '-'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc59 = new OleDbCommand(sql59, connection);

            OleDbDataReader or59 = dc59.ExecuteReader();
            or59.Read();
            String s59 = or59[0].ToString();
            //конец пільги







            //сироти
            string sql511 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc511 = new OleDbCommand(sql511, connection);

            OleDbDataReader or511 = dc511.ExecuteReader();
            or511.Read();
            String s511 = or511[0].ToString();

            string sql521 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc521 = new OleDbCommand(sql521, connection);

            OleDbDataReader or521 = dc521.ExecuteReader();
            or521.Read();
            String s521 = or521[0].ToString();

            string sql531 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc531 = new OleDbCommand(sql531, connection);

            OleDbDataReader or531 = dc531.ExecuteReader();
            or531.Read();
            String s531 = or531[0].ToString();

            string sql541 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc541 = new OleDbCommand(sql541, connection);

            OleDbDataReader or541 = dc541.ExecuteReader();
            or541.Read();
            String s541 = or541[0].ToString();

            string sql551 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc551 = new OleDbCommand(sql551, connection);

            OleDbDataReader or551 = dc551.ExecuteReader();
            or551.Read();
            String s551 = or551[0].ToString();


            string sql561 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc561 = new OleDbCommand(sql561, connection);

            OleDbDataReader or561 = dc561.ExecuteReader();
            or561.Read();
            String s561 = or561[0].ToString();

            string sql571 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc571 = new OleDbCommand(sql571, connection);

            OleDbDataReader or571 = dc571.ExecuteReader();
            or571.Read();
            String s571 = or571[0].ToString();


            string sql581 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc581 = new OleDbCommand(sql581, connection);

            OleDbDataReader or581 = dc581.ExecuteReader();
            or581.Read();
            String s581 = or581[0].ToString();

            string sql591 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Пільги = 'сирота'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc591 = new OleDbCommand(sql591, connection);

            OleDbDataReader or591 = dc591.ExecuteReader();
            or591.Read();
            String s591 = or591[0].ToString();
            // сироти конец


            // Чорнобилець
            string sql512 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Пільги = 'Чорнобилець'))  AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc512 = new OleDbCommand(sql512, connection);

            OleDbDataReader or512 = dc512.ExecuteReader();
            or512.Read();
            String s512 = or512[0].ToString();

            string sql522 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc522 = new OleDbCommand(sql522, connection);

            OleDbDataReader or522 = dc522.ExecuteReader();
            or522.Read();
            String s522 = or522[0].ToString();

            string sql532 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc532 = new OleDbCommand(sql532, connection);

            OleDbDataReader or532 = dc532.ExecuteReader();
            or532.Read();
            String s532 = or532[0].ToString();

            string sql542 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc542 = new OleDbCommand(sql542, connection);

            OleDbDataReader or542 = dc542.ExecuteReader();
            or542.Read();
            String s542 = or542[0].ToString();

            string sql552 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc552 = new OleDbCommand(sql552, connection);

            OleDbDataReader or552 = dc552.ExecuteReader();
            or552.Read();
            String s552 = or552[0].ToString();


            string sql562 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc562 = new OleDbCommand(sql562, connection);

            OleDbDataReader or562 = dc562.ExecuteReader();
            or562.Read();
            String s562 = or562[0].ToString();

            string sql572 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc572 = new OleDbCommand(sql572, connection);

            OleDbDataReader or572 = dc572.ExecuteReader();
            or572.Read();
            String s572 = or572[0].ToString();


            string sql582 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc582 = new OleDbCommand(sql582, connection);

            OleDbDataReader or582 = dc582.ExecuteReader();
            or582.Read();
            String s582 = or582[0].ToString();

            string sql592 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Пільги = 'Чорнобилець'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc592 = new OleDbCommand(sql592, connection);

            OleDbDataReader or592 = dc592.ExecuteReader();
            or592.Read();
            String s592 = or592[0].ToString();
            //конец Чорнобилець






            //інвалід
            string sql513 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc513 = new OleDbCommand(sql513, connection);

            OleDbDataReader or513 = dc513.ExecuteReader();
            or513.Read();
            String s513 = or513[0].ToString();

            string sql523 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc523 = new OleDbCommand(sql523, connection);

            OleDbDataReader or523 = dc523.ExecuteReader();
            or523.Read();
            String s523 = or523[0].ToString();

            string sql533 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc533 = new OleDbCommand(sql533, connection);

            OleDbDataReader or533 = dc533.ExecuteReader();
            or533.Read();
            String s533 = or533[0].ToString();

            string sql543 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc543 = new OleDbCommand(sql543, connection);

            OleDbDataReader or543 = dc543.ExecuteReader();
            or543.Read();
            String s543 = or543[0].ToString();

            string sql553 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc553 = new OleDbCommand(sql553, connection);

            OleDbDataReader or553 = dc553.ExecuteReader();
            or553.Read();
            String s553 = or553[0].ToString();


            string sql563 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc563 = new OleDbCommand(sql563, connection);

            OleDbDataReader or563 = dc563.ExecuteReader();
            or563.Read();
            String s563 = or563[0].ToString();

            string sql573 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc573 = new OleDbCommand(sql573, connection);

            OleDbDataReader or573 = dc573.ExecuteReader();
            or573.Read();
            String s573 = or573[0].ToString();


            string sql583 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc583 = new OleDbCommand(sql583, connection);

            OleDbDataReader or583 = dc583.ExecuteReader();
            or583.Read();
            String s583 = or583[0].ToString();

            string sql593 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Пільги = 'інвалід І-ІІ групи'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc593 = new OleDbCommand(sql593, connection);

            OleDbDataReader or593 = dc593.ExecuteReader();
            or593.Read();
            String s593 = or593[0].ToString();

            //інвалід конец







            //багатодытня
            string sql514 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc514 = new OleDbCommand(sql514, connection);

            OleDbDataReader or514 = dc514.ExecuteReader();
            or514.Read();
            String s514 = or514[0].ToString();

            string sql524 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc524 = new OleDbCommand(sql524, connection);

            OleDbDataReader or524 = dc524.ExecuteReader();
            or524.Read();
            String s524 = or524[0].ToString();

            string sql534 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc534 = new OleDbCommand(sql534, connection);

            OleDbDataReader or534 = dc534.ExecuteReader();
            or534.Read();
            String s534 = or534[0].ToString();

            string sql544 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc544 = new OleDbCommand(sql544, connection);

            OleDbDataReader or544 = dc544.ExecuteReader();
            or544.Read();
            String s544 = or544[0].ToString();

            string sql554 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc554 = new OleDbCommand(sql554, connection);

            OleDbDataReader or554 = dc554.ExecuteReader();
            or554.Read();
            String s554 = or554[0].ToString();


            string sql564 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc564 = new OleDbCommand(sql564, connection);

            OleDbDataReader or564 = dc564.ExecuteReader();
            or564.Read();
            String s564 = or564[0].ToString();

            string sql574 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc574 = new OleDbCommand(sql574, connection);

            OleDbDataReader or574 = dc574.ExecuteReader();
            or574.Read();
            String s574 = or574[0].ToString();


            string sql584 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc584 = new OleDbCommand(sql584, connection);

            OleDbDataReader or584 = dc584.ExecuteReader();
            or584.Read();
            String s584 = or584[0].ToString();

            string sql594 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Пільги = 'багатодітна родина'))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc594 = new OleDbCommand(sql594, connection);

            OleDbDataReader or594 = dc594.ExecuteReader();
            or594.Read();
            String s594 = or594[0].ToString();
            //багатодытня конец







            //неповні
            string sql61 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc61 = new OleDbCommand(sql61, connection);

            OleDbDataReader or61 = dc61.ExecuteReader();
            or61.Read();
            String s61 = or61[0].ToString();

            string sql62 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and ([Неповна сім'я] = true )) AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc62 = new OleDbCommand(sql62, connection);

            OleDbDataReader or62 = dc62.ExecuteReader();
            or62.Read();
            String s62 = or62[0].ToString();

            string sql63 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc63 = new OleDbCommand(sql63, connection);

            OleDbDataReader or63 = dc63.ExecuteReader();
            or63.Read();
            String s63 = or63[0].ToString();

            string sql64 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc64 = new OleDbCommand(sql64, connection);

            OleDbDataReader or64 = dc64.ExecuteReader();
            or64.Read();
            String s64 = or64[0].ToString();

            string sql65 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc65 = new OleDbCommand(sql65, connection);

            OleDbDataReader or65 = dc65.ExecuteReader();
            or65.Read();
            String s65 = or65[0].ToString();


            string sql66 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc66 = new OleDbCommand(sql66, connection);

            OleDbDataReader or66 = dc66.ExecuteReader();
            or66.Read();
            String s66 = or66[0].ToString();

            string sql67 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc67 = new OleDbCommand(sql67, connection);

            OleDbDataReader or67 = dc67.ExecuteReader();
            or67.Read();
            String s67 = or67[0].ToString();


            string sql68 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc68 = new OleDbCommand(sql68, connection);

            OleDbDataReader or68 = dc68.ExecuteReader();
            or68.Read();
            String s68 = or68[0].ToString();

            string sql69 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and ([Неповна сім'я] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc69 = new OleDbCommand(sql69, connection);

            OleDbDataReader or69 = dc69.ExecuteReader();
            or69.Read();
            String s69 = or69[0].ToString();

            //неповні конец





            //чоловіча


            string sql71 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc71 = new OleDbCommand(sql71, connection);

            OleDbDataReader or71 = dc71.ExecuteReader();
            or71.Read();
            String s71 = or71[0].ToString();

            string sql72 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc72 = new OleDbCommand(sql72, connection);

            OleDbDataReader or72 = dc72.ExecuteReader();
            or72.Read();
            String s72 = or72[0].ToString();

            string sql73 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Стать = 1 ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc73 = new OleDbCommand(sql73, connection);

            OleDbDataReader or73 = dc73.ExecuteReader();
            or73.Read();
            String s73 = or73[0].ToString();

            string sql74 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc74 = new OleDbCommand(sql74, connection);

            OleDbDataReader or74 = dc74.ExecuteReader();
            or74.Read();
            String s74 = or74[0].ToString();

            string sql75 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc75 = new OleDbCommand(sql75, connection);

            OleDbDataReader or75 = dc75.ExecuteReader();
            or75.Read();
            String s75 = or75[0].ToString();


            string sql76 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Стать = 1 ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc76 = new OleDbCommand(sql76, connection);

            OleDbDataReader or76 = dc76.ExecuteReader();
            or76.Read();
            String s76 = or76[0].ToString();

            string sql77 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc77 = new OleDbCommand(sql77, connection);

            OleDbDataReader or77 = dc77.ExecuteReader();
            or77.Read();
            String s77 = or77[0].ToString();


            string sql78 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc78 = new OleDbCommand(sql78, connection);

            OleDbDataReader or78 = dc78.ExecuteReader();
            or78.Read();
            String s78 = or78[0].ToString();

            string sql79 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Стать = 1 ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc79 = new OleDbCommand(sql79, connection);

            OleDbDataReader or79 = dc79.ExecuteReader();
            or79.Read();
            String s79 = or79[0].ToString();

            //чоловіча конец






            //жіноча
            string sql81 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc81 = new OleDbCommand(sql81, connection);

            OleDbDataReader or81 = dc81.ExecuteReader();
            or81.Read();
            String s81 = or81[0].ToString();

            string sql82 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc82 = new OleDbCommand(sql82, connection);

            OleDbDataReader or82 = dc82.ExecuteReader();
            or82.Read();
            String s82 = or82[0].ToString();

            string sql83 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Стать = 2 ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc83 = new OleDbCommand(sql83, connection);

            OleDbDataReader or83 = dc83.ExecuteReader();
            or83.Read();
            String s83 = or83[0].ToString();

            string sql84 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc84 = new OleDbCommand(sql84, connection);

            OleDbDataReader or84 = dc84.ExecuteReader();
            or84.Read();
            String s84 = or84[0].ToString();

            string sql85 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc85 = new OleDbCommand(sql85, connection);

            OleDbDataReader or85 = dc85.ExecuteReader();
            or85.Read();
            String s85 = or85[0].ToString();


            string sql86 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Стать = 2 ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc86 = new OleDbCommand(sql86, connection);

            OleDbDataReader or86 = dc86.ExecuteReader();
            or86.Read();
            String s86 = or86[0].ToString();

            string sql87 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc87 = new OleDbCommand(sql87, connection);

            OleDbDataReader or87 = dc87.ExecuteReader();
            or87.Read();
            String s87 = or87[0].ToString();


            string sql88 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc88 = new OleDbCommand(sql88, connection);

            OleDbDataReader or88 = dc88.ExecuteReader();
            or88.Read();
            String s88 = or88[0].ToString();

            string sql89 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Стать = 2 ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc89 = new OleDbCommand(sql89, connection);

            OleDbDataReader or89 = dc89.ExecuteReader();
            or89.Read();
            String s89 = or89[0].ToString();
            //жыноча конец







            //гуртожиток
            string sql91 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and ([Потребує гуртожитку]= true ))  AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc91 = new OleDbCommand(sql91, connection);

            OleDbDataReader or91 = dc91.ExecuteReader();
            or91.Read();
            String s91 = or91[0].ToString();

            string sql92 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc92 = new OleDbCommand(sql92, connection);

            OleDbDataReader or92 = dc92.ExecuteReader();
            or92.Read();
            String s92 = or92[0].ToString();

            string sql93 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc93 = new OleDbCommand(sql93, connection);

            OleDbDataReader or93 = dc93.ExecuteReader();
            or93.Read();
            String s93 = or93[0].ToString();

            string sql94 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc94 = new OleDbCommand(sql94, connection);

            OleDbDataReader or94 = dc94.ExecuteReader();
            or94.Read();
            String s94 = or94[0].ToString();

            string sql95 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc95 = new OleDbCommand(sql95, connection);

            OleDbDataReader or95 = dc95.ExecuteReader();
            or95.Read();
            String s95 = or95[0].ToString();


            string sql96 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and ([Потребує гуртожитку]= true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc96 = new OleDbCommand(sql96, connection);

            OleDbDataReader or96 = dc96.ExecuteReader();
            or96.Read();
            String s96 = or96[0].ToString();

            string sql97 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc97 = new OleDbCommand(sql97, connection);

            OleDbDataReader or97 = dc97.ExecuteReader();
            or97.Read();
            String s97 = or97[0].ToString();


            string sql98 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc98 = new OleDbCommand(sql98, connection);

            OleDbDataReader or98 = dc98.ExecuteReader();
            or98.Read();
            String s98 = or98[0].ToString();

            string sql99 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and ([Потребує гуртожитку] = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc99 = new OleDbCommand(sql99, connection);

            OleDbDataReader or99 = dc99.ExecuteReader();
            or99.Read();
            String s99 = or99[0].ToString();

            //гуртожиток кінець







            //гуртожитоок чоловіки
            string sql911 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and ([Потребує гуртожитку]= true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc911 = new OleDbCommand(sql911, connection);

            OleDbDataReader or911 = dc911.ExecuteReader();
            or911.Read();
            String s911 = or911[0].ToString();

            string sql921 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc921 = new OleDbCommand(sql921, connection);

            OleDbDataReader or921 = dc921.ExecuteReader();
            or921.Read();
            String s921 = or921[0].ToString();

            string sql931 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc931 = new OleDbCommand(sql931, connection);

            OleDbDataReader or931 = dc931.ExecuteReader();
            or931.Read();
            String s931 = or931[0].ToString();

            string sql941 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc941 = new OleDbCommand(sql941, connection);

            OleDbDataReader or941 = dc941.ExecuteReader();
            or941.Read();
            String s941 = or941[0].ToString();

            string sql951 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc951 = new OleDbCommand(sql951, connection);

            OleDbDataReader or951 = dc951.ExecuteReader();
            or951.Read();
            String s951 = or951[0].ToString();


            string sql961 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and ([Потребує гуртожитку]= true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc961 = new OleDbCommand(sql961, connection);

            OleDbDataReader or961 = dc961.ExecuteReader();
            or961.Read();
            String s961 = or961[0].ToString();

            string sql971 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc971 = new OleDbCommand(sql971, connection);

            OleDbDataReader or971 = dc971.ExecuteReader();
            or971.Read();
            String s971 = or971[0].ToString();


            string sql981 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc981 = new OleDbCommand(sql981, connection);

            OleDbDataReader or981 = dc981.ExecuteReader();
            or981.Read();
            String s981 = or981[0].ToString();

            string sql991 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and ([Потребує гуртожитку] = true ) and (Стать = 1))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc991 = new OleDbCommand(sql991, connection);

            OleDbDataReader or991 = dc991.ExecuteReader();
            or991.Read();
            String s991 = or991[0].ToString();

            //конец гуртожиток чоловіки




            //гуртожиток жінки
            string sql912 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and ([Потребує гуртожитку]= true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc912 = new OleDbCommand(sql912, connection);

            OleDbDataReader or912 = dc912.ExecuteReader();
            or912.Read();
            String s912 = or912[0].ToString();

            string sql922 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc922 = new OleDbCommand(sql922, connection);

            OleDbDataReader or922 = dc922.ExecuteReader();
            or922.Read();
            String s922 = or922[0].ToString();

            string sql932 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc932 = new OleDbCommand(sql932, connection);

            OleDbDataReader or932 = dc932.ExecuteReader();
            or932.Read();
            String s932 = or931[0].ToString();

            string sql942 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc942 = new OleDbCommand(sql942, connection);

            OleDbDataReader or942 = dc942.ExecuteReader();
            or942.Read();
            String s942 = or942[0].ToString();

            string sql952 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc952 = new OleDbCommand(sql952, connection);

            OleDbDataReader or952 = dc952.ExecuteReader();
            or952.Read();
            String s952 = or952[0].ToString();


            string sql962 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and ([Потребує гуртожитку]= true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc962 = new OleDbCommand(sql962, connection);

            OleDbDataReader or962 = dc962.ExecuteReader();
            or962.Read();
            String s962 = or962[0].ToString();

            string sql972 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc972 = new OleDbCommand(sql972, connection);

            OleDbDataReader or972 = dc972.ExecuteReader();
            or972.Read();
            String s972 = or972[0].ToString();


            string sql982 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc982 = new OleDbCommand(sql982, connection);

            OleDbDataReader or982 = dc982.ExecuteReader();
            or982.Read();
            String s982 = or982[0].ToString();

            string sql992 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and ([Потребує гуртожитку] = true ) and (Стать = 2))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc992 = new OleDbCommand(sql992, connection);

            OleDbDataReader or992 = dc992.ExecuteReader();
            or992.Read();
            String s992 = or992[0].ToString();
            //гуртожиток жінки кінець





            //немісцеві
            string sql101 = "Select COUNT(*) FROM Головна where (( Зараховано = true) and (Немісцевий= true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc101 = new OleDbCommand(sql101, connection);

            OleDbDataReader or101 = dc101.ExecuteReader();
            or101.Read();
            String s101 = or101[0].ToString();

            string sql102 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 1) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc102 = new OleDbCommand(sql102, connection);

            OleDbDataReader or102 = dc102.ExecuteReader();
            or102.Read();
            String s102 = or102[0].ToString();

            string sql103 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 5) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc103 = new OleDbCommand(sql103, connection);

            OleDbDataReader or103 = dc103.ExecuteReader();
            or103.Read();
            String s103 = or103[0].ToString();

            string sql104 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 4) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc104 = new OleDbCommand(sql104, connection);

            OleDbDataReader or104 = dc104.ExecuteReader();
            or104.Read();
            String s104 = or104[0].ToString();

            string sql105 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 2) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc105 = new OleDbCommand(sql105, connection);

            OleDbDataReader or105 = dc105.ExecuteReader();
            or105.Read();
            String s105 = or105[0].ToString();


            string sql106 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 3) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc106 = new OleDbCommand(sql106, connection);

            OleDbDataReader or106 = dc106.ExecuteReader();
            or106.Read();
            String s106 = or106[0].ToString();

            string sql107 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 6) and (Немісцевий= true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc107 = new OleDbCommand(sql107, connection);

            OleDbDataReader or107 = dc107.ExecuteReader();
            or107.Read();
            String s107 = or107[0].ToString();


            string sql108 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 7) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc108 = new OleDbCommand(sql108, connection);

            OleDbDataReader or108 = dc108.ExecuteReader();
            or108.Read();
            String s108 = or108[0].ToString();

            string sql109 = "Select COUNT(*) FROM Головна where ((Зараховано = true) and (Спеціальність = 8) and (Немісцевий = true ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc109 = new OleDbCommand(sql109, connection);

            OleDbDataReader or109 = dc109.ExecuteReader();
            or109.Read();
            String s109 = or109[0].ToString();
            //немісцеві конец






            //незадовыльны оцінки
            string sql111 = "Select COUNT(*) FROM Головна where (  ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ) and ([Бал за результатами вступних екзаменів: математика] <> 0))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc111 = new OleDbCommand(sql111, connection);

            OleDbDataReader or111 = dc111.ExecuteReader();
            or111.Read();
            String s111 = or111[0].ToString();

            string sql112 = "Select COUNT(*) FROM Головна where ((Спеціальність = 1) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc112 = new OleDbCommand(sql112, connection);

            OleDbDataReader or112 = dc112.ExecuteReader();
            or112.Read();
            String s112 = or112[0].ToString();

            string sql113 = "Select COUNT(*) FROM Головна where ((Спеціальність = 5) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc113 = new OleDbCommand(sql113, connection);

            OleDbDataReader or113 = dc113.ExecuteReader();
            or113.Read();
            String s113 = or113[0].ToString();

            string sql114 = "Select COUNT(*) FROM Головна where ( (Спеціальність = 4) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc114 = new OleDbCommand(sql114, connection);

            OleDbDataReader or114 = dc114.ExecuteReader();
            or114.Read();
            String s114 = or114[0].ToString();

            string sql115 = "Select COUNT(*) FROM Головна where ( (Спеціальність = 2) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc115 = new OleDbCommand(sql115, connection);

            OleDbDataReader or115 = dc115.ExecuteReader();
            or115.Read();
            String s115 = or115[0].ToString();


            string sql116 = "Select COUNT(*) FROM Головна where ( (Спеціальність = 3) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc116 = new OleDbCommand(sql116, connection);

            OleDbDataReader or116 = dc116.ExecuteReader();
            or116.Read();
            String s116 = or116[0].ToString();

            string sql117 = "Select COUNT(*) FROM Головна where ( (Спеціальність = 6) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: рисунок] < 4) or ([Бал за результатами вступних екзаменів: композиція] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc117 = new OleDbCommand(sql117, connection);

            OleDbDataReader or117 = dc117.ExecuteReader();
            or117.Read();
            String s117 = or117[0].ToString();


            string sql118 = "Select COUNT(*) FROM Головна where ( (Спеціальність = 7) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc118 = new OleDbCommand(sql118, connection);

            OleDbDataReader or118 = dc118.ExecuteReader();
            or118.Read();
            String s118 = or118[0].ToString();

            string sql119 = "Select COUNT(*) FROM Головна where ( (Спеціальність = 8) and ( ([Бал за результатами вступних екзаменів: українська мова] < 4 ) or ([Бал за результатами вступних екзаменів: математика] < 4) ))" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc119 = new OleDbCommand(sql119, connection);

            OleDbDataReader or119 = dc119.ExecuteReader();
            or119.Read();
            String s119 = or119[0].ToString();



            connection.Close();

            //String s = or[0].ToString().Substring(0, or[0].ToString().IndexOf(" "));
            //String s2 = or[0].ToString().Substring(or[0].ToString().IndexOf(" "), or[0].ToString().LastIndexOf(" ") - or[0].ToString().IndexOf(" "));
            //String s3 = or[0].ToString().Substring(or[0].ToString().LastIndexOf(" "), or[0].ToString().Length - or[0].ToString().LastIndexOf(" "));



            Excel.Workbook xlWorkBook;

            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
    "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
     Type.Missing, Type.Missing); ;
            xlWorkSheet = xlWorkBook.Worksheets.Item["zvit1"];
            xlWorkSheet.Activate();

            excelcells = xlWorkSheet.get_Range("D6");
            excelcells.Value2 = s1;
            excelcells = xlWorkSheet.get_Range("F6");
            excelcells.Value2 = s2;
            excelcells = xlWorkSheet.get_Range("H6");
            excelcells.Value2 = s3;
            excelcells = xlWorkSheet.get_Range("J6");
            excelcells.Value2 = s4;
            excelcells = xlWorkSheet.get_Range("L6");
            excelcells.Value2 = s5;
            excelcells = xlWorkSheet.get_Range("N6");
            excelcells.Value2 = s6;
            excelcells = xlWorkSheet.get_Range("P6");
            excelcells.Value2 = s7;
            excelcells = xlWorkSheet.get_Range("R6");
            excelcells.Value2 = s8;
            excelcells = xlWorkSheet.get_Range("T6");
            excelcells.Value2 = s9;

            excelcells = xlWorkSheet.get_Range("D12");
            excelcells.Value2 = s11;
            excelcells = xlWorkSheet.get_Range("F12");
            excelcells.Value2 = s12;
            excelcells = xlWorkSheet.get_Range("H12");
            excelcells.Value2 = s13;
            excelcells = xlWorkSheet.get_Range("J12");
            excelcells.Value2 = s14;
            excelcells = xlWorkSheet.get_Range("L12");
            excelcells.Value2 = s15;
            excelcells = xlWorkSheet.get_Range("N12");
            excelcells.Value2 = s16;
            excelcells = xlWorkSheet.get_Range("P12");
            excelcells.Value2 = s17;
            excelcells = xlWorkSheet.get_Range("R12");
            excelcells.Value2 = s18;
            excelcells = xlWorkSheet.get_Range("T12");
            excelcells.Value2 = s19;

            excelcells = xlWorkSheet.get_Range("D18");
            excelcells.Value2 = s21;
            excelcells = xlWorkSheet.get_Range("F18");
            excelcells.Value2 = s22;
            excelcells = xlWorkSheet.get_Range("H18");
            excelcells.Value2 = s23;
            excelcells = xlWorkSheet.get_Range("J18");
            excelcells.Value2 = s24;
            excelcells = xlWorkSheet.get_Range("L18");
            excelcells.Value2 = s25;
            excelcells = xlWorkSheet.get_Range("N18");
            excelcells.Value2 = s26;
            excelcells = xlWorkSheet.get_Range("P18");
            excelcells.Value2 = s27;
            excelcells = xlWorkSheet.get_Range("R18");
            excelcells.Value2 = s28;
            excelcells = xlWorkSheet.get_Range("T18");
            excelcells.Value2 = s29;

            excelcells = xlWorkSheet.get_Range("D21");
            excelcells.Value2 = s31;
            excelcells = xlWorkSheet.get_Range("F21");
            excelcells.Value2 = s32;
            excelcells = xlWorkSheet.get_Range("H21");
            excelcells.Value2 = s33;
            excelcells = xlWorkSheet.get_Range("J21");
            excelcells.Value2 = s34;
            excelcells = xlWorkSheet.get_Range("L21");
            excelcells.Value2 = s35;
            excelcells = xlWorkSheet.get_Range("N21");
            excelcells.Value2 = s36;
            excelcells = xlWorkSheet.get_Range("P21");
            excelcells.Value2 = s37;
            excelcells = xlWorkSheet.get_Range("R21");
            excelcells.Value2 = s38;
            excelcells = xlWorkSheet.get_Range("T21");
            excelcells.Value2 = s39;


            excelcells = xlWorkSheet.get_Range("D24");
            excelcells.Value2 = s41;
            excelcells = xlWorkSheet.get_Range("F24");
            excelcells.Value2 = s42;
            excelcells = xlWorkSheet.get_Range("H24");
            excelcells.Value2 = s43;
            excelcells = xlWorkSheet.get_Range("J24");
            excelcells.Value2 = s44;
            excelcells = xlWorkSheet.get_Range("L24");
            excelcells.Value2 = s45;
            excelcells = xlWorkSheet.get_Range("N24");
            excelcells.Value2 = s46;
            excelcells = xlWorkSheet.get_Range("P24");
            excelcells.Value2 = s47;
            excelcells = xlWorkSheet.get_Range("R24");
            excelcells.Value2 = s48;
            excelcells = xlWorkSheet.get_Range("T24");
            excelcells.Value2 = s49;


            excelcells = xlWorkSheet.get_Range("D25");
            excelcells.Value2 = s51;
            excelcells = xlWorkSheet.get_Range("F25");
            excelcells.Value2 = s52;
            excelcells = xlWorkSheet.get_Range("H25");
            excelcells.Value2 = s53;
            excelcells = xlWorkSheet.get_Range("J25");
            excelcells.Value2 = s54;
            excelcells = xlWorkSheet.get_Range("L25");
            excelcells.Value2 = s55;
            excelcells = xlWorkSheet.get_Range("N25");
            excelcells.Value2 = s56;
            excelcells = xlWorkSheet.get_Range("P25");
            excelcells.Value2 = s57;
            excelcells = xlWorkSheet.get_Range("R25");
            excelcells.Value2 = s58;
            excelcells = xlWorkSheet.get_Range("T25");
            excelcells.Value2 = s59;

            excelcells = xlWorkSheet.get_Range("D26");
            excelcells.Value2 = s511;
            excelcells = xlWorkSheet.get_Range("F26");
            excelcells.Value2 = s521;
            excelcells = xlWorkSheet.get_Range("H26");
            excelcells.Value2 = s531;
            excelcells = xlWorkSheet.get_Range("J26");
            excelcells.Value2 = s541;
            excelcells = xlWorkSheet.get_Range("L26");
            excelcells.Value2 = s551;
            excelcells = xlWorkSheet.get_Range("N26");
            excelcells.Value2 = s561;
            excelcells = xlWorkSheet.get_Range("P26");
            excelcells.Value2 = s571;
            excelcells = xlWorkSheet.get_Range("R26");
            excelcells.Value2 = s581;
            excelcells = xlWorkSheet.get_Range("T26");
            excelcells.Value2 = s591;


            excelcells = xlWorkSheet.get_Range("D27");
            excelcells.Value2 = s512;
            excelcells = xlWorkSheet.get_Range("F27");
            excelcells.Value2 = s522;
            excelcells = xlWorkSheet.get_Range("H27");
            excelcells.Value2 = s532;
            excelcells = xlWorkSheet.get_Range("J27");
            excelcells.Value2 = s542;
            excelcells = xlWorkSheet.get_Range("L27");
            excelcells.Value2 = s552;
            excelcells = xlWorkSheet.get_Range("N27");
            excelcells.Value2 = s562;
            excelcells = xlWorkSheet.get_Range("P27");
            excelcells.Value2 = s572;
            excelcells = xlWorkSheet.get_Range("R27");
            excelcells.Value2 = s582;
            excelcells = xlWorkSheet.get_Range("T27");
            excelcells.Value2 = s592;


            excelcells = xlWorkSheet.get_Range("D28");
            excelcells.Value2 = s513;
            excelcells = xlWorkSheet.get_Range("F28");
            excelcells.Value2 = s523;
            excelcells = xlWorkSheet.get_Range("H28");
            excelcells.Value2 = s533;
            excelcells = xlWorkSheet.get_Range("J28");
            excelcells.Value2 = s543;
            excelcells = xlWorkSheet.get_Range("L28");
            excelcells.Value2 = s553;
            excelcells = xlWorkSheet.get_Range("N28");
            excelcells.Value2 = s563;
            excelcells = xlWorkSheet.get_Range("P28");
            excelcells.Value2 = s573;
            excelcells = xlWorkSheet.get_Range("R28");
            excelcells.Value2 = s583;
            excelcells = xlWorkSheet.get_Range("T28");
            excelcells.Value2 = s593;

            excelcells = xlWorkSheet.get_Range("D29");
            excelcells.Value2 = s514;
            excelcells = xlWorkSheet.get_Range("F29");
            excelcells.Value2 = s524;
            excelcells = xlWorkSheet.get_Range("H29");
            excelcells.Value2 = s534;
            excelcells = xlWorkSheet.get_Range("J29");
            excelcells.Value2 = s544;
            excelcells = xlWorkSheet.get_Range("L29");
            excelcells.Value2 = s554;
            excelcells = xlWorkSheet.get_Range("N29");
            excelcells.Value2 = s564;
            excelcells = xlWorkSheet.get_Range("P29");
            excelcells.Value2 = s574;
            excelcells = xlWorkSheet.get_Range("R29");
            excelcells.Value2 = s584;
            excelcells = xlWorkSheet.get_Range("T29");
            excelcells.Value2 = s594;

            excelcells = xlWorkSheet.get_Range("D30");
            excelcells.Value2 = s61;
            excelcells = xlWorkSheet.get_Range("F30");
            excelcells.Value2 = s62;
            excelcells = xlWorkSheet.get_Range("H30");
            excelcells.Value2 = s63;
            excelcells = xlWorkSheet.get_Range("J30");
            excelcells.Value2 = s64;
            excelcells = xlWorkSheet.get_Range("L30");
            excelcells.Value2 = s65;
            excelcells = xlWorkSheet.get_Range("N30");
            excelcells.Value2 = s66;
            excelcells = xlWorkSheet.get_Range("P30");
            excelcells.Value2 = s67;
            excelcells = xlWorkSheet.get_Range("R30");
            excelcells.Value2 = s68;
            excelcells = xlWorkSheet.get_Range("T30");
            excelcells.Value2 = s69;


            excelcells = xlWorkSheet.get_Range("D31");
            excelcells.Value2 = s11;
            excelcells = xlWorkSheet.get_Range("F31");
            excelcells.Value2 = s12;
            excelcells = xlWorkSheet.get_Range("H31");
            excelcells.Value2 = s13;
            excelcells = xlWorkSheet.get_Range("J31");
            excelcells.Value2 = s14;
            excelcells = xlWorkSheet.get_Range("L31");
            excelcells.Value2 = s15;
            excelcells = xlWorkSheet.get_Range("N31");
            excelcells.Value2 = s16;
            excelcells = xlWorkSheet.get_Range("P31");
            excelcells.Value2 = s17;
            excelcells = xlWorkSheet.get_Range("R31");
            excelcells.Value2 = s18;
            excelcells = xlWorkSheet.get_Range("T31");
            excelcells.Value2 = s19;
            xlApp.SaveWorkspace();
            xlApp.Visible = true;

            excelcells = xlWorkSheet.get_Range("D32");
            excelcells.Value2 = s71;
            excelcells = xlWorkSheet.get_Range("F32");
            excelcells.Value2 = s72;
            excelcells = xlWorkSheet.get_Range("H32");
            excelcells.Value2 = s73;
            excelcells = xlWorkSheet.get_Range("J32");
            excelcells.Value2 = s74;
            excelcells = xlWorkSheet.get_Range("L32");
            excelcells.Value2 = s75;
            excelcells = xlWorkSheet.get_Range("N32");
            excelcells.Value2 = s76;
            excelcells = xlWorkSheet.get_Range("P32");
            excelcells.Value2 = s77;
            excelcells = xlWorkSheet.get_Range("R32");
            excelcells.Value2 = s78;
            excelcells = xlWorkSheet.get_Range("T32");
            excelcells.Value2 = s79;

            excelcells = xlWorkSheet.get_Range("D33");
            excelcells.Value2 = s81;
            excelcells = xlWorkSheet.get_Range("F33");
            excelcells.Value2 = s82;
            excelcells = xlWorkSheet.get_Range("H33");
            excelcells.Value2 = s83;
            excelcells = xlWorkSheet.get_Range("J33");
            excelcells.Value2 = s84;
            excelcells = xlWorkSheet.get_Range("L33");
            excelcells.Value2 = s85;
            excelcells = xlWorkSheet.get_Range("N33");
            excelcells.Value2 = s86;
            excelcells = xlWorkSheet.get_Range("P33");
            excelcells.Value2 = s87;
            excelcells = xlWorkSheet.get_Range("R33");
            excelcells.Value2 = s88;
            excelcells = xlWorkSheet.get_Range("T33");
            excelcells.Value2 = s89;

            excelcells = xlWorkSheet.get_Range("D34");
            excelcells.Value2 = s91;
            excelcells = xlWorkSheet.get_Range("F34");
            excelcells.Value2 = s92;
            excelcells = xlWorkSheet.get_Range("H34");
            excelcells.Value2 = s93;
            excelcells = xlWorkSheet.get_Range("J34");
            excelcells.Value2 = s94;
            excelcells = xlWorkSheet.get_Range("L34");
            excelcells.Value2 = s95;
            excelcells = xlWorkSheet.get_Range("N34");
            excelcells.Value2 = s96;
            excelcells = xlWorkSheet.get_Range("P34");
            excelcells.Value2 = s97;
            excelcells = xlWorkSheet.get_Range("R34");
            excelcells.Value2 = s98;
            excelcells = xlWorkSheet.get_Range("T34");
            excelcells.Value2 = s99;

            excelcells = xlWorkSheet.get_Range("D35");
            excelcells.Value2 = s911;
            excelcells = xlWorkSheet.get_Range("F35");
            excelcells.Value2 = s921;
            excelcells = xlWorkSheet.get_Range("H35");
            excelcells.Value2 = s931;
            excelcells = xlWorkSheet.get_Range("J35");
            excelcells.Value2 = s941;
            excelcells = xlWorkSheet.get_Range("L35");
            excelcells.Value2 = s951;
            excelcells = xlWorkSheet.get_Range("N35");
            excelcells.Value2 = s961;
            excelcells = xlWorkSheet.get_Range("P35");
            excelcells.Value2 = s971;
            excelcells = xlWorkSheet.get_Range("R35");
            excelcells.Value2 = s981;
            excelcells = xlWorkSheet.get_Range("T35");
            excelcells.Value2 = s991;


            excelcells = xlWorkSheet.get_Range("D36");
            excelcells.Value2 = s912;
            excelcells = xlWorkSheet.get_Range("F36");
            excelcells.Value2 = s922;
            excelcells = xlWorkSheet.get_Range("H36");
            excelcells.Value2 = s932;
            excelcells = xlWorkSheet.get_Range("J36");
            excelcells.Value2 = s942;
            excelcells = xlWorkSheet.get_Range("L36");
            excelcells.Value2 = s952;
            excelcells = xlWorkSheet.get_Range("N36");
            excelcells.Value2 = s962;
            excelcells = xlWorkSheet.get_Range("P36");
            excelcells.Value2 = s972;
            excelcells = xlWorkSheet.get_Range("R36");
            excelcells.Value2 = s982;
            excelcells = xlWorkSheet.get_Range("T36");
            excelcells.Value2 = s992;

            excelcells = xlWorkSheet.get_Range("D38");
            excelcells.Value2 = s101;
            excelcells = xlWorkSheet.get_Range("F38");
            excelcells.Value2 = s102;
            excelcells = xlWorkSheet.get_Range("H38");
            excelcells.Value2 = s103;
            excelcells = xlWorkSheet.get_Range("J38");
            excelcells.Value2 = s104;
            excelcells = xlWorkSheet.get_Range("L38");
            excelcells.Value2 = s105;
            excelcells = xlWorkSheet.get_Range("N38");
            excelcells.Value2 = s106;
            excelcells = xlWorkSheet.get_Range("P38");
            excelcells.Value2 = s107;
            excelcells = xlWorkSheet.get_Range("R38");
            excelcells.Value2 = s108;
            excelcells = xlWorkSheet.get_Range("T38");
            excelcells.Value2 = s109;


            excelcells = xlWorkSheet.get_Range("D39");
            excelcells.Value2 = s111;
            excelcells = xlWorkSheet.get_Range("F39");
            excelcells.Value2 = s112;
            excelcells = xlWorkSheet.get_Range("H39");
            excelcells.Value2 = s113;
            excelcells = xlWorkSheet.get_Range("J39");
            excelcells.Value2 = s114;
            excelcells = xlWorkSheet.get_Range("L39");
            excelcells.Value2 = s115;
            excelcells = xlWorkSheet.get_Range("N39");
            excelcells.Value2 = s116;
            excelcells = xlWorkSheet.get_Range("P39");
            excelcells.Value2 = s117;
            excelcells = xlWorkSheet.get_Range("R39");
            excelcells.Value2 = s118;
            excelcells = xlWorkSheet.get_Range("T39");
            excelcells.Value2 = s119;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            string lZvit = CopyExcelDocFromTemplate();
            connection.Close();

            int count1 = 7;

            Excel.Workbook xlWorkBook;

            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
    "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
     Type.Missing, Type.Missing, Type.Missing, Type.Missing,
     Type.Missing, Type.Missing); ;
            xlWorkSheet = xlWorkBook.Worksheets.Item["zvit2"];
            xlWorkSheet.Activate();



            string sql1 = "Select ПІБ, [Документ про освіту: серія], [Документ про освіту: номер], [Бал за результатами вступних екзаменів: українська мова], [Бал за результатами вступних екзаменів: математика], [Бал за підготовчі курси], [Шкільний середній бал], [Сума балів (прохідний бал)] , Спеціальність FROM Головна where ( Зараховано = true)" + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) ";
            OleDbCommand dc1 = new OleDbCommand(sql1, connection);
            connection.Open();
            OleDbDataReader or1 = dc1.ExecuteReader();

            int i = 1;
            while (or1.Read())
            {
                string[] FIO = or1[0].ToString().Split(' ');
                String s1 = FIO.First();
                String s2 = FIO[1];
                String s3 = FIO.Last();
                String s4 = or1[1].ToString();
                String s5 = or1[2].ToString();
                String s6 = or1[3].ToString();
                String s7 = or1[4].ToString();
                String s8 = or1[5].ToString();
                String s9 = or1[6].ToString();
                String s10 = or1[7].ToString();
                string s11 = "";
                if (Convert.ToInt32(or1[8].ToString()) == 1)
                {
                    s11 = "Монтаж і експлуатація електроустаткування підприємств і цивільних споруд";
                }
                if (Convert.ToInt32(or1[8].ToString()) == 2)
                {
                    s11 = "Технологія обробки матеріалів на верстатах і автоматичних лініях";
                }
                if (Convert.ToInt32(or1[8].ToString()) == 3)
                {
                    s11 = "Розробка програмного забезпечення";
                }
                if (Convert.ToInt32(or1[8].ToString()) == 4)
                {
                    s11 = "Виробництво електричних машин і апаратів";
                }
                if (Convert.ToInt32(or1[8].ToString()) == 5)
                {
                    s11 = "Конструювання, виробництво і технічне обслуговування виробів електронної техніки";
                }
                if (Convert.ToInt32(or1[8].ToString()) == 6)
                {
                    s11 = "Дизайн";
                }
                if (Convert.ToInt32(or1[8].ToString()) == 7)
                {
                    s11 = "Економіка підприємства(на базі 9 класів)";

                }
                if (Convert.ToInt32(or1[8].ToString()) == 8)
                {
                    s11 = "Економіка підприємства(на базі 11 класів)";

                }
                excelcells = xlWorkSheet.get_Range("A" + count1);
                excelcells.Value2 = i;
                excelcells = xlWorkSheet.get_Range("B" + count1);
                excelcells.Value2 = s11;
                excelcells = xlWorkSheet.get_Range("C" + count1);
                excelcells.Value2 = s1;
                excelcells = xlWorkSheet.get_Range("D" + count1);
                excelcells.Value2 = s2;
                excelcells = xlWorkSheet.get_Range("E" + count1);
                excelcells.Value2 = s3;
                excelcells = xlWorkSheet.get_Range("F" + count1);
                excelcells.Value2 = s4;
                excelcells = xlWorkSheet.get_Range("G" + count1);
                excelcells.Value2 = s5;
                excelcells = xlWorkSheet.get_Range("K" + count1);
                excelcells.Value2 = s6;
                excelcells = xlWorkSheet.get_Range("P" + count1);
                excelcells.Value2 = s7;
                excelcells = xlWorkSheet.get_Range("X" + count1);
                excelcells.Value2 = s8;
                excelcells = xlWorkSheet.get_Range("Y" + count1);
                excelcells.Value2 = s9;
                excelcells = xlWorkSheet.get_Range("Z" + count1);
                excelcells.Value2 = s10;
                count1++;
                i++;
            }

            xlApp.SaveWorkspace();
            xlApp.Visible = true;
        }


        private void button7_Click(object sender, EventArgs e)
        {
            string lZvit = CopyExcelDocFromTemplate();


            int count1 = 4;
            int count2 = 4;
            int count3 = 4;
            int count4 = 4;
            int count5 = 4;
            int count6 = 4;
            int count7 = 4;
            int count8 = 4;
            int count9 = 4;
            int count10 = 4;
            int count11 = 4;
            int count12 = 4;
            int count13 = 4;
            int count14 = 4;
            int count15 = 4;
            int count16 = 4;
            connection.Close();
            
            // string sql2= "Select count(ПІБ) FROM Головна where (( Зараховано = true) and (Спеціальність = 1))";
            // OleDbCommand dc2 = new OleDbCommand(sql2, connection);
            // connection.Open();
            // OleDbDataReader or2 = dc2.ExecuteReader();
            // or2.Read();
            //int s2= Convert.ToInt32(or2[0].ToString());
            connection.Open();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            if (lcbSpec.comboBox.Text == "ДЗ")
            { xlWorkSheet = xlWorkBook.Worksheets.Item["zvit4_dz"]; }
            else
            { xlWorkSheet = xlWorkBook.Worksheets.Item["zvit4"]; }
            xlWorkSheet.Activate();
            string show_deleted;
            if (cbShowDeleted.Checked==true)
            {
                show_deleted = " ";
                //show_deleted = " AND [Забрав документи]=true ";
            }
            else
            {
                show_deleted = " AND [Забрав документи]=false ";
            }
         



            string sql1 = "Select ПІБ FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() +show_deleted+ ") AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc1 = new OleDbCommand(sql1, connection);

            OleDbDataReader or1 = dc1.ExecuteReader();

            string sql2 = "Select [№ реєстрації] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc2 = new OleDbCommand(sql2, connection);

            OleDbDataReader or2 = dc2.ExecuteReader();

            string sql3 = "Select [Бал за результатами вступних екзаменів: математика] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc3 = new OleDbCommand(sql3, connection);

            OleDbDataReader or3 = dc3.ExecuteReader();

            string sql4 = "Select [Бал за результатами вступних екзаменів: українська мова] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + " AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc4 = new OleDbCommand(sql4, connection);

            OleDbDataReader or4 = dc4.ExecuteReader();

            string sql5 = "Select [Шкільний середній бал] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc5 = new OleDbCommand(sql5, connection);

            OleDbDataReader or5 = dc5.ExecuteReader();




            string sql6 = "Select [Бал за підготовчі курси] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc6 = new OleDbCommand(sql6, connection);

            OleDbDataReader or6 = dc6.ExecuteReader();



            string sql7 = "Select [Сума балів (прохідний бал)] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc7 = new OleDbCommand(sql7, connection);

            OleDbDataReader or7 = dc7.ExecuteReader();

            string sql8 = "Select Пільги FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc8 = new OleDbCommand(sql8, connection);

            OleDbDataReader or8 = dc8.ExecuteReader();



            string sql9 = "Select [Закінчили підготовчі курси] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc9 = new OleDbCommand(sql9, connection);

            OleDbDataReader or9 = dc9.ExecuteReader();
            //   or1.Read();


            string sql10 = "Select [Друга спеціальність] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dc10 = new OleDbCommand(sql10, connection);

            OleDbDataReader or10 = dc10.ExecuteReader();

            string sqlList = "Select [Лист від підприємства] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + ")AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
            "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#))";
            OleDbCommand dcList = new OleDbCommand(sqlList, connection);

            OleDbDataReader orList = dcList.ExecuteReader();


            string sql11 = "Select [Бал за результатами вступних екзаменів: рисунок] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc11 = new OleDbCommand(sql11, connection);
            OleDbDataReader or11 = dc11.ExecuteReader();


            string sql12 = "Select [Бал за результатами вступних екзаменів: композиція] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc12 = new OleDbCommand(sql12, connection);
            OleDbDataReader or12 = dc12.ExecuteReader();

            string sql13 = "Select [Шкільна оцінка по алгебрі] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                  "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc13 = new OleDbCommand(sql13, connection);
            OleDbDataReader or13 = dc13.ExecuteReader();

            string sql14 = "Select [Шкільна оцінка по геометрії] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                  "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc14 = new OleDbCommand(sql14, connection);
            OleDbDataReader or14 = dc14.ExecuteReader();

            string sql15 = "Select [Шкільна оцінка по українській мові] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                  "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc15 = new OleDbCommand(sql15, connection);
            OleDbDataReader or15 = dc15.ExecuteReader();

            string sql16 = "Select [Документи],[Забрав документи] FROM Головна where ((Спеціальність = " + lcbSpec.keyValue.ToString() + show_deleted + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                 "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() /*+ dtpCreatedTill.Value.Date.ToString("dd/MM/yyyy")*/ + "#)))";
            OleDbCommand dc16 = new OleDbCommand(sql16, connection);
            OleDbDataReader or16 = dc16.ExecuteReader();

           
            string sql17 = GetMainSql()
                // + "   AND  Головна.[Зараховано] = true "
                + "   AND  Головна.[Забрав документи] = true "
                 + "   AND  Головна.[Потребує гуртожитку] = true "
                  + " AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                 "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)"
                 + " ORDER BY val(Головна.[Сума балів (прохідний бал)])  DESC ";
            
            OleDbCommand dc17 = new OleDbCommand(sql17, connection);
            OleDbDataReader or17 = dc17.ExecuteReader();

            excelcells = xlWorkSheet.get_Range("D1");
            excelcells.Value2 = "\"" + lcbSpec.comboBox.Text + "\"";
           


            /* while (orList.Read())
             {
                 String sList = orList["Лист від підприємства"].ToString();

                 excelcells = xlWorkSheet.get_Range("L" + count3);
                 excelcells.Value2 = sList;

                 count3++;
             }
             count3 = count1;
         */        

            while (or1.Read())
            {
                String s1 = or1["ПІБ"].ToString();

                excelcells = xlWorkSheet.get_Range("C" + count1);
                excelcells.Value2 = s1;
                excelcells = xlWorkSheet.get_Range("A" + count1);
                excelcells.Value2 = count1 - 3;

               
                // rang++;
                count1++;
            }


            while (or2.Read())
            {
                String s2 = or2["№ реєстрації"].ToString();

                excelcells = xlWorkSheet.get_Range("B" + count2);
                excelcells.Value2 = s2;
                // rang++;
                count2++;
            }

            while (or13.Read())
            {
                String s13 = or13["Шкільна оцінка по алгебрі"].ToString();
                excelcells = xlWorkSheet.get_Range("D" + count13);
                excelcells.Value2 = s13;
                count13++;
            }

            while (or14.Read())
            {
                String s14 = or14["Шкільна оцінка по геометрії"].ToString();
                excelcells = xlWorkSheet.get_Range("E" + count14);
                excelcells.Value2 = s14;
                count14++;
            }

            while (or15.Read())
            {
                String s15 = or15["Шкільна оцінка по українській мові"].ToString();
                excelcells = xlWorkSheet.get_Range("F" + count15);
                excelcells.Value2 = s15;
                count15++;
            }


            while (or3.Read())
            {
                String s3;
                if (lcbSpec.comboBox.Text == "ДЗ")
                { }
                else
                {
                    s3 = or3["Бал за результатами вступних екзаменів: математика"].ToString();
                    excelcells = xlWorkSheet.get_Range("H" + count3);
                    excelcells.Value2 = s3;
                    count3++;
                }


            }

            while (or4.Read())
            {
                if (lcbSpec.comboBox.Text == "ДЗ")
                {
                    String s4 = or4["Бал за результатами вступних екзаменів: українська мова"].ToString();
                    excelcells = xlWorkSheet.get_Range("H" + count4);
                    excelcells.Value2 = s4;
                    count4++;
                }
                else
                {
                    String s4 = or4["Бал за результатами вступних екзаменів: українська мова"].ToString();
                    excelcells = xlWorkSheet.get_Range("I" + count4);
                    excelcells.Value2 = s4;
                    count4++;
                }
            }
            if (lcbSpec.comboBox.Text == "ДЗ")
            {
                while (or11.Read())
                {
                    String s11 = or11["Бал за результатами вступних екзаменів: рисунок"].ToString();

                    excelcells = xlWorkSheet.get_Range("I" + count11);
                    excelcells.Value2 = s11;

                    count11++;
                }

                while (or12.Read())
                {
                    String s12 = or12["Бал за результатами вступних екзаменів: композиція"].ToString();

                    excelcells = xlWorkSheet.get_Range("J" + count12);
                    excelcells.Value2 = s12;

                    count12++;
                }
            }

            while (or5.Read())
            {
                String s5 = or5["Шкільний середній бал"].ToString();
                excelcells = xlWorkSheet.get_Range("G" + count5);
                excelcells.Value2 = s5;
                count5++;
            }

            while (or6.Read())
            {
                String s6 = or6["Бал за підготовчі курси"].ToString();
                if (lcbSpec.comboBox.Text == "ДЗ")
                {
                    excelcells = xlWorkSheet.get_Range("K" + count6);
                }
                else
                {
                    excelcells = xlWorkSheet.get_Range("J" + count6);
                }

                excelcells.Value2 = s6;
                count6++;
            }

            while (or7.Read())
            {
                if (lcbSpec.comboBox.Text == "ДЗ")
                {
                    String s7 = or7["Сума балів (прохідний бал)"].ToString();
                    excelcells = xlWorkSheet.get_Range("L" + count7);
                    excelcells.Value2 = Double.Parse(s7);
                    count7++;
                }
                else
                {
                    String s7 = or7["Сума балів (прохідний бал)"].ToString();
                    excelcells = xlWorkSheet.get_Range("K" + count7);
                    excelcells.Value2 = Double.Parse(s7);
                    count7++;
                }
            }


            while (or8.Read())
            {
                String s8 = or8["Пільги"].ToString();
                if (lcbSpec.comboBox.Text == "ДЗ")
                { excelcells = xlWorkSheet.get_Range("M" + count8); }
                else
                {
                    excelcells = xlWorkSheet.get_Range("L" + count8);
                }
                excelcells.Value2 = s8;
                count8++;
            }




            while (or9.Read())
            {

                String s9 = or9["Закінчили підготовчі курси"].ToString();
                if (s9 == "1")
                {
                    s9 = "-";
                }

                if (s9 == "2")
                {
                    s9 = "1";
                }
                if (s9 == "3")
                {
                    s9 = "3";
                }
                if (s9 == "4")
                {
                    s9 = "6";
                }
                if (lcbSpec.comboBox.Text == "ДЗ")
                {
                    excelcells = xlWorkSheet.get_Range("N" + count9);
                }
                else
                {
                    excelcells = xlWorkSheet.get_Range("M" + count9);
                }

                excelcells.Value2 = s9;
                count9++;
            }

            while (or10.Read())
            {

                String s10 = or10["Друга спеціальність"].ToString();
                if (s10 == "1")
                {
                    s10 = "МЕУб";
                }

                if (s10 == "2")
                {
                    s10 = "ТОМб";
                }
                if (s10 == "3")
                {
                    s10 = "РПЗб";
                }
                if (s10 == "4")
                {
                    s10 = "ЕМАб";
                }
                if (s10 == "5")
                {
                    s10 = "КВЕТб";
                }
                if (s10 == "6")
                {
                    s10 = "ДЗб";
                }
                if (s10 == "7")
                {
                    s10 = "ЕП9б";
                }
                if (s10 == "8")
                {
                    s10 = "ЕП11б";
                }
                if (s10 == "11")
                {
                    s10 = "-";
                }
                if (lcbSpec.comboBox.Text == "ДЗ")
                { excelcells = xlWorkSheet.get_Range("O" + count10); }
                else
                { excelcells = xlWorkSheet.get_Range("N" + count10); }
                excelcells.Value2 = s10;

                count10++;
            }

            while (or16.Read())
            {
                String s16 = or16["Документи"].ToString();
                if (lcbSpec.comboBox.Text == "ДЗ")
                { excelcells = xlWorkSheet.get_Range("P" + count16); }
                else
                {
                    excelcells = xlWorkSheet.get_Range("O" + count16);
                }
                if (cbShowDeleted.Checked==true)
                {
                        if ((Boolean)or16["Забрав документи"] == true)
                    {
                        excelcells.EntireRow.Font.ColorIndex = 3;
                        excelcells.EntireRow.Font.Bold = true;
                    }
                }
                


                excelcells.Value2 = s16;
                count16++;

            }
          /*  while (or17.Read())
            {
                if ((Boolean)or17["Забрав документи"] == true)
                {
                    excelcells.EntireRow.Font.ColorIndex = 3;
                    excelcells.EntireRow.Font.Bold = true;
                }          
            }
            */



            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            if (lcbSpec.comboBox.Text == "ДЗ")
            {
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Item["zvit4_dz"];
                Excel.Range tempRange = ObjWorkSheet.get_Range("A3", "P1000");
                tempRange.Sort(tempRange.Columns[12, Type.Missing],
                      Excel.XlSortOrder.xlDescending,
                      Type.Missing, Type.Missing,
                      Excel.XlSortOrder.xlDescending,
                      Type.Missing,
                      Excel.XlSortOrder.xlDescending,
                      Excel.XlYesNoGuess.xlYes,
                      Type.Missing,
                      Type.Missing,
                      Excel.XlSortOrientation.xlSortColumns,
                      Excel.XlSortMethod.xlPinYin,
                      Excel.XlSortDataOption.xlSortNormal,
                      Excel.XlSortDataOption.xlSortNormal,
                      Excel.XlSortDataOption.xlSortNormal);
            }
            else
            {
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Item["zvit4"];
                Excel.Range tempRange = ObjWorkSheet.get_Range("A3", "O1000");
                tempRange.Sort(tempRange.Columns[11, Type.Missing],
                      Excel.XlSortOrder.xlDescending,
                      Type.Missing, Type.Missing,
                      Excel.XlSortOrder.xlDescending,
                      Type.Missing,
                      Excel.XlSortOrder.xlDescending,
                      Excel.XlYesNoGuess.xlYes,
                      Type.Missing,
                      Type.Missing,
                      Excel.XlSortOrientation.xlSortColumns,
                      Excel.XlSortMethod.xlPinYin,
                      Excel.XlSortDataOption.xlSortNormal,
                      Excel.XlSortDataOption.xlSortNormal,
                      Excel.XlSortDataOption.xlSortNormal);
            }

            xlApp.SaveWorkspace();
            xlApp.Visible = true;
            connection.Close();
        }


        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {


            if (Char.IsNumber(e.KeyChar) || e.KeyChar == '\b')
            {
                //MessageBox.Show("ввод только цифр");
                //  ErrorProvider.SetError(Control, "Ололо ошибка");

            }
            else
            {
                e.Handled = true;
            }

        }

        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {

            if (!Char.IsDigit(e.KeyChar))
            {

                if (e.KeyChar != '.' || e.KeyChar != (Char)ConsoleKey.Backspace || textBox1.Text.IndexOf(".") != -1)
                {
                    e.Handled = true;
                }
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            {
                connection.Close();
                viknoAbitur da = new viknoAbitur();
                if (da.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string sql =
                        "Insert into Головна " +
                        "( [№ реєстрації], " +//1
                        "  ПІБ, " +//2
                        "  Спеціальність, " +//2+
                        "  [Шкільна оцінка по алгебрі], " +//3
                        "  [Шкільна оцінка по геометрії], " +//4
                        "  [Шкільна оцінка по українській мові], " +//5
                        "  [Шкільний середній бал], " +//6
                        "  [Бал за результатами вступних екзаменів: українська мова], " +//7
                        "  [Бал за результатами вступних екзаменів: математика], " +//8
                        "  [Бал за результатами вступних екзаменів: рисунок], " +//9
                        "  [Бал за результатами вступних екзаменів: композиція], " +//10
                        "  [Бал за підготовчі курси], " +//11
                        "  [Сума балів (прохідний бал)], " +//12
                        "  Пільги, " +//13
                        "  Відзнака, " +//14
                        "  [Закінчили підготовчі курси], " +//15
                        "  [Документ про освіту від 7 до 12 балів], " +//16
                        "  Немісцевий, " +//17
                        "  [Потребує гуртожитку], " +//18
                        "  Стать, " +//19
                        "  [Неповна сім'я], " +//20
                        "  Освіта, " +//21
                        "  [Документ про освіту: серія], " +//22
                        "  [Документ про освіту: номер], " +//23
                        "  [Друга спеціальність], " +//24
                        "  Зараховано," +//25
                        "  [Третя спеціальність]," +//26
                        "  [Забрав документи]," +//27
                        "  [Лист від підприємства]," +//28
                        "  [Мова1]," +//29
                        "  [Мова2]," +//30
                        "  [Дата_створення]," +//31
                        "  Форма_навчання,  " +//32
                        "  [Документи]) " +//33
                        "values " +
                        "('" + da.tbRegNumber.Text + "','" +//1
                        "  " + da.tbPib.Text + "'," +//2
                        "  " + da.lcbSpec1.keyValue + "," +//2+
                        "  " + da.tbSchoolAlgebra.Text + "," +//3
                        "  " + da.tbSchoolGeometry.Text + "," +//4
                        "  " + da.tbSchoolUkr.Text + "," +//5
                        "  '" + da.tbSchoolSr.Text + "'," +//6
                        "  " + da.tbEkzUkr.Text + "," +//7
                        "  " + da.tbEkzMat.Text + "," +//8
                        "  " + da.tbEkzPict.Text + "," +//9
                        "  " + da.tEkzKomp.Text + "," +//10
                        "  '" + da.tbEkzKyrs.Text + "'," +//11
                        "  '" + da.tbSumPB.Text + "', " +//12
                        "  '" + da.pilgi + "'," +//13
                        "  '" + da.tbVidznaka.Text + "'," +//14
                        "  " + da.kurs + "," +//15
                        "  '" + da.isEducatonDocGood + "'," +//16
                        "  '" + da.mesniy + "'," +//17
                        "  '" + da.needHostel + "'," +//18
                        "  " + da.isSexMale + "," +//19
                        "  '" + da.isFamalyNotFull + "'," +//20
                        "  " + da.education + "," +//21
                        "  '" + da.tbEdDocSerion.Text + "'," +//22
                        "  '" + da.tbEdDocNumber.Text + "'," +//23
                        "  " + da.lcbSpec2.keyValue + "," +//24
                        "  '" + da.z + "'," +//25
                        "  " + da.lcbSpec3.keyValue + "," +//26
                        "  " + da.cbNoDocs.Checked + ", " +//27
                        "  '" + da.tbLeter.Text + "', " +//28
                        "  " + da.lcbLenguage1.keyValue + ", " +//29
                        "  " + da.lcbLenguage2.keyValue + ", " +//30
                        "  '" + DateTime.Now.ToString("dd.MM.yyyy") + "', " +//31 
                        " ' " + da.cbEdForm.Text + "', " +//32
                        " ' " + da.cbDocType.Text + "' " +//33
                        "  )";

                    OleDbCommand dc = new OleDbCommand(sql, connection);

                    connection.Open();
                    OleDbDataReader or = dc.ExecuteReader();
                    connection.Close();
                    UpdateDataWithFilters();
                    //  this.головнаTableAdapter.Fill(this.probaDataSet.Головна);
                }

            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            groupBox9.Visible = true;
            this.Height = 725;
        }

        private void поискToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label16.Visible = true;
            label17.Visible = true;
            textBox16.Visible = true;
            comboBox1.Visible = true;
            groupBox4.Visible = true;
            groupBox3.Visible = true;




            this.Height = 595;
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.головнаBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.probaDataSet = new proect.probaDataSet();
            this.головнаTableAdapter = new proect.probaDataSetTableAdapters.ГоловнаTableAdapter();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.label17 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.radioButton6 = new System.Windows.Forms.RadioButton();
            this.radioButton5 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.button7 = new System.Windows.Forms.Button();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.button14 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.поискToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.справочникиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.льготыToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.языкиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.специальностиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button8 = new System.Windows.Forms.Button();
            this.cbShowDeleted = new System.Windows.Forms.CheckBox();
            this.dtpCreatedFrom = new System.Windows.Forms.DateTimePicker();
            this.dtpCreatedTill = new System.Windows.Forms.DateTimePicker();
            this.label20 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.comboBox8 = new System.Windows.Forms.ComboBox();
            this.label26 = new System.Windows.Forms.Label();
            this.button12 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button5 = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.radioButton9 = new System.Windows.Forms.RadioButton();
            this.radioButton10 = new System.Windows.Forms.RadioButton();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.radioButton7 = new System.Windows.Forms.RadioButton();
            this.radioButton8 = new System.Windows.Forms.RadioButton();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.radioButton11 = new System.Windows.Forms.RadioButton();
            this.radioButton12 = new System.Windows.Forms.RadioButton();
            this.label19 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.comboBox5 = new System.Windows.Forms.ComboBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.comboBox4 = new System.Windows.Forms.ComboBox();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.label24 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.textBox20 = new System.Windows.Forms.TextBox();
            this.textBox19 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.comboBox6 = new System.Windows.Forms.ComboBox();
            this.textBox15 = new System.Windows.Forms.ComboBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.головнаBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.probaDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.groupBox8.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.SuspendLayout();
            // 
            // головнаBindingSource
            // 
            this.головнаBindingSource.DataMember = "Головна";
            this.головнаBindingSource.DataSource = this.probaDataSet;
            // 
            // probaDataSet
            // 
            this.probaDataSet.DataSetName = "probaDataSet";
            this.probaDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // головнаTableAdapter
            // 
            this.головнаTableAdapter.ClearBeforeFill = true;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridView1.Location = new System.Drawing.Point(0, 24);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(1008, 512);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick_1);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(12, 608);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(140, 23);
            this.button3.TabIndex = 4;
            this.button3.Text = "Видалити запис";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox16
            // 
            this.textBox16.Location = new System.Drawing.Point(237, 657);
            this.textBox16.Name = "textBox16";
            this.textBox16.Size = new System.Drawing.Size(162, 20);
            this.textBox16.TabIndex = 18;
            this.textBox16.Visible = false;
            this.textBox16.TextChanged += new System.EventHandler(this.textBox16_TextChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label16.Location = new System.Drawing.Point(234, 636);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(191, 13);
            this.label16.TabIndex = 19;
            this.label16.Text = "Пошук по прізвищу абітурієнта";
            this.label16.Visible = false;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "(не обрано)",
            "МЕУ",
            "ТОМ",
            "РПЗ",
            "ЕМА",
            "КВЕТ",
            "ДЗ",
            "ЕП 9",
            "ЕП 11"});
            this.comboBox1.Location = new System.Drawing.Point(15, 655);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(162, 21);
            this.comboBox1.TabIndex = 60;
            this.comboBox1.TabStop = false;
            this.comboBox1.Visible = false;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label17.Location = new System.Drawing.Point(12, 636);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(198, 13);
            this.label17.TabIndex = 61;
            this.label17.Text = "Пошук по обраній спеціальності";
            this.label17.Visible = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.radioButton3);
            this.groupBox3.Controls.Add(this.radioButton2);
            this.groupBox3.Controls.Add(this.radioButton1);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox3.Location = new System.Drawing.Point(453, 553);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(129, 96);
            this.groupBox3.TabIndex = 62;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Пошук по статі абітурієнта";
            this.groupBox3.Visible = false;
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton3.Location = new System.Drawing.Point(6, 71);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(87, 17);
            this.radioButton3.TabIndex = 2;
            this.radioButton3.TabStop = true;
            this.radioButton3.Text = "Не обрано";
            this.radioButton3.UseVisualStyleBackColor = true;
            this.radioButton3.CheckedChanged += new System.EventHandler(this.radioButton3_CheckedChanged);
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton2.Location = new System.Drawing.Point(6, 53);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(61, 17);
            this.radioButton2.TabIndex = 1;
            this.radioButton2.TabStop = true;
            this.radioButton2.Text = "Жіноча";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton1.Location = new System.Drawing.Point(6, 35);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(70, 17);
            this.radioButton1.TabIndex = 0;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "Чоловіча";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.radioButton6);
            this.groupBox4.Controls.Add(this.radioButton5);
            this.groupBox4.Controls.Add(this.radioButton4);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox4.Location = new System.Drawing.Point(588, 553);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(188, 96);
            this.groupBox4.TabIndex = 63;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Пошук по абітурієнтам, які потербують гуртожиток";
            this.groupBox4.Visible = false;
            this.groupBox4.Enter += new System.EventHandler(this.groupBox4_Enter);
            // 
            // radioButton6
            // 
            this.radioButton6.AutoSize = true;
            this.radioButton6.Location = new System.Drawing.Point(6, 71);
            this.radioButton6.Name = "radioButton6";
            this.radioButton6.Size = new System.Drawing.Size(87, 17);
            this.radioButton6.TabIndex = 2;
            this.radioButton6.TabStop = true;
            this.radioButton6.Text = "Не обрано";
            this.radioButton6.UseVisualStyleBackColor = true;
            this.radioButton6.CheckedChanged += new System.EventHandler(this.radioButton6_CheckedChanged);
            // 
            // radioButton5
            // 
            this.radioButton5.AutoSize = true;
            this.radioButton5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton5.Location = new System.Drawing.Point(6, 53);
            this.radioButton5.Name = "radioButton5";
            this.radioButton5.Size = new System.Drawing.Size(162, 17);
            this.radioButton5.TabIndex = 1;
            this.radioButton5.TabStop = true;
            this.radioButton5.Text = "Не потребують гуртожиток";
            this.radioButton5.UseVisualStyleBackColor = true;
            this.radioButton5.CheckedChanged += new System.EventHandler(this.radioButton5_CheckedChanged);
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButton4.Location = new System.Drawing.Point(6, 32);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(147, 17);
            this.radioButton4.TabIndex = 0;
            this.radioButton4.TabStop = true;
            this.radioButton4.Text = "Потребують гуртожиток";
            this.radioButton4.UseVisualStyleBackColor = true;
            this.radioButton4.CheckedChanged += new System.EventHandler(this.radioButton4_CheckedChanged);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(7, 19);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(185, 21);
            this.button7.TabIndex = 86;
            this.button7.Text = "Вивести звіт по спеціальності";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.button14);
            this.groupBox9.Controls.Add(this.button13);
            this.groupBox9.Controls.Add(this.button9);
            this.groupBox9.Controls.Add(this.button7);
            this.groupBox9.Location = new System.Drawing.Point(782, 550);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(214, 125);
            this.groupBox9.TabIndex = 91;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "Звіти";
            this.groupBox9.Visible = false;
            // 
            // button14
            // 
            this.button14.Location = new System.Drawing.Point(7, 99);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(185, 21);
            this.button14.TabIndex = 94;
            this.button14.Text = "Мови";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Click += new System.EventHandler(this.button14_Click);
            // 
            // button13
            // 
            this.button13.Location = new System.Drawing.Point(7, 72);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(185, 21);
            this.button13.TabIndex = 93;
            this.button13.Text = "За пільговими категоріями";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // button9
            // 
            this.button9.Location = new System.Drawing.Point(7, 45);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(185, 21);
            this.button9.TabIndex = 91;
            this.button9.Text = "Вивести звіт для гуртожитку";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.поискToolStripMenuItem,
            this.toolStripMenuItem1,
            this.toolStripMenuItem2,
            this.справочникиToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1008, 24);
            this.menuStrip1.TabIndex = 92;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.menuStrip1_ItemClicked);
            // 
            // поискToolStripMenuItem
            // 
            this.поискToolStripMenuItem.Name = "поискToolStripMenuItem";
            this.поискToolStripMenuItem.Size = new System.Drawing.Size(58, 20);
            this.поискToolStripMenuItem.Text = "Пошук";
            this.поискToolStripMenuItem.Visible = false;
            this.поискToolStripMenuItem.Click += new System.EventHandler(this.поискToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(58, 20);
            this.toolStripMenuItem1.Text = "Додати";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(47, 20);
            this.toolStripMenuItem2.Text = "Звіти";
            this.toolStripMenuItem2.Visible = false;
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // справочникиToolStripMenuItem
            // 
            this.справочникиToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.льготыToolStripMenuItem,
            this.языкиToolStripMenuItem,
            this.специальностиToolStripMenuItem});
            this.справочникиToolStripMenuItem.Name = "справочникиToolStripMenuItem";
            this.справочникиToolStripMenuItem.Size = new System.Drawing.Size(94, 20);
            this.справочникиToolStripMenuItem.Text = "Справочники";
            // 
            // льготыToolStripMenuItem
            // 
            this.льготыToolStripMenuItem.Name = "льготыToolStripMenuItem";
            this.льготыToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.льготыToolStripMenuItem.Text = "Льготы";
            this.льготыToolStripMenuItem.Click += new System.EventHandler(this.льготыToolStripMenuItem_Click);
            // 
            // языкиToolStripMenuItem
            // 
            this.языкиToolStripMenuItem.Name = "языкиToolStripMenuItem";
            this.языкиToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.языкиToolStripMenuItem.Text = "Языки";
            this.языкиToolStripMenuItem.Click += new System.EventHandler(this.языкиToolStripMenuItem_Click);
            // 
            // специальностиToolStripMenuItem
            // 
            this.специальностиToolStripMenuItem.Name = "специальностиToolStripMenuItem";
            this.специальностиToolStripMenuItem.Size = new System.Drawing.Size(160, 22);
            this.специальностиToolStripMenuItem.Text = "Специальности";
            this.специальностиToolStripMenuItem.Click += new System.EventHandler(this.специальностиToolStripMenuItem_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(12, 577);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(140, 23);
            this.button8.TabIndex = 94;
            this.button8.Text = "Переглянути/змінити";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // cbShowDeleted
            // 
            this.cbShowDeleted.AutoSize = true;
            this.cbShowDeleted.BackColor = System.Drawing.SystemColors.Control;
            this.cbShowDeleted.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cbShowDeleted.Location = new System.Drawing.Point(168, 608);
            this.cbShowDeleted.Name = "cbShowDeleted";
            this.cbShowDeleted.Size = new System.Drawing.Size(269, 17);
            this.cbShowDeleted.TabIndex = 96;
            this.cbShowDeleted.Text = "Показувати абітурієнтів, що забрали документи";
            this.cbShowDeleted.UseVisualStyleBackColor = false;
            this.cbShowDeleted.CheckedChanged += new System.EventHandler(this.cbShowDeleted_CheckedChanged);
            // 
            // dtpCreatedFrom
            // 
            this.dtpCreatedFrom.Location = new System.Drawing.Point(292, 550);
            this.dtpCreatedFrom.Name = "dtpCreatedFrom";
            this.dtpCreatedFrom.Size = new System.Drawing.Size(145, 20);
            this.dtpCreatedFrom.TabIndex = 97;
            this.dtpCreatedFrom.Value = new System.DateTime(2013, 6, 20, 18, 26, 0, 0);
            this.dtpCreatedFrom.ValueChanged += new System.EventHandler(this.dtpCreatedFrom_ValueChanged);
            // 
            // dtpCreatedTill
            // 
            this.dtpCreatedTill.Location = new System.Drawing.Point(292, 579);
            this.dtpCreatedTill.Name = "dtpCreatedTill";
            this.dtpCreatedTill.Size = new System.Drawing.Size(145, 20);
            this.dtpCreatedTill.TabIndex = 98;
            this.dtpCreatedTill.Value = new System.DateTime(2014, 6, 20, 18, 26, 0, 0);
            this.dtpCreatedTill.ValueChanged += new System.EventHandler(this.dtpCreatedTill_ValueChanged);
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label20.Location = new System.Drawing.Point(165, 549);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(114, 13);
            this.label20.TabIndex = 99;
            this.label20.Text = "Дата створення з";
            this.label20.Click += new System.EventHandler(this.label20_Click);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label22.Location = new System.Drawing.Point(165, 580);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(121, 13);
            this.label22.TabIndex = 100;
            this.label22.Text = "Дата створення до";
            this.label22.Click += new System.EventHandler(this.label22_Click);
            // 
            // comboBox8
            // 
            this.comboBox8.FormattingEnabled = true;
            this.comboBox8.Location = new System.Drawing.Point(406, 611);
            this.comboBox8.Name = "comboBox8";
            this.comboBox8.Size = new System.Drawing.Size(185, 21);
            this.comboBox8.TabIndex = 103;
            this.comboBox8.SelectedIndexChanged += new System.EventHandler(this.comboBox8_SelectedIndexChanged);
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(369, 611);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(34, 13);
            this.label26.TabIndex = 102;
            this.label26.Text = "Mова";
            // 
            // button12
            // 
            this.button12.Location = new System.Drawing.Point(12, 547);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(139, 23);
            this.button12.TabIndex = 104;
            this.button12.Text = "Оновити";
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBox4);
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Location = new System.Drawing.Point(118, 81);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(223, 160);
            this.groupBox1.TabIndex = 94;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Шкільні оцінки";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(40, 102);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(73, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Середній бал";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(18, 77);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(95, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Українська мова";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(55, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 5;
            this.label2.Text = "Геометрія";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(64, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Алгебра";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(119, 99);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(30, 20);
            this.textBox4.TabIndex = 3;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(119, 74);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(30, 20);
            this.textBox3.TabIndex = 2;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(119, 48);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(30, 20);
            this.textBox2.TabIndex = 1;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(119, 22);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(30, 20);
            this.textBox1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button5);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.textBox13);
            this.groupBox2.Controls.Add(this.textBox12);
            this.groupBox2.Controls.Add(this.textBox11);
            this.groupBox2.Controls.Add(this.textBox10);
            this.groupBox2.Controls.Add(this.textBox9);
            this.groupBox2.Controls.Add(this.textBox8);
            this.groupBox2.Controls.Add(this.groupBox6);
            this.groupBox2.Location = new System.Drawing.Point(389, 81);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(196, 284);
            this.groupBox2.TabIndex = 95;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Бали за результатами вступних екзаменів";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(17, 167);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(143, 23);
            this.button5.TabIndex = 12;
            this.button5.Text = "Перерахувати суму балів";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label13.Location = new System.Drawing.Point(14, 200);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(92, 13);
            this.label13.TabIndex = 11;
            this.label13.Text = "Прохідний бал";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(37, 132);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(90, 13);
            this.label12.TabIndex = 10;
            this.label12.Text = "Підготовчі курси";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(61, 106);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(66, 13);
            this.label11.TabIndex = 9;
            this.label11.Text = "Композиція";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(78, 81);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(49, 13);
            this.label10.TabIndex = 8;
            this.label10.Text = "Рисунок";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(57, 55);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(70, 13);
            this.label9.TabIndex = 7;
            this.label9.Text = "Математика";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(32, 29);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(95, 13);
            this.label8.TabIndex = 6;
            this.label8.Text = "Українська мова";
            // 
            // textBox13
            // 
            this.textBox13.Location = new System.Drawing.Point(112, 197);
            this.textBox13.Name = "textBox13";
            this.textBox13.Size = new System.Drawing.Size(67, 20);
            this.textBox13.TabIndex = 5;
            // 
            // textBox12
            // 
            this.textBox12.Location = new System.Drawing.Point(133, 129);
            this.textBox12.Name = "textBox12";
            this.textBox12.Size = new System.Drawing.Size(46, 20);
            this.textBox12.TabIndex = 4;
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(133, 103);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(46, 20);
            this.textBox11.TabIndex = 3;
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(133, 77);
            this.textBox10.Name = "textBox10";
            this.textBox10.Size = new System.Drawing.Size(46, 20);
            this.textBox10.TabIndex = 2;
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(133, 52);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(46, 20);
            this.textBox9.TabIndex = 1;
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(133, 26);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(46, 20);
            this.textBox8.TabIndex = 0;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.radioButton9);
            this.groupBox6.Controls.Add(this.radioButton10);
            this.groupBox6.Location = new System.Drawing.Point(6, 222);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(164, 57);
            this.groupBox6.TabIndex = 69;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Склад сім\'ї";
            // 
            // radioButton9
            // 
            this.radioButton9.AutoSize = true;
            this.radioButton9.Location = new System.Drawing.Point(77, 19);
            this.radioButton9.Name = "radioButton9";
            this.radioButton9.Size = new System.Drawing.Size(69, 17);
            this.radioButton9.TabIndex = 47;
            this.radioButton9.TabStop = true;
            this.radioButton9.Text = "Неповна";
            this.radioButton9.UseVisualStyleBackColor = true;
            // 
            // radioButton10
            // 
            this.radioButton10.AutoSize = true;
            this.radioButton10.Location = new System.Drawing.Point(14, 19);
            this.radioButton10.Name = "radioButton10";
            this.radioButton10.Size = new System.Drawing.Size(57, 17);
            this.radioButton10.TabIndex = 46;
            this.radioButton10.TabStop = true;
            this.radioButton10.Text = "Повна";
            this.radioButton10.UseVisualStyleBackColor = true;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(188, 250);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(40, 13);
            this.label14.TabIndex = 96;
            this.label14.Text = "Пільги";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(176, 277);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(52, 13);
            this.label15.TabIndex = 97;
            this.label15.Text = "Відзнака";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.radioButton7);
            this.groupBox5.Controls.Add(this.radioButton8);
            this.groupBox5.Controls.Add(this.checkBox1);
            this.groupBox5.Location = new System.Drawing.Point(208, 304);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(175, 57);
            this.groupBox5.TabIndex = 98;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Місце проживання:";
            // 
            // radioButton7
            // 
            this.radioButton7.AutoSize = true;
            this.radioButton7.Location = new System.Drawing.Point(95, 16);
            this.radioButton7.Name = "radioButton7";
            this.radioButton7.Size = new System.Drawing.Size(85, 17);
            this.radioButton7.TabIndex = 39;
            this.radioButton7.TabStop = true;
            this.radioButton7.Text = "Немісцевий";
            this.radioButton7.UseVisualStyleBackColor = true;
            // 
            // radioButton8
            // 
            this.radioButton8.AutoSize = true;
            this.radioButton8.Location = new System.Drawing.Point(17, 15);
            this.radioButton8.Name = "radioButton8";
            this.radioButton8.Size = new System.Drawing.Size(72, 17);
            this.radioButton8.TabIndex = 38;
            this.radioButton8.TabStop = true;
            this.radioButton8.Text = "Місцевий";
            this.radioButton8.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(17, 36);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(134, 17);
            this.checkBox1.TabIndex = 41;
            this.checkBox1.Text = "Потребує гуртожитку";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.radioButton11);
            this.groupBox7.Controls.Add(this.radioButton12);
            this.groupBox7.Location = new System.Drawing.Point(102, 302);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(100, 62);
            this.groupBox7.TabIndex = 99;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Стать:";
            // 
            // radioButton11
            // 
            this.radioButton11.AutoSize = true;
            this.radioButton11.Location = new System.Drawing.Point(16, 39);
            this.radioButton11.Name = "radioButton11";
            this.radioButton11.Size = new System.Drawing.Size(61, 17);
            this.radioButton11.TabIndex = 44;
            this.radioButton11.TabStop = true;
            this.radioButton11.Text = "Жіноча";
            this.radioButton11.UseVisualStyleBackColor = true;
            // 
            // radioButton12
            // 
            this.radioButton12.AutoSize = true;
            this.radioButton12.Location = new System.Drawing.Point(16, 18);
            this.radioButton12.Name = "radioButton12";
            this.radioButton12.Size = new System.Drawing.Size(70, 17);
            this.radioButton12.TabIndex = 43;
            this.radioButton12.TabStop = true;
            this.radioButton12.Text = "Чоловіча";
            this.radioButton12.UseVisualStyleBackColor = true;
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(104, 399);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(141, 13);
            this.label19.TabIndex = 100;
            this.label19.Text = "Закінчили підготовчі курси";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(201, 423);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(43, 13);
            this.label21.TabIndex = 101;
            this.label21.Text = "Освіта:";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(134, 450);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(110, 13);
            this.label25.TabIndex = 102;
            this.label25.Text = "Друга спеціальність";
            // 
            // comboBox5
            // 
            this.comboBox5.FormattingEnabled = true;
            this.comboBox5.Items.AddRange(new object[] {
            "-",
            "1 місяць",
            "3 місяці",
            "6 місяців"});
            this.comboBox5.Location = new System.Drawing.Point(252, 396);
            this.comboBox5.Name = "comboBox5";
            this.comboBox5.Size = new System.Drawing.Size(115, 21);
            this.comboBox5.TabIndex = 103;
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "9 кл. загальноосвітньої школи",
            "11 кл. загальноосвітньої школи",
            "ПТУ"});
            this.comboBox3.Location = new System.Drawing.Point(252, 423);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(115, 21);
            this.comboBox3.TabIndex = 104;
            // 
            // comboBox4
            // 
            this.comboBox4.FormattingEnabled = true;
            this.comboBox4.Items.AddRange(new object[] {
            "-",
            "МЕУ",
            "ТОМ",
            "РПЗ",
            "ЕМА",
            "КВЕТ",
            "ДЗ",
            "ЕП 9",
            "ЕП 11"});
            this.comboBox4.Location = new System.Drawing.Point(252, 450);
            this.comboBox4.Name = "comboBox4";
            this.comboBox4.Size = new System.Drawing.Size(115, 21);
            this.comboBox4.TabIndex = 105;
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.label24);
            this.groupBox8.Controls.Add(this.label23);
            this.groupBox8.Controls.Add(this.textBox20);
            this.groupBox8.Controls.Add(this.textBox19);
            this.groupBox8.Location = new System.Drawing.Point(401, 370);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(184, 85);
            this.groupBox8.TabIndex = 106;
            this.groupBox8.TabStop = false;
            this.groupBox8.Text = "Документ про освіту:";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(20, 40);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(39, 13);
            this.label24.TabIndex = 55;
            this.label24.Text = "номер";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(26, 18);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(33, 13);
            this.label23.TabIndex = 54;
            this.label23.Text = "серія";
            // 
            // textBox20
            // 
            this.textBox20.Location = new System.Drawing.Point(69, 41);
            this.textBox20.Name = "textBox20";
            this.textBox20.Size = new System.Drawing.Size(100, 20);
            this.textBox20.TabIndex = 53;
            // 
            // textBox19
            // 
            this.textBox19.Location = new System.Drawing.Point(69, 15);
            this.textBox19.Name = "textBox19";
            this.textBox19.Size = new System.Drawing.Size(100, 20);
            this.textBox19.TabIndex = 52;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(369, 478);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(152, 42);
            this.button2.TabIndex = 107;
            this.button2.Text = "Змінити";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(118, 367);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(216, 17);
            this.checkBox2.TabIndex = 108;
            this.checkBox2.Text = "Документ про освіту від 7 до 12 балів";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(406, 457);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(86, 17);
            this.checkBox3.TabIndex = 109;
            this.checkBox3.Text = "Зараховано";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // comboBox6
            // 
            this.comboBox6.FormattingEnabled = true;
            this.comboBox6.Items.AddRange(new object[] {
            "-",
            "сирота",
            "Чорнобилець ",
            "інвалід І-ІІ групи",
            "багатодітна родина"});
            this.comboBox6.Location = new System.Drawing.Point(234, 247);
            this.comboBox6.Name = "comboBox6";
            this.comboBox6.Size = new System.Drawing.Size(107, 21);
            this.comboBox6.TabIndex = 110;
            // 
            // textBox15
            // 
            this.textBox15.FormattingEnabled = true;
            this.textBox15.Location = new System.Drawing.Point(234, 277);
            this.textBox15.Name = "textBox15";
            this.textBox15.Size = new System.Drawing.Size(107, 21);
            this.textBox15.TabIndex = 111;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(304, 14);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(191, 20);
            this.textBox5.TabIndex = 112;
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(304, 40);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(191, 20);
            this.textBox6.TabIndex = 113;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(225, 17);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(73, 13);
            this.label5.TabIndex = 114;
            this.label5.Text = "№ реєстрації";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(273, 43);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(25, 13);
            this.label6.TabIndex = 115;
            this.label6.Text = "ПІБ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(181, 69);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(117, 13);
            this.label7.TabIndex = 116;
            this.label7.Text = "Обрана спеціальність";
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "МЕУ",
            "ТОМ",
            "РПЗ",
            "ЕМА",
            "КВЕТ",
            "ДЗ",
            "ЕП 9",
            "ЕП 11"});
            this.comboBox2.Location = new System.Drawing.Point(304, 66);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(191, 21);
            this.comboBox2.TabIndex = 117;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.comboBox2);
            this.groupBox10.Controls.Add(this.comboBox8);
            this.groupBox10.Controls.Add(this.label7);
            this.groupBox10.Controls.Add(this.label26);
            this.groupBox10.Controls.Add(this.label6);
            this.groupBox10.Controls.Add(this.label5);
            this.groupBox10.Controls.Add(this.textBox6);
            this.groupBox10.Controls.Add(this.textBox5);
            this.groupBox10.Controls.Add(this.textBox15);
            this.groupBox10.Controls.Add(this.comboBox6);
            this.groupBox10.Controls.Add(this.checkBox3);
            this.groupBox10.Controls.Add(this.checkBox2);
            this.groupBox10.Controls.Add(this.button2);
            this.groupBox10.Controls.Add(this.groupBox8);
            this.groupBox10.Controls.Add(this.comboBox4);
            this.groupBox10.Controls.Add(this.comboBox3);
            this.groupBox10.Controls.Add(this.comboBox5);
            this.groupBox10.Controls.Add(this.label25);
            this.groupBox10.Controls.Add(this.label21);
            this.groupBox10.Controls.Add(this.label19);
            this.groupBox10.Controls.Add(this.groupBox7);
            this.groupBox10.Controls.Add(this.groupBox5);
            this.groupBox10.Controls.Add(this.label15);
            this.groupBox10.Controls.Add(this.label14);
            this.groupBox10.Controls.Add(this.groupBox2);
            this.groupBox10.Controls.Add(this.groupBox1);
            this.groupBox10.Location = new System.Drawing.Point(1144, 51);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(596, 349);
            this.groupBox10.TabIndex = 95;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "groupBox10";
            this.groupBox10.Visible = false;
            // 
            // golovna
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1008, 729);
            this.Controls.Add(this.button12);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.label20);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.dtpCreatedTill);
            this.Controls.Add(this.dtpCreatedFrom);
            this.Controls.Add(this.groupBox10);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.groupBox9);
            this.Controls.Add(this.cbShowDeleted);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.textBox16);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "golovna";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Головна";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.головнаBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.probaDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            this.groupBox8.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private void льготыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var lForm = new fmSpravBenefit();
            //lForm.Initialize();
            lForm.Show();
        }

        private void языкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var lForm = new fmSpravLanguage();
            lForm.Show();
        }

        private void специальностиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var lForm = new fmSpravSpecialty();
            lForm.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {

            connection.Close();
            viknoAbitur da = new viknoAbitur();

            string sql = GetMainSql() + " and [№ п/п] = " + ind;

            OleDbCommand dc = new OleDbCommand(sql, connection);
            connection.Open();

            OleDbDataReader currentRow = dc.ExecuteReader();
            currentRow.Read();
            if (currentRow.HasRows)
            {

                da.tbRegNumber.Text = currentRow[1].ToString();
                da.tbPib.Text = currentRow[2].ToString();

                da.lcbSpec1.comboBox.Text = currentRow[3].ToString();
                da.lcbSpec2.comboBox.Text = currentRow[25].ToString();
                da.lcbSpec3.comboBox.Text = currentRow[26].ToString();

                da.lcbLenguage1.comboBox.Text = currentRow[31].ToString();
                da.lcbLenguage2.comboBox.Text = currentRow[32].ToString();

                da.tbSchoolAlgebra.Text = currentRow[4].ToString();
                da.tbSchoolGeometry.Text = currentRow[5].ToString();
                da.tbSchoolUkr.Text = currentRow[6].ToString();
                da.tbSchoolSr.Text = currentRow[7].ToString();
                da.tbEkzUkr.Text = currentRow[8].ToString();
                da.tbEkzMat.Text = currentRow[9].ToString();
                da.tbEkzPict.Text = currentRow[10].ToString();
                da.tEkzKomp.Text = currentRow[11].ToString();
                da.tbEkzKyrs.Text = currentRow[12].ToString();
                da.tbSumPB.Text = currentRow[13].ToString();
                /*da.tbSumPB.Text = (
                        Convert.ToDouble(da.tbSchoolSr.Text) +
                        Convert.ToDouble(da.tbEkzMat.Text) +
                        Convert.ToDouble(da.tbEkzUkr.Text) +
                        Convert.ToDouble(da.tbEkzPict.Text) +
                        Convert.ToDouble(da.tEkzKomp.Text) +
                        Convert.ToDouble(da.tbEkzKyrs.Text)
                    ).ToString();*/
                da.cbBenefit.Text = currentRow[14].ToString();
                da.tbVidznaka.Text = currentRow[15].ToString();
                da.cbPk.SelectedIndex = Convert.ToInt32(currentRow[27].ToString()) - 1;
                da.cbEduDockGood.Checked = Convert.ToBoolean(currentRow[17].ToString());
                da.radioButton2.Checked = Convert.ToBoolean(currentRow[18].ToString());
                da.radioButton1.Checked = Convert.ToBoolean(currentRow[18].ToString()) == false;

                da.cbNeedHostel.Checked = Convert.ToBoolean(currentRow[19].ToString());

                da.radioButton3.Checked = currentRow[20].ToString() == "Чоловіча";
                da.radioButton4.Checked = currentRow[20].ToString() != "Чоловіча";
                da.radioButton6.Checked = Convert.ToBoolean(currentRow[21].ToString());
                da.radioButton5.Checked = Convert.ToBoolean(currentRow[21].ToString()) == false;
                da.cbEducation.Text = currentRow[22].ToString();
                da.cbDocType.Text = currentRow[34].ToString();
                da.tbEdDocSerion.Text = currentRow[23].ToString();
                da.tbEdDocNumber.Text = currentRow[24].ToString();
                da.checkBox3.Checked = Convert.ToBoolean(currentRow[28].ToString());
                da.cbNoDocs.Checked = Convert.ToBoolean(currentRow[29].ToString()); //            ",  [Забрав документи] =" 
                da.tbLeter.Text = currentRow[30].ToString();//            ",  [Лист від підприємства]
                if (currentRow[33].ToString() != "")
                {
                    da.cbEdForm.Text = currentRow[33].ToString();
                }

                if (da.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    sql = "Update Головна" +
                        " Set  " +
                        "[№ реєстрації]='" + da.tbRegNumber.Text + "', " +
                        "  ПІБ = '" + da.tbPib.Text +
                        "', Спеціальність = " + da.lcbSpec1.keyValue +
                        ",  [Шкільна оцінка по алгебрі] = " + da.tbSchoolAlgebra.Text +
                        ",  [Шкільна оцінка по геометрії] = " + da.tbSchoolGeometry.Text +
                        ",  [Шкільна оцінка по українській мові] = " + da.tbSchoolUkr.Text +
                        ",  [Шкільний середній бал] = '" + da.tbSchoolSr.Text +
                        "',  [Бал за результатами вступних екзаменів: українська мова] = " + da.tbEkzUkr.Text +
                        ",  [Бал за результатами вступних екзаменів: математика] = " + da.tbEkzMat.Text +
                        ",  [Бал за результатами вступних екзаменів: рисунок] = " + da.tbEkzPict.Text +
                        ",  [Бал за результатами вступних екзаменів: композиція] = " + da.tEkzKomp.Text +
                        ",  [Бал за підготовчі курси] = '" + da.tbEkzKyrs.Text +
                        "',  [Сума балів (прохідний бал)] = '" + da.tbSumPB.Text +
                        "',  Пільги = '" + da.pilgi +
                        "',  Відзнака = '" + da.tbVidznaka.Text +
                        "',  [Закінчили підготовчі курси] = " + da.kurs +
                        ",  [Документ про освіту від 7 до 12 балів] ='" + da.isEducatonDocGood +
                        "',  Немісцевий ='" + da.mesniy +
                        "',  [Потребує гуртожитку] = '" + da.needHostel +
                        "',  Стать =" + da.isSexMale +
                        ",  [Неповна сім'я] = '" + da.isFamalyNotFull +
                        "',  Освіта = " + da.education +
                        ",  [Документ про освіту: серія] = '" + da.tbEdDocSerion.Text +
                        "',  [Документ про освіту: номер] = '" + da.tbEdDocNumber.Text +
                        "',  [Друга спеціальність] = " + da.lcbSpec2.keyValue +
                        ",  Зараховано = " + da.z +
                        ",  [Третя спеціальність] = " + da.lcbSpec3.keyValue +
                        ",  [Забрав документи] =" + da.cbNoDocs.Checked +
                        ",  [Лист від підприємства] = " + "'" + da.tbLeter.Text + "'" +
                        ",  [Мова1] = " + da.lcbLenguage1.keyValue + "" +
                        ",  [Мова2] = " + da.lcbLenguage2.keyValue + "" +
                        ",  [Форма_навчання] = '" + da.cbEdForm.Text + "'" +
                        ",  [Документи] = '" + da.cbDocType.Text + "'" +
                        " where [№ п/п] = " + ind +
                        " ";

                    dc = new OleDbCommand(sql, connection);

                    //connection.Open();
                    OleDbDataReader or = dc.ExecuteReader();

                    connection.Close();
                    UpdateDataWithFilters();
                    //  this.головнаTableAdapter.Fill(this.probaDataSet.Головна);
                }
                else
                {

                    MessageBox.Show("Помилка!Ви не вибрали рядок.");
                }
                ;
            };
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void UpdateDataWithFilters()
        {
            connection.Close();
            string sql = GetMainSql()
                + filterGender
                + filterHostel
                + filterSpets
                + "  AND (Головна.ПІБ like '%" + textBox16.Text + "%')"
                + filterLenguage
                //+ " AND Головна.[Забрав документи] = false "
               + " ORDER BY val(Головна.[Сума балів (прохідний бал)])  DESC ";
            ;


            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataSet ds = new DataSet();
            da.Fill(ds, "Головна");
            dataGridView1.DataSource = ds.Tables["Головна"].DefaultView;
            connection.Close();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string lZvit = CopyExcelDocFromTemplate();
            connection.Close();
            string sql = GetMainSql()
               // + "   AND  Головна.[Зараховано] = true "
                + "   AND  Головна.[Потребує гуртожитку] = true "
                 + " AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)"
                + " ORDER BY val(Головна.[Сума балів (прохідний бал)])  DESC ";

            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataSet ds = new DataSet();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["zvit5"];
            xlWorkSheet.Activate();


            OleDbDataReader or1 = command.ExecuteReader();

            excelcells = xlWorkSheet.get_Range("D1");

            int count = 3;




            while (or1.Read())
            {
                excelcells = xlWorkSheet.get_Range("B" + count);
                excelcells.Value2 = or1["ПІБ"].ToString();
                excelcells = xlWorkSheet.get_Range("A" + count);
                excelcells.Value2 = count - 2;

                excelcells = xlWorkSheet.get_Range("C" + count);
                excelcells.Value2 = or1["Стать"].ToString();
                excelcells = xlWorkSheet.get_Range("D" + count);
                excelcells.Value2 = or1["Обрана Спеціальність"].ToString();
                if ((Boolean)or1["Забрав документи"] == true)
                {
                    excelcells.EntireRow.Font.ColorIndex = 3;
                    excelcells.EntireRow.Font.Bold = true;
                }

                // rang++;
                count++;

            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;
            connection.Close();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            connection.Close();
            string lSql = " SELECT " +
                "   [№] as [Код], " +
                "   [Спеціальність], " +
                "   [Розшифровка], " +
                "   [Активна] " +
                " FROM " +
                "   [Спеціальність] ";
            var ld = new Dictionary<string, string>();
            int[] Arr = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int number = 0;
            OleDbCommand comm = new OleDbCommand(lSql, connection);
            connection.Open();
            OleDbDataAdapter dbda = new OleDbDataAdapter(comm);
            //DataSet dst = new DataSet();
            OleDbDataReader dr = comm.ExecuteReader();
            while (dr.Read())
            {
                if (dr["Код"].ToString() != "11")
                {
                    ld.Add(dr["Код"].ToString(), dr["Спеціальність"].ToString());
                    Arr[number] = Convert.ToInt32(dr["Код"].ToString());
                    number++;
                }
            }

            number = 0;
            string lZvit = CopyExcelDocFor9klFromTemplate();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            for (number = 0; number < ld.Count; number++)
            {
                connection.Close();
                string sql = GetMainSql()
                   + "AND Головна.Освіта = 1"
                    + "   AND (Головна.Спеціальність = " + Arr[number].ToString() + " AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)) "
                    + " ORDER BY Головна.[ПІБ]  ";

                OleDbCommand command = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataSet ds = new DataSet();


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["page" + (number + 1).ToString()];
                xlWorkSheet.Activate();
                xlWorkSheet.Name = ld[Arr[number].ToString()];


                OleDbDataReader or1 = command.ExecuteReader();

                excelcells = xlWorkSheet.get_Range("D1");

                int count = 4;

                while (or1.Read())
                {

                    excelcells = xlWorkSheet.get_Range("C" + count);
                    excelcells.Value2 = or1["ПІБ"].ToString();
                    excelcells = xlWorkSheet.get_Range("A" + count);
                    excelcells.Value2 = count - 3;

                    excelcells = xlWorkSheet.get_Range("B" + count);
                    excelcells.Value2 = or1["№ реєстрації"].ToString();

                    excelcells = xlWorkSheet.get_Range("D" + count);
                    excelcells.Value2 = or1["Шкільна оцінка по алгебрі"].ToString();
                    excelcells = xlWorkSheet.get_Range("E" + count);
                    excelcells.Value2 = or1["Шкільна оцінка по геометрії"].ToString();
                    excelcells = xlWorkSheet.get_Range("F" + count);
                    excelcells.Value2 = or1["Шкільна оцінка по українській мові"].ToString();


                    excelcells = xlWorkSheet.get_Range("G" + count);
                    excelcells.Value2 = or1["Шкільний середній бал"].ToString();

                    excelcells = xlWorkSheet.get_Range("O" + count);
                    excelcells.Value2 = or1["Документи"].ToString();


                    excelcells = xlWorkSheet.get_Range("H" + count);
                    excelcells.Value2 = or1["Бал за результатами вступних екзаменів: композиція"].ToString();
                    excelcells = xlWorkSheet.get_Range("I" + count);
                    excelcells.Value2 = or1["Бал за результатами вступних екзаменів: рисунок"].ToString();
                    excelcells = xlWorkSheet.get_Range("J" + count);
                    excelcells.Value2 = or1["Бал за результатами вступних екзаменів: українська мова"].ToString();

                    excelcells = xlWorkSheet.get_Range("K" + count);
                    excelcells.Value2 = or1["Бал за підготовчі курси"].ToString();
                    excelcells = xlWorkSheet.get_Range("L" + count);
                    excelcells.Value2 = (
                      Convert.ToDouble(or1["Шкільна оцінка по алгебрі"].ToString()) +
                      Convert.ToDouble(or1["Шкільна оцінка по геометрії"].ToString()) +
                      Convert.ToDouble(or1["Шкільна оцінка по українській мові"].ToString()) +
                      Convert.ToDouble(or1["Шкільний середній бал"].ToString()) +
                      Convert.ToDouble(or1["Бал за результатами вступних екзаменів: композиція"].ToString()) +
                      Convert.ToDouble(or1["Бал за результатами вступних екзаменів: рисунок"].ToString()) +
                      Convert.ToDouble(or1["Бал за результатами вступних екзаменів: українська мова"].ToString()) +
                      Convert.ToDouble(or1["Бал за підготовчі курси"].ToString())).ToString();

                    excelcells = xlWorkSheet.get_Range("M" + count);
                    excelcells.Value2 = or1["Пільги"].ToString();
                    excelcells = xlWorkSheet.get_Range("N" + count);
                    if ((Boolean)or1["Забрав документи"] == true)
                    {
                        excelcells.EntireRow.Font.ColorIndex = 3;
                        excelcells.EntireRow.Font.Bold = true;
                    }
                    //excelcells.Value2 = or1["Шкільна оцінка по українській мові"].ToString();
                    /*
                    "   Головна.[], " + 
                    "   Головна.[], " + 
                    "   Головна.[Бал за результатами вступних екзаменів: математика], " + 
                    "   Головна.[], " +//10 
                    "   Головна.[], " + 
                    "   Головна.[], " + 
                    "   Головна.[Сума балів (прохідний бал)], " + 
                
                     */
                    // rang++;
                    count++;

                }
                connection.Close();
            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;
        }
        private void button15_Click(object sender, EventArgs e)
        {
            connection.Close();
            string lSql = " SELECT " +
                "   [№] as [Код], " +
                "   [Спеціальність], " +
                "   [Розшифровка], " +
                "   [Активна] " +
                " FROM " +
                "   [Спеціальність] ";
            var ld = new Dictionary<string, string>();
            int[] Arr = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int number = 0;
            OleDbCommand comm = new OleDbCommand(lSql, connection);
            connection.Open();
            OleDbDataAdapter dbda = new OleDbDataAdapter(comm);
            //DataSet dst = new DataSet();
            OleDbDataReader dr = comm.ExecuteReader();
            while (dr.Read())
            {
                if (dr["Код"].ToString() != "11")
                {
                    ld.Add(dr["Код"].ToString(), dr["Спеціальність"].ToString());
                    Arr[number] = Convert.ToInt32(dr["Код"].ToString());
                    number++;
                }
            }

            number = 0;
            string lZvit = CopyExcelDocFor9klFromTemplate();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            for (number = 0; number < ld.Count; number++)
            {
                connection.Close();
                string sql = GetMainSql()
                    + "   AND (Головна.Спеціальність = " + Arr[number].ToString() + ") "
                    + " ORDER BY Головна.[ПІБ] ";

                OleDbCommand command = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataSet ds = new DataSet();


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["page" + (number + 1).ToString()];
                xlWorkSheet.Activate();
                xlWorkSheet.Name = ld[Arr[number].ToString()];


                OleDbDataReader or1 = command.ExecuteReader();

                excelcells = xlWorkSheet.get_Range("D1");

                int count = 3;




                while (or1.Read())
                {
                    excelcells = xlWorkSheet.get_Range("C" + count);
                    excelcells.Value2 = or1["ПІБ"].ToString();
                    excelcells = xlWorkSheet.get_Range("A" + count);
                    excelcells.Value2 = count - 2;

                    excelcells = xlWorkSheet.get_Range("B" + count);
                    excelcells.Value2 = or1["№ реєстрації"].ToString();

                    excelcells = xlWorkSheet.get_Range("D" + count);
                    excelcells.Value2 = or1["Шкільна оцінка по алгебрі"].ToString();
                    excelcells = xlWorkSheet.get_Range("E" + count);
                    excelcells.Value2 = or1["Шкільна оцінка по геометрії"].ToString();
                    excelcells = xlWorkSheet.get_Range("F" + count);
                    excelcells.Value2 = or1["Шкільна оцінка по українській мові"].ToString();


                    excelcells = xlWorkSheet.get_Range("G" + count);
                    excelcells.Value2 = or1["Шкільний середній бал"].ToString();

                    excelcells = xlWorkSheet.get_Range("H" + count);
                    excelcells.Value2 = or1["Документи"].ToString();


                    excelcells = xlWorkSheet.get_Range("H" + count);
                    excelcells.Value2 = or1["Бал за результатами вступних екзаменів: композиція"].ToString();
                    excelcells = xlWorkSheet.get_Range("I" + count);
                    excelcells.Value2 = or1["Бал за результатами вступних екзаменів: рисунок"].ToString();
                    excelcells = xlWorkSheet.get_Range("J" + count);
                    excelcells.Value2 = or1["Бал за результатами вступних екзаменів: українська мова"].ToString();

                    excelcells = xlWorkSheet.get_Range("K" + count);
                    excelcells.Value2 = or1["Бал за підготовчі курси"].ToString();
                    excelcells = xlWorkSheet.get_Range("L" + count);
                    excelcells.Value2 = (
                      Convert.ToDouble(or1["Шкільна оцінка по алгебрі"].ToString()) +
                      Convert.ToDouble(or1["Шкільна оцінка по геометрії"].ToString()) +
                      Convert.ToDouble(or1["Шкільна оцінка по українській мові"].ToString()) +
                      Convert.ToDouble(or1["Шкільний середній бал"].ToString()) +
                      Convert.ToDouble(or1["Бал за результатами вступних екзаменів: композиція"].ToString()) +
                      Convert.ToDouble(or1["Бал за результатами вступних екзаменів: рисунок"].ToString()) +
                      Convert.ToDouble(or1["Бал за результатами вступних екзаменів: українська мова"].ToString()) +
                      Convert.ToDouble(or1["Бал за підготовчі курси"].ToString())).ToString();

                    excelcells = xlWorkSheet.get_Range("M" + count);
                    excelcells.Value2 = or1["Пільги"].ToString();
                    excelcells = xlWorkSheet.get_Range("N" + count);
                    if ((Boolean)or1["Забрав документи"] == true)
                    {
                        excelcells.EntireRow.Font.ColorIndex = 3;
                        excelcells.EntireRow.Font.Bold = true;
                    }
                    //excelcells.Value2 = or1["Шкільна оцінка по українській мові"].ToString();
                    /*
                    "   Головна.[], " + 
                    "   Головна.[], " + 
                    "   Головна.[Бал за результатами вступних екзаменів: математика], " + 
                    "   Головна.[], " +//10 
                    "   Головна.[], " + 
                    "   Головна.[], " + 
                    "   Головна.[Сума балів (прохідний бал)], " + 
                
                     */
                    // rang++;
                    count++;

                }
                connection.Close();
            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;
        }
        private void cbShowDeleted_CheckedChanged(object sender, EventArgs e)
        {
            //UpdateDataWithFilters();
        }

        private void dtpCreatedFrom_ValueChanged(object sender, EventArgs e)
        {
            //UpdateDataWithFilters();
        }

        private void dtpCreatedTill_ValueChanged(object sender, EventArgs e)
        {
            // UpdateDataWithFilters();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            connection.Close();
            string sql = "Delete * from Головна ";


            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open(); //открытие соединения
            command.ExecuteReader(); //выполнение запроса на удаление

            connection.Close();
            UpdateDataWithFilters();
        }

        private string CopyExcelDocFromTemplate()
        {
            string lPath = AppDomain.CurrentDomain.BaseDirectory;
            string lNewFileName = "print\\zvit_" + DateTime.Now.ToString("dd.MM.yyyy_hh.mm.ss") + ".xls";
            System.IO.File.Copy(lPath + "zvit_template.xls", lPath + lNewFileName);
            return lNewFileName;
        }

        private string CopyExcelDocFor9klFromTemplate()
        {
            string lPath = AppDomain.CurrentDomain.BaseDirectory;
            string lNewFileName = "print\\zvit_" + DateTime.Now.ToString("dd.MM.yyyy_hh.mm.ss") + ".xls";
            System.IO.File.Copy(lPath + "zvit9_template.xls", lPath + lNewFileName);
            return lNewFileName;
        }
        private string CopyExcelDocFor11klFromTemplate()
        {
            string lPath = AppDomain.CurrentDomain.BaseDirectory;
            string lNewFileName = "print\\zvit_" + DateTime.Now.ToString("dd.MM.yyyy_hh.mm.ss") + ".xls";
            System.IO.File.Copy(lPath + "zvit11_template.xls", lPath + lNewFileName);
            return lNewFileName;
        }
        private string CopyExcelDocForPTUFromTemplate()
        {
            string lPath = AppDomain.CurrentDomain.BaseDirectory;
            string lNewFileName = "print\\zvit_" + DateTime.Now.ToString("dd.MM.yyyy_hh.mm.ss") + ".xls";
            System.IO.File.Copy(lPath + "zvitPTU_template.xls", lPath + lNewFileName);
            return lNewFileName;
        }
        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            filterLenguage = " ";
            if (lcbLenguage1 != null)
            {
                if (lcbLenguage1.keyValue != 6)
                {
                    filterLenguage = " AND ((Головна.[Мова1] = " + lcbLenguage1.keyValue + ") OR (  Головна.[Мова2] = " + lcbLenguage1.keyValue + ")) "
                        + " AND ((Головна.[Мова1] <> 6) OR (  Головна.[Мова2] <> 6)) ";
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            UpdateDataWithFilters();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string lZvit = CopyExcelDocFromTemplate();
            connection.Close();
            string sql = GetMainSql()
              //  + "   AND  Головна.[Зараховано] = true "
                + "   AND  not isnull(Головна.Пільги) and Головна.Пільги <> '-' "
                + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) "
                + " ORDER BY Головна.[ПІБ] ";

            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataSet ds = new DataSet();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["zvit7"];
            xlWorkSheet.Activate();


            OleDbDataReader or1 = command.ExecuteReader();

            excelcells = xlWorkSheet.get_Range("D1");

            int count = 3;




            while (or1.Read())
            {
                excelcells = xlWorkSheet.get_Range("B" + count);
                excelcells.Value2 = or1["ПІБ"].ToString();
                excelcells = xlWorkSheet.get_Range("A" + count);
                excelcells.Value2 = count - 2;


                excelcells = xlWorkSheet.get_Range("C" + count);
                excelcells.Value2 = or1["Обрана Спеціальність"].ToString();

                excelcells = xlWorkSheet.get_Range("D" + count);
                excelcells.Value2 = or1["Пільги"].ToString();
                if ((Boolean)or1["Забрав документи"] == true)
                {
                    excelcells.EntireRow.Font.ColorIndex = 3;
                    excelcells.EntireRow.Font.Bold = true;
                }
                count++;

            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;
            connection.Close();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            string lZvit = CopyExcelDocFromTemplate();
            connection.Close();
            string sql = GetMainSql() + filterGender
                + filterHostel
                + filterSpets
                + filterLenguage
                //+ " AND  Головна.[Зараховано] = true "
                + "AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#) "
                + " ORDER BY Головна.[№ п/п] ";

            OleDbCommand command = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(command);
            DataSet ds = new DataSet();

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["zvit6"];
            xlWorkSheet.Activate();


            OleDbDataReader or1 = command.ExecuteReader();

            excelcells = xlWorkSheet.get_Range("D1");

            int count = 4;

            while (or1.Read())
            {
                excelcells = xlWorkSheet.get_Range("C" + count);
                excelcells.Value2 = or1["ПІБ"].ToString();
                excelcells = xlWorkSheet.get_Range("A" + count);
                excelcells.Value2 = count - 2;

                excelcells = xlWorkSheet.get_Range("B" + count);
                excelcells.Value2 = or1["Обрана Спеціальність"].ToString();
                excelcells = xlWorkSheet.get_Range("D" + count);
                excelcells.Value2 = or1["Мова1"].ToString();
                excelcells = xlWorkSheet.get_Range("E" + count);
                excelcells.Value2 = or1["Мова2"].ToString();
                if ((Boolean)or1["Забрав документи"] == true)
                {
                    excelcells.EntireRow.Font.ColorIndex = 3;
                    excelcells.EntireRow.Font.Bold = true;
                }
                // rang++;
                count++;

            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;
            connection.Close();
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label22_Click(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button16_Click(object sender, EventArgs e)
        {
            connection.Close();
            string lSql = " SELECT " +
                "   [№] as [Код], " +
                "   [Спеціальність], " +
                "   [Розшифровка], " +
                "   [Активна] " +
                " FROM " +
                "   [Спеціальність] ";
            var ld = new Dictionary<string, string>();
            int[] Arr = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int number = 0;
            OleDbCommand comm = new OleDbCommand(lSql, connection);
            connection.Open();
            OleDbDataAdapter dbda = new OleDbDataAdapter(comm);
            //DataSet dst = new DataSet();
            OleDbDataReader dr = comm.ExecuteReader();
            while (dr.Read())
            {
                if (dr["Код"].ToString() != "11")
                {
                    ld.Add(dr["Код"].ToString(), dr["Спеціальність"].ToString());
                    Arr[number] = Convert.ToInt32(dr["Код"].ToString());
                    number++;
                }
            }

            number = 0;
            string lZvit = CopyExcelDocForPTUFromTemplate();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            for (number = 0; number < ld.Count; number++)
            {
                connection.Close();
                string sql = GetMainSql()
                   + "AND Головна.Освіта = 3"
                    + "AND (Головна.Спеціальність = " + Arr[number].ToString() + " AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)) "
                    + " ORDER BY Головна.[ПІБ]  ";

                OleDbCommand command = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataSet ds = new DataSet();


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["page" + (number + 1).ToString()];
                xlWorkSheet.Activate();
                xlWorkSheet.Name = ld[Arr[number].ToString()];


                OleDbDataReader or1 = command.ExecuteReader();

                excelcells = xlWorkSheet.get_Range("D1");

                int count = 4;

                while (or1.Read())
                {

                    excelcells = xlWorkSheet.get_Range("C" + count);
                    excelcells.Value2 = or1["ПІБ"].ToString();
                    excelcells = xlWorkSheet.get_Range("A" + count);
                    excelcells.Value2 = count - 3;

                    excelcells = xlWorkSheet.get_Range("B" + count);
                    excelcells.Value2 = or1["№ реєстрації"].ToString();

                    excelcells = xlWorkSheet.get_Range("D" + count);
                    excelcells.Value2 = or1["Шкільний середній бал"].ToString();

                    excelcells = xlWorkSheet.get_Range("G" + count);
                    excelcells.Value2 = or1["Документи"].ToString();

                    excelcells = xlWorkSheet.get_Range("E" + count);
                    excelcells.Value2 = or1["Фаховий"].ToString();
                    excelcells = xlWorkSheet.get_Range("F" + count);
                    excelcells.Value2 = (
                      Convert.ToDouble(or1["Шкільний середній бал"].ToString()) +
                      Convert.ToDouble(or1["Фаховий"].ToString()));
                    if ((Boolean)or1["Забрав документи"] == true)
                    {
                        excelcells.EntireRow.Font.ColorIndex = 3;
                        excelcells.EntireRow.Font.Bold = true;
                    }
                    count++;
                    count++;

                }
                connection.Close();
            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;

        }

        private void button15_Click_1(object sender, EventArgs e)
        {
            connection.Close();
            string lSql = " SELECT " +
                "   [№] as [Код], " +
                "   [Спеціальність], " +
                "   [Розшифровка], " +
                "   [Активна] " +
                " FROM " +
                "   [Спеціальність] ";
            var ld = new Dictionary<string, string>();
            int[] Arr = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
                            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            int number = 0;
            OleDbCommand comm = new OleDbCommand(lSql, connection);
            connection.Open();
            OleDbDataAdapter dbda = new OleDbDataAdapter(comm);
            //DataSet dst = new DataSet();
            OleDbDataReader dr = comm.ExecuteReader();
            while (dr.Read())
            {
                if (dr["Код"].ToString() != "11")
                {
                    ld.Add(dr["Код"].ToString(), dr["Спеціальність"].ToString());
                    Arr[number] = Convert.ToInt32(dr["Код"].ToString());
                    number++;
                }
            }

            number = 0;
            string lZvit = CopyExcelDocFor11klFromTemplate();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range excelcells;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + lZvit,
                      Type.Missing, Type.Missing, Type.Missing,
                      "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                      Type.Missing, Type.Missing);
            for (number = 0; number < ld.Count; number++)
            {
                connection.Close();
                string sql = GetMainSql()
                   + "AND Головна.Освіта = 2"
                    + "AND (Головна.Спеціальність = " + Arr[number].ToString() + " AND ([Дата_створення] BETWEEN #" + dtpCreatedFrom.Value.Month.ToString() + '/' + dtpCreatedFrom.Value.Day.ToString() + '/' + dtpCreatedFrom.Value.Year.ToString() +
                "# AND #" + dtpCreatedTill.Value.Month.ToString() + '/' + dtpCreatedTill.Value.Day.ToString() + '/' + dtpCreatedTill.Value.Year.ToString() + "#)) "
                    + " ORDER BY Головна.[ПІБ]  ";

                OleDbCommand command = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataSet ds = new DataSet();


                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item["page" + (number + 1).ToString()];
                xlWorkSheet.Activate();
                xlWorkSheet.Name = ld[Arr[number].ToString()];


                OleDbDataReader or1 = command.ExecuteReader();

                excelcells = xlWorkSheet.get_Range("D1");

                int count = 4;

                while (or1.Read())
                {
                    excelcells = xlWorkSheet.get_Range("C" + count);
                    excelcells.Value2 = or1["ПІБ"].ToString();
                    excelcells = xlWorkSheet.get_Range("A" + count);
                    excelcells.Value2 = count - 3;

                    excelcells = xlWorkSheet.get_Range("B" + count);
                    excelcells.Value2 = or1["№ реєстрації"].ToString();


                    excelcells = xlWorkSheet.get_Range("D" + count);
                    excelcells.Value2 = or1["Шкільний середній бал"].ToString();

                    excelcells = xlWorkSheet.get_Range("E" + count);
                    excelcells.Value2 = or1["ЗНО1"].ToString();
                    excelcells = xlWorkSheet.get_Range("F" + count);
                    excelcells.Value2 = or1["ЗНО2"].ToString();
                    excelcells = xlWorkSheet.get_Range("G" + count);
                    excelcells.Value2 = (
                      Convert.ToDouble(or1["Шкільний середній бал"].ToString()) +
                      Convert.ToDouble(or1["ЗНО1"].ToString()) +
                      Convert.ToDouble(or1["ЗНО2"].ToString()));

                    excelcells = xlWorkSheet.get_Range("H" + count);
                    excelcells.Value2 = or1["Документи"].ToString();
                    if ((Boolean)or1["Забрав документи"] == true)
                    {
                        excelcells.EntireRow.Font.ColorIndex = 3;
                        excelcells.EntireRow.Font.Bold = true;
                    }
                    count++;

                }
                connection.Close();
            }
            xlApp.SaveWorkspace();
            xlApp.Visible = true;

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

    }
}

