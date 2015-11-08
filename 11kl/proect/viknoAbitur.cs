using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace proect
{
    public partial class viknoAbitur : Form
    {
        public LookupComboBox lcbSpec1, lcbSpec2, lcbSpec3;
        public LookupComboBox lcbBenefit;
        public LookupComboBox lcbLenguage1, lcbLenguage2;
        OleDbConnection connection = new OleDbConnection(proect.Properties.Settings.Default.probaConnectionString);
        public int mesniy = 0;
        public int spec1;
        public int spec2;
        public int spec3;
        public double suma = 0;

        public viknoAbitur()
        {
            InitializeComponent();

            lcbSpec1 = new LookupComboBox(
                "select [№], [Спеціальність] FROM [Спеціальність]",
                "Спеціальність", 
                cbSpec1
              );
            lcbSpec2 = new LookupComboBox(
                "select [№], [Спеціальність] FROM [Спеціальність]",
                "Спеціальність",
                cbSpec2
              );
            lcbSpec3 = new LookupComboBox(
                "select [№], [Спеціальність] FROM [Спеціальність]",
                "Спеціальність",
                cbSpec3
              );
            lcbBenefit = new LookupComboBox(
                "select PRIVILEGE_RECID, PRIVILEGE_NAME FROM SPRAV_BENEFIT",
                "SPRAV_BENEFIT",
                cbBenefit
              );
            lcbLenguage1 = new LookupComboBox(
                "select LANGUAGE_RECID, LANGUAGE_NAME FROM SPRAV_LANGUAGE",
                "SPRAV_LANGUAGE",
                cbLeng1
              );
            lcbLenguage2 = new LookupComboBox(
                "select LANGUAGE_RECID, LANGUAGE_NAME FROM SPRAV_LANGUAGE",
                "SPRAV_LANGUAGE",
                cbLeng2
                );
           /* lcbSpec2.comboBox.SelectedIndex = 8;
            lcbSpec3.comboBox.SelectedIndex = 8;
            lcbLenguage1.keyValue = 6;
            lcbLenguage2.keyValue = 6;*/
                        
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
          //  string sql = "Insert into Головна ( [№ реєстрації], ПІБ, Спеціальність, [Шкільна оцінка по алгебрі], [Шкільна оцінка по геометрії], [Шкільна оцінка по українській мові], [Шкільний середній бал], [Бал за результатами вступних екзаменів: українська мова], [Бал за результатами вступних екзаменів: математика], [Бал за результатами вступних екзаменів: рисунок], [Бал за результатами вступних екзаменів: композиція], [Бал за підготовчі курси], [Сума балів (прохідний бал)], Пільги, Відзнака, [Закінчили підготовчі курси], [Документ про освіту від 7 до 12 балів], Немісцевий, [Потребує гуртожитку], Стать, [Неповна сім'я], Освіта, [Документ про освіту: серія], [Документ про освіту: номер], [Друга спеціальність] ) values (" + textBox1.Text + ",'" + textBox2.Text + "'," + s1 + "," + textBox4.Text + "," + textBox5.Text + "," + textBox6.Text + "," + textBox7.Text + "," + textBox8.Text + "," + textBox9.Text + "," + textBox10.Text + "," + textBox11.Text + "," + textBox12.Text + "," + textBox13.Text + ", '" + textBox14.Text + "','" + textBox15.Text + "'," + k + ",'" + textBox17.Text + "','" + m + "','" + g + "'," + i + ",'" + n + "'," + o + ",'" + textBox19.Text + "','" + textBox20.Text + "'," + s2 + ")";
          //  MessageBox.Show(sql);
          //OleDbCommand dc = new OleDbCommand(sql, connection);

          //connection.Open();
          //  OleDbDataReader or = dc.ExecuteReader();

          //  connection.Close();
          //  this.Close();
            
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                cbNeedHostel.Enabled = true;
                mesniy = 1;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true) {
                cbNeedHostel.Enabled = false;
                mesniy = 0;
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSpec1.SelectedItem == cbSpec2.SelectedItem)
            {
                MessageBox.Show("ПОМИЛКА! Обрана та друга спеціальності не можуть бути однаковими");

            }
            else
            {
                if ((cbSpec2.SelectedItem == "МЕУ") || (cbSpec2.SelectedItem == "ТОМ") || (cbSpec2.SelectedItem == "РПЗ") || (cbSpec2.SelectedItem == "ЕМА")
                    || (cbSpec2.SelectedItem == "КВЕТ") || (cbSpec2.SelectedItem == "ЕП 9") || (cbSpec2.SelectedItem == "ЕП 11"))
                {
                    label10.Enabled = false;
                    label11.Enabled = false;
                    tbEkzPict.Enabled = false;
                    tEkzKomp.Enabled = false;
                    label9.Enabled = true;
                    tbEkzMat.Enabled = true;

                }

                if (cbSpec2.SelectedItem == "ДЗ")
                {
                    label9.Enabled = false;
                    tbEkzMat.Enabled = false;
                    label10.Enabled = true;
                    label11.Enabled = true;
                    tbEkzPict.Enabled = true;
                    tEkzKomp.Enabled = true;

                }

                if (lcbSpec1 != null)
                {
                    spec1 = lcbSpec1.keyValue;
                }
                /*
                if (comboBox1.SelectedItem == "МЕУ")
                {
                    s1 = 1;
                }
                if (comboBox1.SelectedItem == "ТОМ")
                {
                    s1 = 2;
                }
                if (comboBox1.SelectedItem == "РПЗ")
                {
                    s1 = 3;
                }
                if (comboBox1.SelectedItem == "ЕМА")
                {
                    s1 = 4;
                }
                if (comboBox1.SelectedItem == "КВЕТ")
                {
                    s1 = 5;
                }
                if (comboBox1.SelectedItem == "ДЗ")
                {
                    s1 = 6;
                }
                if (comboBox1.SelectedItem == "ЕП 9")
                {
                    s1 = 7;
                }
                if (comboBox1.SelectedItem == "ЕП 11")
                {
                    s1 = 8;
                }*/
            }
        }
        public int isSexMale = 0;
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked == true) {
                isSexMale = 1;
            } 
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked == true)
            {
                isSexMale = 2;
            } 
        }
        public int education = 1;
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbEducation.SelectedItem == "9 кл. загальноосвітньої школи")
            {

                education = 1;

            }
            if (cbEducation.SelectedItem == "11 кл. загальноосвітньої школи")
            {

                education = 2;

            }
            if (cbEducation.SelectedItem == "ПТУ")
            {

                education = 3;

            }
            
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {

        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbSpec1.SelectedItem == cbSpec2.SelectedItem)
            {
                MessageBox.Show("ПОМИЛКА! Обрана та друга спеціальності не можуть бути однаковими");

            }
            else
            {
                if (lcbSpec2 != null)
                {                    
                    spec2 = lcbSpec2.keyValue;
               }
                /*if (comboBox4.SelectedItem == "МЕУ")
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
                }*/
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*if ((comboBox1.SelectedItem == "МЕУ") || (comboBox1.SelectedItem == "ТОМ") || (comboBox1.SelectedItem == "РПЗ") || (comboBox1.SelectedItem == "ЕМА")
               || (comboBox1.SelectedItem == "КВЕТ") || (comboBox1.SelectedItem == "ЕП 9") || (comboBox1.SelectedItem == "ЕП 11"))
            {
           
                suma = Convert.ToDouble(textBox7.Text) + (Convert.ToDouble(textBox8.Text) + Convert.ToDouble(textBox9.Text)) / 2 + Convert.ToDouble(textBox12.Text);
                textBox13.Text = Convert.ToString(suma);
            }

            if (comboBox1.SelectedItem == "ДЗ")
            {
          
                suma = Convert.ToDouble(textBox7.Text) + (Convert.ToDouble(textBox8.Text) + Convert.ToDouble(textBox10.Text) + Convert.ToDouble(textBox11.Text)) / 3;
                textBox13.Text = Convert.ToString(suma);

            }*/

            tbSumPB.Text = (
                Convert.ToDouble(tbSchoolSr.Text) +
                Convert.ToDouble(tbEkzMat.Text)+
                Convert.ToDouble(tbEkzUkr.Text) +
                Convert.ToDouble(tbEkzPict.Text) +
                Convert.ToDouble(tEkzKomp.Text) +
                Convert.ToDouble(tbEkzKyrs.Text)
                /*+Convert.ToDouble(tbEkzUkr.Text) +
                Convert.ToDouble(tbEkzMat.Text) + м
                Convert.ToDouble(tbEkzPict.Text) +
                Convert.ToDouble(tEkzKomp.Text) +
                Convert.ToDouble(tbEkzKyrs.Text)*/
                ).ToString();
        }
        public int kurs = 0;
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbPk.SelectedItem == "-")
            {

                kurs = 1;

            }
            if (cbPk.SelectedItem == "1 місяць")
            {

                kurs = 2;

            }
            if (cbPk.SelectedItem == "3 місяці")
            {

                kurs = 3;

            }
            if (cbPk.SelectedItem == "6 місяців")
            {

                kurs = 4;

            }
        }
        public int isFamalyNotFull = 0;
        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton6.Checked == true)
            {
                isFamalyNotFull = 1;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton5.Checked == true)
            {
                isFamalyNotFull = 0;
            }
        }
        public int needHostel = 0;
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNeedHostel.Checked == true)
            {
                needHostel = 1;
            }
            if (cbNeedHostel.Checked == false)
            {
                needHostel = 0;
            }
        }

              
       
        public int isEducatonDocGood = 0;
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (cbEduDockGood.Checked == true)
            {
                isEducatonDocGood = 1;
            }
            if (cbEduDockGood.Checked == false)
            {
                isEducatonDocGood = 0;
            }
        }
        public int z = 0;
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
        public string pilgi = "";
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            pilgi =  Convert.ToString(cbBenefit.SelectedItem);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ((cbSpec3.SelectedItem == cbSpec2.SelectedItem) || (cbSpec1.SelectedItem == cbSpec3.SelectedItem))
            {
                MessageBox.Show("ПОМИЛКА! Обрана та друга спеціальності не можуть бути однаковими");
            }
            else
            {
                if (lcbSpec3 != null)
                {
                    spec3 = lcbSpec3.keyValue;
                }
            }
        }

        private void viknoAbitur_Shown(object sender, EventArgs e)
        {
            button2_Click(sender, e);
        }

       
        private void tbLeter_TextChanged(object sender, EventArgs e)
        {

        }

        private void label27_Click(object sender, EventArgs e)
        {
            
        }

        private void tbSchoolUkr_TextChanged(object sender, EventArgs e)
        {

        }

        private void tbSchoolSr_TextChanged(object sender, EventArgs e)
        {

        }

        private void tbSchoolGeometry_TextChanged(object sender, EventArgs e)
        {

        }

        private void tbSchoolAlgebra_TextChanged(object sender, EventArgs e)
        {

        }

        private void groupBox5_Enter_1(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void viknoAbitur_Load(object sender, EventArgs e)
        {

        }

        private void tbEkzKyrs_TextChanged(object sender, EventArgs e)
        {

        }

      
       

       
    }
}
