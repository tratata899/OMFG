using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace proect
{
    public partial class fmSpravBenefit : Form
    {
        OleDbConnection connection = new OleDbConnection(proect.Properties.Settings.Default.probaConnectionString);
        int state = 0; //состояние 0 - ничего, 1 - вставка, 2 - редактирование
        int CurrentId = 0;
        public fmSpravBenefit()
        {
            InitializeComponent();
        }

        private void fmSpravBenefit_Load(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            RefreshData();
        }

        private void RefreshData()
        {
            connection.Close();

            string sql =
                " SELECT " +
                "   PRIVILEGE_RECID as [Код], " +
                "   PRIVILEGE_NAME as [Название] " +
                " FROM " +
                "   SPRAV_BENEFIT " +
                " ";
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "SPRAV_BENEFIT");
            dataGridView1.DataSource = ds.Tables["SPRAV_BENEFIT"].DefaultView;
            connection.Close();
        }

        private void SetControlsEnabled(bool Status)
        {
            button1.Enabled = Status;
            button2.Enabled = Status;
            button3.Enabled = Status;
            button4.Enabled = Status;
            dataGridView1.Enabled = Status;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            state = 0;
            pnItem.Visible = false;
            SetControlsEnabled(true);
        }
        
            
        private void button1_Click(object sender, EventArgs e)
        {
            state = 1;
            tbName.Text = "";
            pnItem.Visible = true;

            SetControlsEnabled(false);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            state = 2;
            pnItem.Visible = true;
            SetControlsEnabled(false);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sql =
                    " DELETE * " +
                    " FROM  SPRAV_BENEFIT " +
                    " WHERE PRIVILEGE_RECID = " + CurrentId.ToString() +
                    "  ";
            OleDbCommand dc = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataReader or = dc.ExecuteReader();
            connection.Close();

            RefreshData();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (state == 1)
            {
                string sql =
                    " INSERT INTO SPRAV_BENEFIT " +
                    " ( " +
                    "   PRIVILEGE_NAME " +
                    " ) " +
                    " VALUES ( " +
                    "   '" + tbName.Text + "' " +
                    " ) ";
                OleDbCommand dc = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataReader or = dc.ExecuteReader();
                connection.Close();

                RefreshData();
            }
            if (state == 2)
            {
                string sql =
                    " UPDATE " + 
                    "   SPRAV_BENEFIT " +
                    " SET " +
                    "   PRIVILEGE_NAME = '" + tbName.Text + "' " +
                    " WHERE " +
                    "   PRIVILEGE_RECID = " + CurrentId.ToString() +
                    "  ";
                OleDbCommand dc = new OleDbCommand(sql, connection);
                connection.Open();
                OleDbDataReader or = dc.ExecuteReader();
                connection.Close();

                RefreshData();
            }

            pnItem.Visible = false;
            SetControlsEnabled(true);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            CurrentId = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString()); 
            tbName.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
        }

        private void InitializeComponent()
        {
            this.button3 = new System.Windows.Forms.Button();
            this.pnItem = new System.Windows.Forms.Panel();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.tbName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.pnItem.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(291, 99);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(77, 23);
            this.button3.TabIndex = 15;
            this.button3.Text = "Удалить";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // pnItem
            // 
            this.pnItem.Controls.Add(this.button6);
            this.pnItem.Controls.Add(this.button5);
            this.pnItem.Controls.Add(this.tbName);
            this.pnItem.Controls.Add(this.label1);
            this.pnItem.Location = new System.Drawing.Point(12, 211);
            this.pnItem.Name = "pnItem";
            this.pnItem.Size = new System.Drawing.Size(361, 59);
            this.pnItem.TabIndex = 17;
            this.pnItem.Visible = false;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(257, 33);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(100, 23);
            this.button6.TabIndex = 3;
            this.button6.Text = "Отмена";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(151, 33);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(100, 23);
            this.button5.TabIndex = 2;
            this.button5.Text = "Ок";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // tbName
            // 
            this.tbName.Location = new System.Drawing.Point(75, 8);
            this.tbName.Name = "tbName";
            this.tbName.Size = new System.Drawing.Size(283, 20);
            this.tbName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Название";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(291, 12);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(77, 23);
            this.button4.TabIndex = 16;
            this.button4.Text = "Обновить";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(291, 70);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(77, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "Изменить";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(291, 41);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(77, 23);
            this.button1.TabIndex = 13;
            this.button1.Text = "Добавить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(273, 193);
            this.dataGridView1.TabIndex = 12;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // fmSpravBenefit
            // 
            this.ClientSize = new System.Drawing.Size(380, 281);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.pnItem);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "fmSpravBenefit";
            this.pnItem.ResumeLayout(false);
            this.pnItem.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

    }
}
