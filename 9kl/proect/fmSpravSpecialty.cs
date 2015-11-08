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
    public partial class fmSpravSpecialty : Form
    {

        OleDbConnection connection = new OleDbConnection(proect.Properties.Settings.Default.probaConnectionString);
        int state = 0; //состояние 0 - ничего, 1 - вставка, 2 - редактирование
        int CurrentId = 0;

        public fmSpravSpecialty()
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
                "   [№] as [Код], " +
                "   [Спеціальність], " +
                "   [Розшифровка], " +
                "   [Активна] " +
                " FROM " +
                "   [Спеціальність] " +
                " ";
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, "Спеціальність");
            dataGridView1.DataSource = ds.Tables["Спеціальність"].DefaultView;
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
            tbShortName.Text = "";
            tbName.Text = "";
            cbActive.Checked = true;
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
                    " FROM  [Спеціальність] " +
                    " WHERE [№] = " + CurrentId.ToString() +
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
                    " INSERT INTO [Спеціальність] " +
                    " ( " +
                    "   [Спеціальність], " +
                    "   [Розшифровка], " +
                    "   [Активна] " +
                    " ) " +
                    " VALUES ( " +
                    "   '" + tbShortName.Text + "', " +
                    "   '" + tbName.Text + "', " +
                    "   " + cbActive.Checked.ToString() + " " +
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
                    "   [Спеціальність] " +
                    " SET " +
                    "   [Спеціальність] = '" + tbShortName.Text + "', " +
                    "   [Розшифровка] = '" + tbName.Text + "', " +
                    "   [Активна] = " + cbActive.Checked.ToString() + " " +
                    " WHERE " +
                    "   [№] = " + CurrentId.ToString() +
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
            tbShortName.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString(); ;
            tbName.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString(); ;
            cbActive.Checked = Convert.ToBoolean(dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString()); 
        }
    }
}
