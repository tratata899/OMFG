using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace proect
{
    public class LookupComboBox
    {
        OleDbConnection connection = new OleDbConnection(proect.Properties.Settings.Default.probaConnectionString);
        public string sqlText { get; set; }
        public string tableName { get; set; }

        public ComboBox comboBox { get { return fComboBox; } }
        public int keyValue { get { return data[fComboBox.Text]; } set { fComboBox.Text = data.Where(i => i.Value == value).First().Key; } }

        private Dictionary<string, int> data = new Dictionary<string, int>(); 
        private ComboBox fComboBox;

        public void Refresh()
        {
            connection.Close();

            string sql = sqlText;
            OleDbCommand myCommand = new OleDbCommand(sql, connection);
            connection.Open();
            OleDbDataAdapter da = new OleDbDataAdapter(myCommand);
            DataSet ds = new DataSet();
            da.Fill(ds, tableName);
            //dataGridView1.DataSource = ds.Tables["SPRAV_BENEFIT"].DefaultView;

            comboBox.Items.Clear();
            data.Clear();
            for (int i = 0; i < ds.Tables[tableName].Rows.Count; i++ )
            {
                string str;
                var row = ds.Tables[tableName].Rows[i];
                var itemId = Convert.ToInt32(row[0].ToString());
                var itemText = row[1].ToString();

                comboBox.Items.Add(itemText);
                data.Add(itemText, itemId);
            }
            connection.Close();
            comboBox.SelectedItem = comboBox.Items[0];
        }

        public LookupComboBox(string aSqlText, string aTableName, ComboBox aCombobox)
        {
            sqlText = aSqlText;
            tableName = aTableName;
            fComboBox = aCombobox;

            Refresh();
        }

        private LookupComboBox()
        { 
        }
    }
}
