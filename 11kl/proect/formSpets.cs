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
    public partial class formSpets : Form
    {
        public formSpets()
        {
            InitializeComponent();
        }

        public void Initialize()
        {
            var gConnection = new OleDbConnection(aConnectedString);
            dataSetSpets1.Load(gConnection , null, null);
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
