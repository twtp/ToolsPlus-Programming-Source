using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ClearATote
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM LocationMaster WHERE Division1=" + textBox1.Text + " AND Division2 IS NULL");

            if (results.Count == 1)
            {
                Connectivity.SQLCalls.sqlQuery("DELETE FROM LocationContents WHERE LocationID=" + results[0].Split('|')[0]);
                MessageBox.Show("Complete. Removed Items From Location ID: " + results[0].Split('|')[0]);
            }
            else
            {
                MessageBox.Show("Either nothing is in the tote or the location doesn't exist!");
            }

        }
    }
}
