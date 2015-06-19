using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace masterCatItemsInserter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] data = System.IO.File.ReadAllLines(textBox1.Text);
            int counter = 0;
            int totCount = data.GetLength(0);
            foreach (string d in data)
            {
                string ItemNumber = d.Split('\t')[0];
                string CatID = d.Split('\t')[1];
                //MessageBox.Show("ItemNumber: " + ItemNumber + "\r\n\r\n" + "CatID: " + CatID);
                Connectivity.SQLCalls.sqlQuery("INSERT INTO ItemMasterCategories (ItemNumber,WebPathID) VALUES('" + ItemNumber + "'," + CatID + ")");
                counter++;
                label1.Text = "Status: Working on Item #" + counter.ToString() + " / " + totCount.ToString();
                label1.Refresh();
            }

        }
    }
}
