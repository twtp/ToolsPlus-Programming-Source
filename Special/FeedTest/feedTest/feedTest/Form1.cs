using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace feedTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessStart();
        }
        private void ProcessStart()
        {
            SetupLVI();
            label1.Text = "Status: Loading All DB Items. This could take a minute...";
            label1.Refresh();
            List<string> ItemsInDB = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT ItemNumber,dbo.GetMasterCategory(ItemNumber) AS MasterCategory FROM InventoryMaster");
            long counter = 0;
            foreach (string itm in ItemsInDB)
            {
                counter++;
                string ItemNumber = itm.Split('|')[0];
                string MasterCategory = itm.Split('|')[1];
                ListViewItem tmpItm = new ListViewItem();
                tmpItm.Text = ItemNumber;
                tmpItm.SubItems.Add(MasterCategory);
                if (MasterCategory.Contains('>')==true)
                {
                    tmpItm.BackColor = Color.Green;
                }
                else if (MasterCategory.Length < 5 & MasterCategory.Length > 0)
                {
                    tmpItm.BackColor = Color.Goldenrod;
                }
                else
                {
                    tmpItm.BackColor = Color.Red;
                }
                listView1.Items.Add(tmpItm);
                label1.Text = "Status: Processing Line #" + counter.ToString() + "/" + ItemsInDB.Count.ToString();
                label1.Refresh();
            }
        }
        private void SetupLVI()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.Refresh();
            ColumnHeader ItemNumber = new ColumnHeader();
            ColumnHeader MasterCategory = new ColumnHeader();
            ItemNumber.Text = "Item Number:";
            ItemNumber.Width = listView1.Width / 4;
            MasterCategory.Text = "Master Category: ";
            MasterCategory.Width = listView1.Width - (listView1.Width / 4);
            listView1.Columns.Add(ItemNumber);
            listView1.Columns.Add(MasterCategory);
            listView1.Refresh();
        }
    }
}
