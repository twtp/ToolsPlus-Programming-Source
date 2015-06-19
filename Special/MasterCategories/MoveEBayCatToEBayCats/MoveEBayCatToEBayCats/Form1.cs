using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MoveEBayCatToEBayCats
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> eBayData = Connectivity.SQLCalls.sqlQuery("SELECT CategoryID,CategoryName FROM EBayCategory");
            if (eBayData.Count > 0)
            {
                int counter = 0;
                foreach (string dbLine in eBayData)
                {
                    counter++;
                    Connectivity.SQLCalls.sqlQuery("INSERT INTO EBayCategories (ID,Name) VALUES(" + dbLine.Split('|')[0] + ",'" + dbLine.Split('|')[1] + "')");
                    label1.Text = "Status: Inserting Line #" + counter.ToString() + " of " + eBayData.Count.ToString();
                    label1.Refresh();
                    this.Refresh();
                }
            }
            else { MessageBox.Show("Damnit, what happened now?"); return; }
        }
    }
}
