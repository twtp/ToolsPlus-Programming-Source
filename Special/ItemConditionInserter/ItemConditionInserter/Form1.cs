using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ItemConditionInserter
{
    public partial class Form1 : Form
    {
        Dictionary<string, bool> ItemCondition = new Dictionary<string, bool>();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

            Step1();



            Step2();


        }

        private void Step1()
        {
            button1.Enabled = false;
            ItemCondition = new Dictionary<string, bool>();
            SetStatus("Status: Gathering all items and their original condition... Please wait a minute.");
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber,IsReconditioned FROM InventoryMaster");
            if (qr.Rows.Count > 0)
            {
                foreach (string row in qr.Rows)
                {
                    string ItemNumber = row.Split('|')[0];
                    if (ItemNumber == "AST3688R")
                    {
                       
                    }
                    bool Condition = (row.Split('|')[1] == "True" ? true : false);
                    ItemCondition.Add(ItemNumber, Condition);
                }
            }
            else
            {
                MessageBox.Show("Something happened when trying to pull data from InventoryMaster....damnit!");
                return;
            }
        }
        private void Step2()
        {
            SetStatus("Preparing to update item conditions...");
            int count = 0;
            foreach (KeyValuePair<string,bool> kvp in ItemCondition)
            {
                SetStatus("Updating Item (" + count.ToString() + "/" + ItemCondition.Count.ToString() + ")...");
                string sqlCommand = "INSERT INTO ItemConditionLines (ItemNumber,ConditionID) VALUES('" + kvp.Key + "'," + (kvp.Value == true ? "2" : "1") + ")";
                Connectivity.SQLCalls.sqlQuery(sqlCommand);
                count++;
            }
        }


        private void SetStatus(string statusMessage)
        {
            statusLbl.Text = statusMessage;
            statusLbl.Refresh();
        }
    }
}
