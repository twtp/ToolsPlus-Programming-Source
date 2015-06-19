using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace KirksEmptyLineRemover
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string WhereClause = "WHERE ";
            if (checkBox1.Checked == true)
            {
                foreach (string pLine in textBox1.Text.Split(','))
                {
                    if (pLine != "")
                    {
                        if (WhereClause.Length > 6)
                        {
                            WhereClause += " OR";
                        }
                        WhereClause += " InventoryMaster.ProductLine='" + pLine.ToUpper() + "'";
                    }
                }
            }
            else
            {
                WhereClause = "";
            }
            if (checkBox2.Checked == true)
            {
                string sqlQuery = "SELECT InventoryMaster.ItemNumber,InventoryMaster.Components FROM InventoryMaster " + WhereClause;

                List<string> results = Connectivity.SQLCalls.sqlQuery(sqlQuery);
                if (results.Count > 0)
                {
                    foreach (string result in results)
                    {
                        string itemNumber = result.Split('|')[0];
                        string newComponentText = result.Split('|')[1];
                        if (newComponentText.Contains("\r\n") == true)
                        {
                            newComponentText = newComponentText.Replace("\r\n\r\n", "\r\n").Trim();
                            Connectivity.SQLCalls.sqlQuery("UPDATE InventoryMaster SET Components='" + newComponentText + "' WHERE InventoryMaster.ItemNumber='" + itemNumber + "'");
                            label1.Text = "Status: Updated " + itemNumber;
                            label1.Refresh();
                        }
                        else
                        {
                            label1.Text = "Status: Skipped " + itemNumber + " (no edit needed)";
                            label1.Refresh();
                        }
                    }

                }
            }
            if (checkBox3.Checked == true)
            {
                string sqlQuery = "SELECT InventoryMaster.ItemNumber,InventoryMaster.EPN FROM InventoryMaster " + WhereClause;

                List<string> results = Connectivity.SQLCalls.sqlQuery(sqlQuery);
                if (results.Count > 0)
                {
                    foreach (string result in results)
                    {
                        string itemNumber = result.Split('|')[0];
                        string newEPNText = result.Split('|')[1];
                        if (newEPNText.Contains("\r\n") == true)
                        {
                            newEPNText = newEPNText.Replace("\r\n\r\n", "\r\n").Trim();
                            Connectivity.SQLCalls.sqlQuery("UPDATE InventoryMaster SET EPN='" + newEPNText + "' WHERE InventoryMaster.ItemNumber='" + itemNumber + "'");
                            label1.Text = "Status: Updated " + itemNumber;
                            label1.Refresh();
                        }
                        else
                        {
                            label1.Text = "Status: Skipped " + itemNumber + " (no edit needed)";
                            label1.Refresh();
                        }
                    }

                }
            }

                //string WhereClause = "WHERE InventoryMaster.ProductLine='"

            
        }
    }
}
