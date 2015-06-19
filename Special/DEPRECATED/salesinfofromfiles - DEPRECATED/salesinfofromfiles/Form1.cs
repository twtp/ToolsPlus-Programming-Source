using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace salesinfofromfiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string CustomerInfo = File.ReadAllText("c:\\users\\tomwestbrook\\desktop\\merge\\customerinfo.txt");
            textBox1.Text = "";
            long lineCount = 0;
            long totLines = CustomerInfo.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length;
            foreach (string str in CustomerInfo.Split(new string[]{"\r\n"},StringSplitOptions.RemoveEmptyEntries))
            {
                lineCount++;
                //label1.Text = "Status: " + lineCount.ToString() + "/" + totLines.ToString();
                if (str.TrimStart().StartsWith("E") | str.TrimStart().StartsWith("Y"))
                {
                    textBox1.Text += str.TrimStart().Split(' ')[0];
                    textBox1.Text += ",";
                    textBox1.Text += str.Substring(10).TrimStart().Split(' ')[0];
                    textBox1.Text += ",";
                    textBox1.Text += str.Substring(35).TrimStart().Split(' ')[0];
                    textBox1.Text += ",";
                    textBox1.Text += str.Substring(56).TrimStart().Split(new string[] {"  "},StringSplitOptions.None)[0];
                    textBox1.Text += ",";
                    textBox1.Text += str.Substring(75).TrimStart().Split(' ')[0];
                    textBox1.Text += ",\r\n";
                    //label1.Refresh();
                    //textBox1.Refresh();

                }
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string salesOrders = File.ReadAllText("c:\\users\\tomwestbrook\\desktop\\merge\\salesinfo.txt");
            string checkFile = File.ReadAllText("c:\\users\\tomwestbrook\\desktop\\merge\\customerinfo.txt");
            int total = salesOrders.Split(new string[] { "\r\n" }, StringSplitOptions.None).Length;
            int count = 0;

            foreach (string str in salesOrders.Split(new string[] { "\r\n" }, StringSplitOptions.None))
            {
                count++;
                try
                {
                    label1.Text = "Status: " + count.ToString() + "/" + total.ToString();
                    label1.Refresh();

                    string compID = str.Split('|')[0];
                    SQLCalls.SQLCalls newSql = new SQLCalls.SQLCalls();
                    List<string> results = newSql.sqlQuery("SELECT InventoryComponentMap.ItemNumber,ItemDescription FROM InventoryComponentMap,InventoryMaster WHERE ComponentID='" + compID + "' AND InventoryMaster.ItemNumber=InventoryComponentMap.ItemNumber");
                    string ItemNumber = "N/A";
                    string itemDescription = "N/A";
                    if (results.Count > 0)
                    {
                        ItemNumber = results[0].Split('|')[0];
                        itemDescription = results[0].Split('|')[1];
                    }
                    if (ItemNumber.StartsWith("MET") == true | ItemNumber.StartsWith("D-W") == true)
                    {
                        string orderNo = str.Split('|')[3];
                        if (orderNo == "")
                        {
                            orderNo = "N/A";
                        }
                        if (orderNo.StartsWith("E") | orderNo.StartsWith("Y"))
                        {
                            string customerName = "N/A";
                            string customerPhone = "N/A";
                            string customerOrderDate = "N/A";
                            if (checkFile.Contains(orderNo) == true)
                            {
                                int indexOfCustomer = checkFile.IndexOf(orderNo, 1);

                                string customer = checkFile.Substring((int)indexOfCustomer).Split(new string[] { "\r\n" }, StringSplitOptions.None)[0];
                                if (customer.Length > 105)
                                {
                                    customerName = customer.Substring(55).Split(new string[] { "  " }, StringSplitOptions.None)[0];
                                    customerPhone = customer.Substring(103).Split(new string[] { "  " }, StringSplitOptions.None)[0];
                                    customerOrderDate = customer.Substring(19).Split(' ')[0];
                                    //"0063510            4/4/2013   C            02-BRA3486  Victor M. Bravo                                 520-303-3486
                                    if (customerOrderDate == "")
                                    {
                                        customerOrderDate = "N/A";
                                    }
                                    textBox1.AppendText(customerOrderDate + "," + ItemNumber + "," + itemDescription + "," + customerName + "," + customerPhone + "," + orderNo + ",\r\n");
                                    textBox1.Refresh();
                                }
                            }
                        }
                    }
                }
                catch { }

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
