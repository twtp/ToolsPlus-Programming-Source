using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RestockingStuffer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int qty = 0;
            if (int.TryParse(textBox2.Text,out qty) == true)
            {
                if (qty > 0)
                {
                    List<string> result = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM InventoryMaster WHERE ItemNumber='" + textBox1.Text + "'");
                    if (result.Count > 0)
                    {
                        List<string> componentID = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID FROM InventoryComponentMap WHERE ItemNumber='" + textBox1.Text + "'");
                        if (componentID.Count > 0 )
                        {
                            int counter = 0;
                            foreach (string cid in componentID)
                            {
                                string comID = cid.Split('|')[0];
                                string query = "INSERT INTO LocationContents (LocationID,ComponentID,MasterID,Quantity,LastInventoriedDate) VALUES(10830," + comID + ",10830," + qty + ",'"+DateTime.Now.Year.ToString("0000") + "-" + DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + " " + DateTime.Now.TimeOfDay.Hours.ToString("00") + ":" + DateTime.Now.TimeOfDay.Minutes.ToString("00") + ":" + DateTime.Now.TimeOfDay.Seconds.ToString("00") + ".000')";
                                Connectivity.SQLCalls.sqlQuery(query);
                                //textBox3.Text += query + "\r\n\r\n";
                                counter++;
                                MessageBox.Show("Finished adding " + counter.ToString() + " lines to Restocking Staging");
                            }

                            //should double check here....

                        }
                        else
                        {
                            MessageBox.Show("No data returned. Failed.");
                        }                        
                    }
                    else
                    {
                        MessageBox.Show("That item doesn't already exist in our system and needs to in order to use this tool.");
                    }
                }
                else
                {
                    MessageBox.Show("Qty must be greater than 0");
                }
            }
            else
            {
                MessageBox.Show("Not a numerical value inputted into qty.");
            }
        }
    }
}
