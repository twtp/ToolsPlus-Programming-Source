using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (string str in textBox1.Text.Split(new string[]{"\r\n"},StringSplitOptions.RemoveEmptyEntries))
            {
                List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ComponentID,QuantityShipped,TimeInserted,SalesOrderNo FROM WhseTaskSalesOrderLines INNER JOIN WhseTaskSalesOrderHeader ON WhseTaskSalesOrderLines.HeaderID = WhseTaskSalesOrderHeader.ID WHERE ComponentID IN (SELECT ComponentID FROM InventoryComponentMap WHERE ItemNumber='" + str + "')");
                foreach (string res in results)
                {
                    textBox2.Text += res + "\r\n";
                }
               

            }
        }
    }
}
