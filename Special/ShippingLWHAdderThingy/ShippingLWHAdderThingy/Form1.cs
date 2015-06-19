using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ShippingLWHAdderThingy
{
    public partial class Form1 : Form
    {

        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

           
        }

        private void SetStarShipConnectionString()
        {
            string connStr = "Data Source=10.0.50.17\\bkupexec;Initial Catalog=StarShip;Integrated Security=SSPI;";
            Connectivity.SQLCalls.connectionString = connStr;
        }

        private string LoadCSVFile(string filepath)
        {
            return System.IO.File.ReadAllText(filepath);
        }

        private void CrapToFile()
        {
            string newFile = "";

            foreach(ColumnHeader c in listView1.Columns)
            {
                newFile += c.Text + ",";
            }
            newFile += "\r\n";

            foreach (ListViewItem lvi in listView1.Items)
            {
                for (int x =0;x < lvi.SubItems.Count;x++)
                {
                    newFile += lvi.SubItems[x].Text + ",";

                }
                newFile += "\r\n";
                    
            }

            System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\shipping_results.csv", newFile);
            MessageBox.Show("Holy Shit, we finished!!!!");
        }


        private void PopulateLVI()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;


            string file = LoadCSVFile("c:\\users\\tomwestbrook\\desktop\\shipping.csv");
            int colCount = file.Split('\r')[0].Split(',').GetLength(0);

            string columns = file.Split('\r')[0];

            List<string> Columns = new List<string>();
            foreach (string col in columns.Split(','))
            {
                Columns.Add(col);
            }

            foreach (string col in Columns)
            {
                ColumnHeader colCol = new ColumnHeader();
                colCol.Text = col;
                listView1.Columns.Add(colCol);
            }
            listView1.Refresh();
            //return;
            int count = 0;
            foreach (string row in file.Split(new string[] { "\r\n"},StringSplitOptions.None))
            {
                if (count > 0)
                {
                    ListViewItem lvi = new ListViewItem();
                    int count2 = 0;
                    foreach (string s in row.Split(','))
                    {
                        if (count2 > 0)
                        {
                            lvi.SubItems.Add(s);
                        }
                        else
                        {
                            lvi.Text = s;
                            
                        }
                        count2++;
                    }
                    listView1.Items.Add(lvi);
                }
                count = count + 1;
            }


            

            //List<string> Lines = new List<string>();
            int csvcount = 0;
            foreach (ListViewItem lvi in listView1.Items)
            {
                string TrackingNumber = lvi.Text;
                Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT Length,Width,Height FROM Pack WHERE TrackingNumber='" + TrackingNumber + "'");
                if (qr.Rows.Count > 0)
                {
                    try
                    {
                        lvi.SubItems[24].Text = qr.Rows[0].Split('|')[0];
                        lvi.SubItems[25].Text = qr.Rows[0].Split('|')[1];
                        lvi.SubItems[26].Text = qr.Rows[0].Split('|')[2];
                    }
                    catch { }
                }
                else if (qr.Rows.Count > 1)
                {
                    try
                    {
                        lvi.SubItems[24].Text = "Multiple Entries Found.";
                        lvi.SubItems[25].Text = "Multiple Entries Found.";
                        lvi.SubItems[26].Text = "Multiple Entries Found.";
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        lvi.SubItems[24].Text = "No Data Found.";
                        lvi.SubItems[25].Text = "No Data Found.";
                        lvi.SubItems[26].Text = "No Data Found.";
                    }
                    catch { }
                }
                label1.Text = csvcount.ToString() + "/" + listView1.Items.Count.ToString();
                label1.Refresh();
                csvcount++;
            }

              


        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetStarShipConnectionString();
            PopulateLVI();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CrapToFile();
        }
    }
}
