using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SearsAttributeMapper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ClearListView();
        }
        private void ClearListView()
        {
            listView1.Items.Clear();           
            listView1.Columns.Clear();
            listView1.Clear();

            ColumnHeader AttributeID = new ColumnHeader();
            ColumnHeader AttributeName = new ColumnHeader();
            AttributeID.Text = "Attribute ID";
            AttributeName.Text = "Attribute Name";
            AttributeID.Width = listView1.Width / 4;
            AttributeName.Width = (int)(listView1.Width / 1.5f);
            listView1.Columns.Add(AttributeID);
            listView1.Columns.Add(AttributeName);
            listView1.Refresh();

        }
        private void PopulateLVI(List<string> Attribs)
        {
            foreach (string att in Attribs)
            {
                ListViewItem newItm = new ListViewItem();
                newItm.Text = att.Split('|')[0];
                newItm.SubItems.Add(att.Split('|')[1]);
                listView1.Items.Add(newItm);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> tmpList = new List<string>();
            int totalLines = textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).GetLength(0);
            int counter = 0;
            foreach (string str in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                counter++;
                statusLbl.Text = "Status: Working on line " + counter.ToString() + " of " + totalLines.ToString();

                if (str.Contains('_') == true)
                {
                    string newAdd = str.Split('_')[0] + "|" + str.Split('_')[1];
                    bool existsFlag = false;
                    foreach (string existingItm in tmpList)
                    {
                        if (existingItm == newAdd)
                        {
                            existsFlag = true;
                            break;
                        }
                    }
                    if (existsFlag == false)
                    {
                        tmpList.Add(newAdd);
                    }
                }
            }
            PopulateLVI(tmpList);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string CSVFile = "";
            foreach (ListViewItem lvi in listView1.Items)
            {
                CSVFile += lvi.Text + "," + lvi.SubItems[1].Text + ",\r\n";
            }
            System.IO.File.WriteAllText(@"C:\users\tomwestbrook\desktop\searsAttributes.csv", CSVFile);
            MessageBox.Show("OK, check your desktop!");
        }
    }
}
