using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GoogleCategoriesLinker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
            ProcessEvents();
        }
        private void ProcessEvents()
        {
            SetupListView();
            string[] file = System.IO.File.ReadAllLines(textBox1.Text);
            long counter = 0;
            int id_offset = int.Parse(idOffsetTxt.Text);
            int category_offset = int.Parse(categoryOffsetTxt.Text);
            foreach (string line in file)
            {
                if (counter == 0)
                {
                    //skip (header)
                }
                else
                {
                    string[] columns = line.Split('\t');
                    string ItemNumber = columns[id_offset];
                    string GoogleCategory = columns[category_offset];
                    string GoogleCategoryID = "N/A";
                    List<string> MasterPath = new List<string>();
                    List<string> MasterPathID = new List<string>();
                    List<string> res = Connectivity.SQLCalls.sqlQuery("SELECT WebPathID,MasterCategories.Name FROM PartNumberWebPaths INNER JOIN MasterCategories ON PartNumberWebPaths.WebPathID=MasterCategories.ID WHERE ItemNumber='" + ItemNumber + "'");
                    if (res.Count > 0)
                    {
                        foreach (string s in res)
                        {
                            MasterPathID.Add(s.Split('|')[0]);
                            MasterPath.Add(s.Split('|')[1]);
                        }
                    }
                    res = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM GoogleCategories WHERE Name='" + GoogleCategory + "'");
                    if (res.Count > 0)
                    {
                        GoogleCategoryID = res[0].Split('|')[0];
                    }
                    
                    for (int x = 0; x < MasterPath.Count; x++)
                    {
                        ListViewItem tmpItm = new ListViewItem();
                        tmpItm.Text = counter.ToString();
                        tmpItm.SubItems.Add(ItemNumber);
                        tmpItm.SubItems.Add(GoogleCategoryID);
                        tmpItm.SubItems.Add(GoogleCategory);
                        tmpItm.SubItems.Add(MasterPathID[x]);
                        tmpItm.SubItems.Add(MasterPath[x]);                        
                        listView1.Items.Add(tmpItm);
                    }

                }
                counter++;

            }
            MessageBox.Show("Total Lines: " + listView1.Items.Count.ToString());
        }

        private void SetupListView()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.Refresh();
            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader itmCol = new ColumnHeader();
            ColumnHeader googleCatIDCol = new ColumnHeader();
            ColumnHeader googleCatCol = new ColumnHeader();
            ColumnHeader masterIDCol = new ColumnHeader();
            ColumnHeader masterCatCol = new ColumnHeader();
            idCol.Text = "ID";
            itmCol.Text = "ItemNumber";
            googleCatIDCol.Text = "Google ID";
            googleCatCol.Text = "Google Category";
            masterIDCol.Text = "Master ID";
            masterCatCol.Text = "Master Category";
            listView1.Columns.Add(idCol);
            listView1.Columns.Add(itmCol);
            listView1.Columns.Add(googleCatIDCol);
            listView1.Columns.Add(googleCatCol);
            listView1.Columns.Add(masterIDCol);
            listView1.Columns.Add(masterCatCol);
            listView1.Refresh();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SetupListView();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int count = 0;
            int added = 0;
            foreach (ListViewItem lvi in listView1.Items)
            {
                statusLbl.Text = "Status: Working on line #" + count.ToString() + " out of " + listView1.Items.Count.ToString();
                statusLbl.Refresh();
                string ConnectorType="5";
                string ConnectorID = lvi.SubItems[2].Text;
                string MasterID = lvi.SubItems[4].Text;

                try
                {
                    int.Parse(ConnectorID);
                    int.Parse(MasterID);

                    List<string> res = Connectivity.SQLCalls.sqlQuery("SELECT MasterID FROM MasterCategoryConnectors WHERE ConnectorType=5 AND MasterID=" + MasterID + " AND ConnectorID=" + ConnectorID);
                    if (res.Count == 0)
                    {
                        Connectivity.SQLCalls.sqlQuery("INSERT INTO MasterCategoryConnectors (MasterID,ConnectorType,ConnectorID) VALUES(" + MasterID + ",5," + ConnectorID + ")");
                        added++;
                    }
                }
                catch { MessageBox.Show("Trouble Connecting... ConnectorID: " + ConnectorID + "  /  MasterID: " + MasterID + "\r\nGoogle Category: " + lvi.SubItems[3].Text + "\r\nMaster Category: " + lvi.SubItems[5].Text); }
                count++;
            }
            MessageBox.Show("Complete! Added " + added.ToString() + " lines out of " + listView1.Items.Count.ToString());
        }
    }
}
