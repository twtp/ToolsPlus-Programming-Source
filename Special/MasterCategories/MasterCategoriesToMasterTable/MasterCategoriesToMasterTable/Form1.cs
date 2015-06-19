using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MasterCategoriesToMasterTable
{
    public partial class Form1 : Form
    {
        public struct MasterCategory
        {
            public string MasterName;
            public int MasterID;           
        }

       
        public List<string> MasterCategories = new List<string>();
       

        private List<MasterCategory> PopulateMasterCategories()
        {
            List<MasterCategory> MasterCategoriesList = new List<MasterCategory>();

            List<string> MainCategories = Connectivity.SQLCalls.sqlQuery("SELECT ID, WebPathName FROM WebPaths WHERE PathLevel=1");
            if (MainCategories.Count > 0)
            {
                foreach (string webPath in MainCategories)
                {
                    List<string> Level1Categories = Connectivity.SQLCalls.sqlQuery("SELECT ID, WebPathName FROM WebPaths WHERE PathLevel=2 AND ParentID=" + webPath.Split('|')[0]);
                    if (Level1Categories.Count > 0)
                    {
                        foreach (string webPath2 in Level1Categories)
                        {
                            List<string> Level2Categories = Connectivity.SQLCalls.sqlQuery("SELECT ID,WebPathName FROM WebPaths WHERE PathLevel=3 AND ParentID=" + webPath2.Split('|')[0]);
                            if (Level2Categories.Count > 0)
                            {
                                foreach (string webPath3 in Level2Categories)
                                {
                                    List<string> Level3Categories = Connectivity.SQLCalls.sqlQuery("SELECT ID,WebPathName FROM WebPaths WHERE PathLevel=4 AND ParentID=" + webPath3.Split('|')[0]);
                                    if (Level3Categories.Count > 0)
                                    {
                                        foreach (string webPath4 in Level3Categories)
                                        {
                                            List<string> Level4Categories = Connectivity.SQLCalls.sqlQuery("SELECT ID,WebPathName FROM WebPaths WHERE PathLevel=5 AND ParentID=" + webPath4.Split('|')[0]);
                                            if (Level4Categories.Count > 0)
                                            {
                                                foreach (string webPath5 in Level4Categories)
                                                {
                                                    MasterCategory MC = new MasterCategory();
                                                    MC.MasterID = int.Parse(webPath5.Split('|')[0]);
                                                    MC.MasterName = webPath.Split('|')[1] + " > " + webPath2.Split('|')[1] + " > " + webPath3.Split('|')[1] + " > " + webPath4.Split('|')[1] + " > " + webPath5.Split('|')[1];
                                                    MasterCategoriesList.Add(MC);
                                                }
                                            }
                                            else
                                            {
                                                //no 5th subcat... finalize the last...
                                                MasterCategory MC = new MasterCategory();
                                                MC.MasterID = int.Parse(webPath4.Split('|')[0]);
                                                MC.MasterName = webPath.Split('|')[1] + " > " + webPath2.Split('|')[1] + " > " + webPath3.Split('|')[1] + " > " + webPath4.Split('|')[1];
                                                MasterCategoriesList.Add(MC);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        MasterCategory MC = new MasterCategory();
                                        MC.MasterID = int.Parse(webPath3.Split('|')[0]);
                                        MC.MasterName = webPath.Split('|')[1] + " > " + webPath2.Split('|')[1] + " > " + webPath3.Split('|')[1];
                                        MasterCategoriesList.Add(MC);

                                    }
                                }
                            }
                            else
                            {
                                MasterCategory MC = new MasterCategory();
                                MC.MasterID = int.Parse(webPath2.Split('|')[0]);
                                MC.MasterName = webPath.Split('|')[1] + " > " + webPath2.Split('|')[1];
                                MasterCategoriesList.Add(MC);
                            }
                        }
                    }
                    else
                    {
                        MasterCategory MC = new MasterCategory();
                        MC.MasterID = int.Parse(webPath.Split('|')[0]);
                        MC.MasterName = webPath.Split('|')[1];
                        MasterCategoriesList.Add(MC);
                    }
                }
            }
            else
            {
                MessageBox.Show("Uhm...wtf???");
                return new List<MasterCategory>();
            }


            return MasterCategoriesList;
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void ResetLVI()
        {
            mcLVI.Items.Clear();
            mcLVI.Columns.Clear();
            mcLVI.Clear();
            mcLVI.Refresh();

            ColumnHeader mcId = new ColumnHeader();
            ColumnHeader mcName = new ColumnHeader();
            mcId.Text = "ID";
            mcName.Text = "Name";
            mcId.Width = mcLVI.Width / 10;
            mcName.Width = mcLVI.Width - (mcLVI.Width / 10);
            mcLVI.Columns.Add(mcId);
            mcLVI.Columns.Add(mcName);
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            ResetLVI();
            List<MasterCategory> MCList = PopulateMasterCategories();

            foreach (MasterCategory MC in MCList)
            {
                ListViewItem tmpLVI = new ListViewItem();
                tmpLVI.Text = MC.MasterID.ToString();
                tmpLVI.SubItems.Add(MC.MasterName);
                mcLVI.Items.Add(tmpLVI);
            }
            totLbl.Text = "Total: " + mcLVI.Items.Count.ToString();
            totLbl.Refresh();

            if (verifyChk.Checked == true)
            {
                int counter = 0;
                foreach (MasterCategory MC in MCList)
                {
                    counter++;
                    statusLbl.Text = "Status: Verifying #" + counter.ToString() + " of " + MCList.Count.ToString();
                    statusLbl.Refresh();
                    List<string> verifyNoChild = Connectivity.SQLCalls.sqlQuery("SELECT ID,WebPathName FROM WebPaths WHERE ParentID=" + MC.MasterID);
                    if (verifyNoChild.Count > 0)
                    {
                        DialogResult res = MessageBox.Show("Web Path " + MC.MasterName + " (ID #" + MC.MasterID.ToString() + ") has a child, meaning it shouldn't be in the master list!\r\nOne of it's child's ID is #" + verifyNoChild[0].Split('|')[0],"Warning",MessageBoxButtons.YesNo);
                        if (res == System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
                        {
                            ListViewItem lvi = FindListViewItemForMessage(MC.MasterID.ToString());
                            if (lvi != null)
                            {
                                this.mcLVI.Items.Remove(lvi);
                            }
                            mcLVI.Refresh();
                            foreach (string cat in verifyNoChild)
                            {
                                MasterCategory newCat = new MasterCategory();
                                newCat.MasterID = int.Parse(cat.Split('|')[0]);
                                newCat.MasterName = cat.Split('|')[1];                               

                                ListViewItem tmpItm = new ListViewItem();
                                tmpItm.Text = newCat.MasterID.ToString();
                                tmpItm.SubItems.Add(MC.MasterName + " > " + newCat.MasterName);
                                mcLVI.Items.Add(tmpItm);
                                totLbl.Text = "Total: " + mcLVI.Items.Count.ToString();
                                totLbl.Refresh();
                            }
                        }
                    }
                }
            }

        }
        private ListViewItem FindListViewItemForMessage(string s)
        {
            foreach (ListViewItem lvi in this.mcLVI.Items)
            {
                if (StringComparer.OrdinalIgnoreCase.Compare(lvi.Text, s) == 0)
                {
                    return lvi;
                }
            }

            return null;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ResetLVI();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int counter =0;
            foreach (ListViewItem itm in mcLVI.Items)
            {
                counter++;
                Connectivity.SQLCalls.sqlQuery("INSERT INTO MasterCategories (ID,Name) VALUES(" + itm.Text + ",'" + itm.SubItems[1].Text + "')");
                statusLbl.Text = "Status: Inserting " + counter.ToString() + " of " + mcLVI.Items.Count.ToString();
                statusLbl.Refresh();
            }
        }

        
    }
}
