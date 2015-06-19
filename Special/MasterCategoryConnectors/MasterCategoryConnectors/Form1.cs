using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MasterCategoryConnectors
{
    public partial class Form1 : Form
    {
        public Color ListViewBackColor = Color.White;
        public Color ListViewForeColor = Color.Black;
        //public Color ListViewHighlightBackColor = Color.Orange;
        //public Color ListViewHighlightForeColor = Color.Black;
        public Color ListViewHighlightBackColor = SystemColors.Highlight;
        public Color ListViewHighlightForeColor = SystemColors.HighlightText;


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            ResetMasterLVI();
            ResetChildLVI();
            ResetConnectionsLVI();
            SetToolTips();

            LoadChildrenList();
            ShowChildrenList();

            LoadMaster();
            WriteMasterToLVI();
            
        }
        private struct ChildrenConnectionInfo
        {
            public int ConnectorID;
            public string ConnectorTableName;
            public string ConnectorName;
        }
        private struct CategoryInfo
        {
            public int CategoryID;
            public string CategoryName;
        }
        private Dictionary<int, ChildrenConnectionInfo> ChildrenData = new Dictionary<int, ChildrenConnectionInfo>();
        private Dictionary<int, CategoryInfo> ChildCategoryData = new Dictionary<int, CategoryInfo>();
        private Dictionary<int, CategoryInfo> MasterCategoryData = new Dictionary<int, CategoryInfo>();
       // private Dictionary<int, string> MasterLookup = new Dictionary<int, string>();
       // private Dictionary<int, string> ChildLookup = new Dictionary<int, string>();
        int CurrentChildTableID = 0;
        bool childNeedsRefresh = false;
        bool masterNeedsRefresh = false;
        private int CurrentChildType = -1;

        private void ResetMasterLVI()
        {
            MasterLVI.Items.Clear();
            MasterLVI.Columns.Clear();
            MasterLVI.Clear();
            MasterLVI.View = View.Details;
            MasterLVI.GridLines = true;
            MasterLVI.FullRowSelect = true;
            //MasterLVI.HideSelection = false;
            MasterLVI.HideSelection = true;
            MasterLVI.BackColor = ListViewBackColor;
            MasterLVI.ForeColor = ListViewForeColor;
            MasterLVI.Refresh();


            ColumnHeader internalIDCol = new ColumnHeader();
            ColumnHeader CatIDCol = new ColumnHeader();
            ColumnHeader CatNameCol = new ColumnHeader();

            internalIDCol.Text = "";
            CatIDCol.Text = "ID";
            CatNameCol.Text = "Category Name";

            internalIDCol.Width = 1;
            CatIDCol.Width = MasterLVI.ClientSize.Width / 10;
            CatNameCol.Width = MasterLVI.ClientSize.Width - CatIDCol.Width - 25;
            
            MasterLVI.Columns.Add(internalIDCol);
            MasterLVI.Columns.Add(CatIDCol);
            MasterLVI.Columns.Add(CatNameCol);                        

        }
        private void ResetChildLVI()
        {
            ChildLVI.Items.Clear();
            ChildLVI.Columns.Clear();
            ChildLVI.Clear();
            ChildLVI.View = View.Details;
            ChildLVI.GridLines = true;
            ChildLVI.FullRowSelect = true;
            //ChildLVI.HideSelection = false;
            ChildLVI.HideSelection = true;
            ChildLVI.BackColor = ListViewBackColor;
            ChildLVI.ForeColor = ListViewForeColor;
            ChildLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader CatIDCol = new ColumnHeader();
            ColumnHeader CatNameCol = new ColumnHeader();
            idCol.Text = "";
            CatIDCol.Text = "ID";
            CatNameCol.Text = "Category Name";
            idCol.Width = 1;
            CatIDCol.Width = ChildLVI.ClientSize.Width / 10;
            CatNameCol.Width = ChildLVI.ClientSize.Width - CatIDCol.Width - 25;

            ChildLVI.Columns.Add(idCol);
            ChildLVI.Columns.Add(CatIDCol);
            ChildLVI.Columns.Add(CatNameCol);
        }
        private void ResetConnectionsLVI()
        {
            connectionsLVI.Items.Clear();
            connectionsLVI.Columns.Clear();
            connectionsLVI.Clear();
            connectionsLVI.View = View.Details;
            connectionsLVI.GridLines = true;
            connectionsLVI.FullRowSelect = true;
            connectionsLVI.HideSelection = true;
            connectionsLVI.BackColor = ListViewBackColor;
            connectionsLVI.ForeColor = ListViewForeColor;
            connectionsLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader MasterIDCol = new ColumnHeader();
            ColumnHeader ChildIDCol = new ColumnHeader();
            ColumnHeader MasterCol = new ColumnHeader();
            ColumnHeader ChildCol = new ColumnHeader();

            idCol.Width = 1;
            idCol.Text = "";
            MasterIDCol.Width = 1;
            MasterIDCol.Text = "";
            ChildIDCol.Width = 1;
            ChildIDCol.Text = "";
            MasterCol.Width = connectionsLVI.ClientSize.Width / 2 - 15;
            ChildCol.Width = connectionsLVI.ClientSize.Width / 2 - 15;
            MasterCol.Text = "Master Category";
            ChildCol.Text = "Child Category";

            connectionsLVI.Columns.Add(idCol);
            connectionsLVI.Columns.Add(MasterIDCol);
            connectionsLVI.Columns.Add(ChildIDCol);
            connectionsLVI.Columns.Add(MasterCol);
            connectionsLVI.Columns.Add(ChildCol);


        }


        private void LoadChildrenList()
        {
            ChildrenData = new Dictionary<int, ChildrenConnectionInfo>();
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ConnectorType,ConnectorTableName FROM MasterCategoryConnectorsTypes");
            if (results.Count > 0)
            {
                foreach (string line in results)
                {
                    string id = line.Split('|')[0];
                    string name = line.Split('|')[1];
                    string common_name = name.Split(new string[] { "Categories" }, StringSplitOptions.None)[0];
                    ChildrenConnectionInfo cci = new ChildrenConnectionInfo();
                    cci.ConnectorID = int.Parse(id);
                    cci.ConnectorName = common_name;
                    cci.ConnectorTableName = name;
                    if (cci.ConnectorTableName != "MasterCategories")
                    {
                        ChildrenData.Add(cci.ConnectorID, cci);
                    }
                }
            }
            else
            {
                MessageBox.Show("Failed to load childrens. Tell Tom you need help getting your computer online.");
            }
        }
        private void ShowChildrenList()
        {
            childCmb.Items.Clear();
            foreach (KeyValuePair<int, ChildrenConnectionInfo> kvp in ChildrenData)
            {
                childCmb.Items.Add(kvp.Value.ConnectorName);
            }
        }
        private void LoadSelectedChild()
        {
            int ConnectorID = 0;
            string TableName = "";
            string Name = childCmb.Items[childCmb.SelectedIndex].ToString();
           
            ChildCategoryData.Clear();
            //ChildLookup = new Dictionary<int, string>();
            ResetChildLVI();

            foreach (KeyValuePair<int,ChildrenConnectionInfo> kvp in ChildrenData)
            {
                if (kvp.Value.ConnectorName == Name)
                {
                    TableName = kvp.Value.ConnectorTableName;
                    ConnectorID = kvp.Value.ConnectorID;
                    CurrentChildTableID = ConnectorID;
                }
            }

            if (TableName.Length > 0)
            {
                int count = 0;
                List<string> categories = Connectivity.SQLCalls.sqlQuery("SELECT ID, Name FROM " + TableName);
                if (categories.Count>0)
                {
                    foreach (string line in categories)
                    {
                        string CatID = line.Split('|')[0];
                        string CatName = line.Split('|')[1];                        
                        CategoryInfo ci = new CategoryInfo();
                        ci.CategoryID = int.Parse(CatID);
                        ci.CategoryName = CatName;
                        ChildCategoryData.Add(ci.CategoryID, ci);
                        //ChildLookup.Add(ci.CategoryID,ci.CategoryName);
                        count++;
                    }

                    WriteChildToLVI();
                }
            }
            else
            {
                MessageBox.Show("Could not find chosen table");
            }


        }
        private void LoadCurrentConnectors()
        {
            ResetConnectionsLVI();
            if (childCmb.SelectedIndex != -1)
            {
                ChildrenConnectionInfo cci = new ChildrenConnectionInfo();
                cci.ConnectorTableName = childCmb.Items[childCmb.SelectedIndex] + "Categories";
                cci.ConnectorName = childCmb.Items[childCmb.SelectedIndex].ToString();
                List<string> result = Connectivity.SQLCalls.sqlQuery("SELECT ConnectorType FROM MasterCategoryConnectorsTypes WHERE ConnectorTableName='" + cci.ConnectorTableName + "'");
                if (result.Count > 0)
                {
                    cci.ConnectorID = int.Parse(result[0].Split('|')[0]);
                    CurrentChildType = cci.ConnectorID;
                }
                else
                {
                    MessageBox.Show("No internet? Que?");
                    return;
                }

                List<string> Connectors = Connectivity.SQLCalls.sqlQuery("SELECT MasterID,ConnectorID FROM MasterCategoryConnectors WHERE ConnectorType=" + cci.ConnectorID.ToString());
                int count = 0;
                connectionsLVI.ListViewItemSorter = null;
                foreach (string line in Connectors)
                {
                    int MasterID = int.Parse(line.Split('|')[0]);
                    int ChildID = int.Parse(line.Split('|')[1]);
                    string MasterName = "";
                    string ChildName = "";
                    if (MasterCategoryData.ContainsKey(MasterID) == true)
                    {
                        MasterName = MasterCategoryData[MasterID].CategoryName;
                    }
                    if (ChildCategoryData.ContainsKey(ChildID) == true)
                    {
                        ChildName = ChildCategoryData[ChildID].CategoryName;
                    }

                    if (MasterName.Length > 0 & ChildName.Length > 0)
                    {
                        ListViewItem lvi = new ListViewItem();
                        lvi.Text = count.ToString();
                        lvi.SubItems.Add(MasterID.ToString());
                        lvi.SubItems.Add(ChildID.ToString());
                        lvi.SubItems.Add(MasterName);
                        lvi.SubItems.Add(ChildName);
                        connectionsLVI.Items.Add(lvi);
                        count++;
                    }
                }
            }
            connectionsLVI.ListViewItemSorter = new ListViewItemComparer(3);
            ConnectedCountLbl.Text = "(" + connectionsLVI.Items.Count.ToString() + ")";
        }
        private void WriteChildToLVI()
        {
            ResetChildLVI();
            if (hideMasterConnectedChk.Checked == true)
            {
                //List<int> ConnectedIDs = new List<int>();
                List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ConnectorID FROM MasterCategoryConnectors WHERE ConnectorType=" + CurrentChildTableID.ToString());
                if (results.Count > 0)
                {
                    foreach (string line in results)
                    {
                        //ConnectedIDs.Add(int.Parse(line.Split('|')[0]));
                        int ID = int.Parse(line.Split('|')[0]);
                        ChildCategoryData.Remove(ID);
                    }
                }
            }
            else
            {
                if (childNeedsRefresh == true)
                {
                    childNeedsRefresh = false;
                    LoadSelectedChild();
                    ResetChildLVI();
                }
            }
            if (hideChildHighPriceChk.Checked == true)
            {
                List<int> removeKeys = new List<int>();
                foreach (KeyValuePair<int,CategoryInfo> highPriced in ChildCategoryData)
                {
                    if (highPriced.Value.CategoryName.Contains("$$$")==true)
                    {
                        //ChildCategoryData.Remove(highPriced.Key);
                        removeKeys.Add(highPriced.Key);
                    }

                }
                foreach (int i in removeKeys)
                {
                    ChildCategoryData.Remove(i);
                }
            }


            int count = 0;
            ChildLVI.ListViewItemSorter = null;
            foreach (KeyValuePair<int, CategoryInfo> kvp in ChildCategoryData)
            {
                int CatID = kvp.Value.CategoryID;//line.Split('|')[0];
                string CatName = kvp.Value.CategoryName;//line.Split('|')[1];
                ListViewItem lvi = new ListViewItem();
                lvi.Text = count.ToString();
                lvi.SubItems.Add(CatID.ToString());
                lvi.SubItems.Add(CatName);
                lvi.BackColor = ListViewBackColor;
                lvi.ForeColor = ListViewForeColor;
                ChildLVI.Items.Add(lvi);
                count++;
            }
            ChildLVI.ListViewItemSorter = new ListViewItemComparer(2);
            ChildCountLbl.Text = "(" + ChildLVI.Items.Count.ToString() + ")";
        }        
        public class ListViewItemComparer : IComparer
        {
            private int col;
            public ListViewItemComparer()
            {
                col = 0;
            }
            public ListViewItemComparer(int column)
            {
                col = column;
            }
            public int Compare(object x, object y)
            {
                int returnVal = -1;
                returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
                ((ListViewItem)y).SubItems[col].Text);
                return returnVal;
            }
        }
        private void LoadMaster()
        {
            MasterCategoryData = new Dictionary<int, CategoryInfo>();
            //MasterLookup = new Dictionary<int,string>();
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ID,Name FROM MasterCategories");
           
            if (results.Count >0)
            {
                foreach (string line in results)
                {
                    int MasterID = int.Parse(line.Split('|')[0]);
                    string MasterName = line.Split('|')[1];

                    CategoryInfo ci = new CategoryInfo();
                    ci.CategoryID = MasterID;
                    ci.CategoryName=MasterName;

                    MasterCategoryData.Add(MasterID, ci);
                    //MasterLookup.Add(MasterID,ci.CategoryName);
                }
            }
            else
            {
                MessageBox.Show("Tell Tom your internet must be down.");
            }
        }
        private void WriteMasterToLVI()
        {
            
            if (hideChildConnectedChk.Checked == true)
            {
                List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT MasterID FROM MasterCategoryConnectors WHERE ConnectorType=" + CurrentChildTableID.ToString());
                if (results.Count > 0)
                {
                    foreach (string line in results)
                    {
                        //ConnectedIDs.Add(int.Parse(line.Split('|')[0]));
                        int ID = int.Parse(line.Split('|')[0]);
                        MasterCategoryData.Remove(ID);
                    }
                }
            }

                if (masterNeedsRefresh == true)
                {
                    masterNeedsRefresh = false;
                    LoadMaster();
                    WriteMasterToLVI();
                    return;
                }
            

            int count =0;
            ResetMasterLVI();
            MasterLVI.ListViewItemSorter = null;
            foreach (KeyValuePair<int,CategoryInfo> mCats in MasterCategoryData)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = count.ToString();
                lvi.SubItems.Add(mCats.Value.CategoryID.ToString());
                lvi.SubItems.Add(mCats.Value.CategoryName);
                lvi.BackColor = ListViewBackColor;
                lvi.ForeColor = ListViewForeColor;
                MasterLVI.Items.Add(lvi);
                count++;
            }
            MasterLVI.ListViewItemSorter = new ListViewItemComparer(2);
            MasterCountLbl.Text = "(" + MasterLVI.Items.Count.ToString() + ")";
        }





        private void SetToolTips()
        {
            ToolTip tt = new ToolTip();
            tt.SetToolTip(hideChildConnectedChk,"Checking this box will filter out all master categories that have already been connected to the current child categories collection.");
            tt.SetToolTip(hideMasterConnectedChk, "Checking this box will filter out all child categories that have already been connected to the master category collection.");
            tt.SetToolTip(hideChildHighPriceChk, "Checking this box will filter out all child categories that contain '$$$' in their name. Most likely used just for EBay Categories.");
            tt.SetToolTip(connectBtn, "Pressing this button will connect selected master category to selected child category.");
            tt.SetToolTip(disconnectBtn, "Pressing this button will remove the connection that is selected.");
        }

        private void childCmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadSelectedChild();
            LoadCurrentConnectors();
            MasterLVI.Focus();
        }

        private void hideMasterConnectedChk_CheckedChanged(object sender, EventArgs e)
        {
            childNeedsRefresh = true;
            WriteChildToLVI();
        }

        private void hideChildConnectedChk_CheckedChanged(object sender, EventArgs e)
        {
            if (hideChildConnectedChk.Checked == false)
            {
                masterNeedsRefresh = true;
            }
            WriteMasterToLVI();
        }

        private List<ListViewItem> PreviouslySelectedMasterLVI = new List<ListViewItem>();
        private void MasterLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in PreviouslySelectedMasterLVI)
            {
                lvi.BackColor = ListViewBackColor;
                lvi.ForeColor = ListViewForeColor;
            }
            PreviouslySelectedMasterLVI.Clear();
            foreach (ListViewItem lvi in MasterLVI.SelectedItems)
            {
                lvi.BackColor = ListViewHighlightBackColor;
                lvi.ForeColor = ListViewHighlightForeColor;
                PreviouslySelectedMasterLVI.Add(lvi);
            }
            
        }
        private List<ListViewItem> PreviouslySelectedChildLVI = new List<ListViewItem>();
        private void ChildLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in PreviouslySelectedChildLVI)
            {
                lvi.BackColor = ListViewBackColor;
                lvi.ForeColor = ListViewForeColor;
            }
            PreviouslySelectedChildLVI.Clear();
            foreach (ListViewItem lvi in ChildLVI.SelectedItems)
            {
                lvi.BackColor = ListViewHighlightBackColor;
                lvi.ForeColor = ListViewHighlightForeColor;
                PreviouslySelectedChildLVI.Add(lvi);
            }
        }

        private void connectBtn_Click(object sender, EventArgs e)
        {
            List<int> SelectedMasters = new List<int>();
            List<int> SelectedChilds = new List<int>();
            foreach (ListViewItem selected in MasterLVI.SelectedItems)
            {
                SelectedMasters.Add(int.Parse(selected.SubItems[1].Text));
            }
            foreach (ListViewItem selected in ChildLVI.SelectedItems)
            {
                SelectedChilds.Add(int.Parse(selected.SubItems[1].Text));
            }

            if (SelectedMasters.Count > 1 & SelectedChilds.Count > 1)
            {
                MessageBox.Show("Cant select multiple masters AND multiple childs. Only one direction at a time.");
                return;
            }
            else
            {
                foreach (int m in SelectedMasters)
                {
                    foreach (int c in SelectedChilds)
                    {
                        Connectivity.SQLCalls.sqlQuery("INSERT INTO MasterCategoryConnectors (MasterID,ConnectorType,ConnectorID) VALUES(" + m.ToString() + "," + CurrentChildTableID.ToString() + "," + c.ToString() + ")");
                        int id = connectionsLVI.Items.Count;
                        ListViewItem lvi = new ListViewItem();
                        lvi.Text = id.ToString();
                        lvi.SubItems.Add(m.ToString());
                        lvi.SubItems.Add(c.ToString());
                        lvi.SubItems.Add(MasterCategoryData[m].CategoryName);                        
                        lvi.SubItems.Add(ChildCategoryData[c].CategoryName);
                        lvi.BackColor = ListViewBackColor;
                        lvi.ForeColor = ListViewForeColor;
                        connectionsLVI.Items.Add(lvi);
                        if (hideMasterConnectedChk.Checked == true)
                        {
                            MasterCategoryData.Remove(c);
                            ResetChildLVI();
                            WriteChildToLVI();
                        }
                       
                    }
                    if (hideChildConnectedChk.Checked == true)
                    {
                        MasterCategoryData.Remove(m);
                        ResetMasterLVI();
                        WriteMasterToLVI();
                    }

                }

            }
            ConnectedCountLbl.Text = "(" + connectionsLVI.Items.Count.ToString() + ")";
        }

        private void disconnectBtn_Click(object sender, EventArgs e)
        {
            List<int> selection = new List<int>();
            foreach (ListViewItem lvi in connectionsLVI.SelectedItems)
            {
                int masterID = int.Parse(lvi.SubItems[1].Text);
                int childID = int.Parse(lvi.SubItems[2].Text);
                string masterName = lvi.SubItems[3].Text;
                string childName = lvi.SubItems[4].Text;
                //int masterID = MasterLookup[masterName];
                //int childID = ChildLookup[childName];
                Connectivity.SQLCalls.sqlQuery("DELETE FROM MasterCategoryConnectors WHERE ConnectorType=" + CurrentChildTableID + " AND MasterID=" + masterID.ToString() + " AND ConnectorID=" + childID.ToString());
                if (hideChildConnectedChk.Checked == true)
                {
                    ListViewItem itm = new ListViewItem();
                    itm.Text = MasterLVI.Items.Count.ToString();
                    itm.SubItems.Add(masterID.ToString());
                    itm.SubItems.Add(masterName);
                    if (MasterCategoryData.ContainsKey(masterID) == false)
                    {
                        itm.BackColor = ListViewBackColor;
                        itm.ForeColor = ListViewForeColor;
                        MasterLVI.Items.Add(itm);
                    }
                    //remember, the cache needs to be updated as well...
                    CategoryInfo ci = new CategoryInfo();
                    ci.CategoryID = masterID;
                    ci.CategoryName = masterName;
                    if (MasterCategoryData.ContainsKey(masterID)==false)
                    { 
                    MasterCategoryData.Add(masterID, ci);
                        }
                }
                if (hideMasterConnectedChk.Checked==true)
                {
                    ListViewItem itm = new ListViewItem();
                    itm.Text = ChildLVI.Items.Count.ToString();
                    itm.SubItems.Add(childID.ToString());                    
                    itm.SubItems.Add(childName);
                    if (ChildCategoryData.ContainsKey(childID) == false)
                    {
                        itm.BackColor = ListViewBackColor;
                        itm.ForeColor = ListViewForeColor;
                        ChildLVI.Items.Add(itm);
                    }
                    CategoryInfo ci = new CategoryInfo();
                    ci.CategoryID = childID;
                    ci.CategoryName = childName;
                    if (ChildCategoryData.ContainsKey(childID) == false)
                    {
                        ChildCategoryData.Add(childID, ci);
                    }
                }
            }
            ResetConnectionsLVI();
            LoadCurrentConnectors();
            ConnectedCountLbl.Text = "(" + connectionsLVI.Items.Count.ToString() + ")";

        }

        private void MasterLVI_ColumnClick(object sender, ColumnClickEventArgs e)
        {/*
            int colNumber = e.Column;
            if (MasterLVI.Sorting == SortOrder.Ascending)
            {
                MasterLVI.Sorting = SortOrder.Descending;
            }
            else if (MasterLVI.Sorting == SortOrder.Descending)
            {
                MasterLVI.Sorting = SortOrder.Ascending;
            }
            else
            {
                MasterLVI.Sorting = SortOrder.Ascending;
            }
            MasterLVI.ListViewItemSorter = new ListViewItemComparer(colNumber);
          */
        }

        private void ChildLVI_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            /*
            int colNumber = e.Column;
            if (ChildLVI.Sorting == SortOrder.Ascending)
            {
                ChildLVI.Sorting = SortOrder.Descending;
            }
            else if (ChildLVI.Sorting == SortOrder.Descending)
            {
                ChildLVI.Sorting = SortOrder.Ascending;
            }
            else
            {
                ChildLVI.Sorting = SortOrder.Ascending;
            }*/
        }

        private void hideChildHighPriceChk_CheckedChanged(object sender, EventArgs e)
        {
            childNeedsRefresh = true;
            WriteChildToLVI();
        }

        private void connectionsLVI_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            int Col = e.Column;
            connectionsLVI.ListViewItemSorter = new ListViewItemComparer(Col);
        }

        private List<ListViewItem> PreviouslySelectedConnectionsLVI = new List<ListViewItem>();
        private void connectionsLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in PreviouslySelectedConnectionsLVI)
            {
                lvi.BackColor = ListViewBackColor;
                lvi.ForeColor = ListViewForeColor;
            }
            PreviouslySelectedConnectionsLVI.Clear();
            foreach (ListViewItem lvi in connectionsLVI.SelectedItems)
            {
                lvi.BackColor = ListViewHighlightBackColor;
                lvi.ForeColor = ListViewHighlightForeColor;
                PreviouslySelectedConnectionsLVI.Add(lvi);
            }
        }




    }
}
