using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

namespace CategoryConnector
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void RefreshForm(object sender, EventArgs e)
        {
            //masterCategoriesLbl
            //connectorsCmb
            //childCategoriesLbl
            //masterLVI
            //connectBtn
            //childLVI
            //hideHighPriceChk
            //checkBox2
            //checkBox3
            //currentConnectionsLbl
            //connectorsLVI
            //disconnectBtn

            masterCategoriesLbl.Top = 5;
            masterCategoriesLbl.Left = 5;
            masterLVI.Top = masterCategoriesLbl.Bottom + 2;
            masterLVI.Left = masterCategoriesLbl.Left;
            masterLVI.Width = (this.Width / 2) - (connectBtn.Width/2)-2;
            masterLVI.Height = this.Height / 2;//this.Height - masterCategoriesLbl.Height - connectorsLVI.Height - disconnectBtn.Height - currentConnectionsLbl.Height;
            connectBtn.Left = masterLVI.Right + 2;
            connectBtn.Top = masterLVI.Top + 5;
            hideHighPriceChk.Left = masterLVI.Right + 2;
            hideHighPriceChk.Top = connectBtn.Bottom + 3;
            hideConnectedChildsChk.Left = hideHighPriceChk.Left;
            hideConnectedChildsChk.Top = hideHighPriceChk.Bottom + 2;
            childCategoriesChk.Left = hideHighPriceChk.Left;
            childCategoriesChk.Top = hideConnectedChildsChk.Bottom + 2;
            globalMastersChk.Left = hideHighPriceChk.Left;
            globalMastersChk.Top = childCategoriesChk.Bottom + 2;
            childMastersChk.Left = hideHighPriceChk.Left;
            childMastersChk.Top = globalMastersChk.Bottom + 2;
            childLVI.Left = connectBtn.Right + 10;
            childLVI.Top = masterLVI.Top;
            childLVI.Width = masterLVI.Width-50;
            childLVI.Height = masterLVI.Height;
            currentConnectionsLbl.Top = masterLVI.Bottom + 2;
            currentConnectionsLbl.Left = masterLVI.Left + 2;
            connectorsLVI.Top = currentConnectionsLbl.Bottom + 2;
            connectorsLVI.Left = currentConnectionsLbl.Left;
            connectorsLVI.Width = this.Width - (connectorsLVI.Left * 2) - 25;
            disconnectBtn.Top = this.Height - disconnectBtn.Height-50;
            disconnectBtn.Left = connectorsLVI.Left + ((connectorsLVI.Width / 2) - (disconnectBtn.Width / 2));           
            connectorsLVI.Height = disconnectBtn.Top - connectorsLVI.Top;//this.Height - currentConnectionsLbl.Bottom - 50;
            
            connectorsCmb.Left = childLVI.Left;
            connectorsCmb.Top = masterCategoriesLbl.Top;
            childCategoriesLbl.Top = connectorsCmb.Top;
            childCategoriesLbl.Left = connectorsCmb.Right + 10;

        }
        private void SetToolTips()
        {
            //masterCategoriesLbl
            //connectorsCmb
            //childCategoriesLbl
            //masterLVI
            //connectBtn
            //childLVI
            //hideHighPriceChk
            //checkBox2
            //checkBox3
            //currentConnectionsLbl
            //connectorsLVI
            //disconnectBtn
            ToolTip tt = new ToolTip();
            tt.SetToolTip(childCategoriesChk, "This checkbox enables/disables the Connectors list from showing all connected connectors, or just the ones from the currently selected Child.");
            tt.SetToolTip(hideConnectedChildsChk, "This checkbox hides/unhides child categories that are already connected.");
            tt.SetToolTip(hideHighPriceChk, "This checkbox hides/unhides high priced categories in the child list. Currently supported by EBay Categories only.");
            tt.SetToolTip(connectBtn, "This button creates a connection between the Master category on the left and the Child category on the right. Multiple categories can be selected.");
            tt.SetToolTip(disconnectBtn, "This button removes the selected Connector from the connectors list.");
            tt.SetToolTip(connectorsCmb, "This selection control allows you to choose a specific Child category list to load.");
            tt.SetToolTip(globalMastersChk, "This checkbox hides/unhides Master categories that have been used in connecting any category");
            tt.SetToolTip(childMastersChk, "This checkbox hides/unhides Master categories that have been used in connecting to the current Child categories.");


        }
        public class ListViewItemComparer : IComparer
        {
            private int col = 0;

            public ListViewItemComparer(int column)
            {
                col = column;
            }
            public int Compare(object x, object y)
            {
                return String.Compare(((ListViewItem)x).SubItems[col].Text, ((ListViewItem)y).SubItems[col].Text);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetToolTips();
            LoadMasterLVI();            
            ResetConnectorsCmb();
            ResetConnectorLVI();
            LoadConnectorsLVI(hideConnectedChildsChk.Checked, sender, e);
            masterLVI.ListViewItemSorter = new ListViewItemComparer(1);
            childLVI.ListViewItemSorter = new ListViewItemComparer(1);
            connectorsLVI.ListViewItemSorter = new ListViewItemComparer(1);
            this.Resize += new EventHandler(RefreshForm);
        }

        private void ResetConnectorLVI()
        {
            connectorsLVI.Items.Clear();
            connectorsLVI.Columns.Clear();
            connectorsLVI.Clear();
            connectorsLVI.Refresh();

            ColumnHeader MasterCat = new ColumnHeader();
            ColumnHeader ChildCatType = new ColumnHeader();
            ColumnHeader ChildCat = new ColumnHeader();

            MasterCat.Text = "Master Category";
            ChildCatType.Text = "Child Type";
            ChildCat.Text = "Child Category";

            MasterCat.Width = (connectorsLVI.Width / 2) - (connectorsLVI.Width / 20);
            ChildCat.Width = (connectorsLVI.Width / 2) - (connectorsLVI.Width / 20);
            ChildCatType.Width = (connectorsLVI.Width / 10);

            connectorsLVI.Columns.Add(MasterCat);
            connectorsLVI.Columns.Add(ChildCatType);
            connectorsLVI.Columns.Add(ChildCat);

        }

        private void ResetMasterLVI()
        {
            masterLVI.Items.Clear();
            masterLVI.Columns.Clear();
            masterLVI.Clear();
            masterLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader nameCol = new ColumnHeader();
            idCol.Width = 0;
            nameCol.Width = masterLVI.Width;
            idCol.Text = "Master ID";
            nameCol.Text = "Master Category";
            masterLVI.Columns.Add(idCol);
            masterLVI.Columns.Add(nameCol);

        }
        private void ResetChildLVI()
        {
            childLVI.Items.Clear();
            childLVI.Columns.Clear();
            childLVI.Clear();
            childLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader nameCol = new ColumnHeader();
            idCol.Width = 0;
            nameCol.Width = childLVI.Width;
            idCol.Text = "Child ID";
            nameCol.Text = "Child Category";
            childLVI.Columns.Add(idCol);
            childLVI.Columns.Add(nameCol);
            childCategoriesLbl.Text = "Categories: ";
            childCategoriesLbl.Refresh();
        }
        private void ResetConnectorsCmb()
        {
            connectorsCmb.Items.Clear();
            connectorsCmb.Refresh();
            List<string> connectorsList = Connectivity.SQLCalls.sqlQuery("SELECT ConnectorType,ConnectorTableName FROM MasterCategoryConnectorsTypes ORDER BY ConnectorType");
            if (connectorsList.Count > 0)
            {
                foreach (string conn in connectorsList)
                {
                    connectorsCmb.Items.Add(conn.Split('|')[0] + ": " + conn.Split('|')[1]);
                }
            }

        }

        private void LoadConnectorsLVI(bool currentChildOnly,object sender, EventArgs e)
        {
            ResetConnectorLVI();
            //if (connectorsCmb.SelectedIndex == -1)
           // {
           //     return;
           // }
            string sqlQuery = "SELECT MasterID,ConnectorType,ConnectorID FROM MasterCategoryConnectors";
            if (currentChildOnly == true & connectorsCmb.SelectedIndex > -1)
            {
               sqlQuery += " WHERE ConnectorType=" + connectorsCmb.Items[connectorsCmb.SelectedIndex].ToString().Split(':')[0].Trim();
            }
            else if (currentChildOnly == false)
            {
                //do nothing additional
            }
            else
            {
                return;
            }

            List<string> connectors = Connectivity.SQLCalls.sqlQuery(sqlQuery);
            if (connectors.Count > 0)
            {
                foreach (string connector in connectors)
                {
                    string MasterID = connector.Split('|')[0];
                    string ConnectorType = connector.Split('|')[1];
                    string ChildID = connector.Split('|')[2];

                    string ChildName = "";
                    string MasterName = "";
                    string ConnectorName = "";
                    List<string> conns = Connectivity.SQLCalls.sqlQuery("SELECT ConnectorTableName FROM MasterCategoryConnectorsTypes WHERE ConnectorType=" + ConnectorType);
                    if (conns.Count > 0)
                    {
                        ConnectorName = conns[0].Split('|')[0];
                    }
                    else
                    {
                        return;
                    }
                    List<string> mast = Connectivity.SQLCalls.sqlQuery("SELECT Name FROM MasterCategories WHERE ID=" + MasterID);
                    if (mast.Count > 0)
                    {
                        MasterName = mast[0].Split('|')[0];
                    }
                    else
                    {
                        return;
                    }
                    List<string> childs = Connectivity.SQLCalls.sqlQuery("SELECT Name FROM " + ConnectorName + " WHERE ID=" + ChildID);
                    if (childs.Count > 0)
                    {
                        ChildName = childs[0].Split('|')[0];
                    }
                    else
                    {
                        return;
                    }

                    ListViewItem tmpItm = new ListViewItem();
                    tmpItm.Text = MasterName;
                    tmpItm.SubItems.Add(ConnectorName);
                    tmpItm.SubItems.Add(ChildName);

                    connectorsLVI.Items.Add(tmpItm);
                    
                    
                }
                currentConnectionsLbl.Text = "Current Connections: " + connectors.Count.ToString();
                currentConnectionsLbl.Refresh();
            }
            else
            {
                currentConnectionsLbl.Text = "Current Connections: " + connectors.Count.ToString();
                currentConnectionsLbl.Refresh();
            }

        }
        private void FilterConnectedChilds(object sender, EventArgs e)
        {
            
            if (hideConnectedChildsChk.Checked == true)
            {
                List<string> connectedChildCats = new List<string>();
                foreach (ListViewItem LVI in connectorsLVI.Items)
                {
                    connectedChildCats.Add(LVI.SubItems[2].Text);
                }
                for (int x = 0; x < childLVI.Items.Count; x++)
                {
                    foreach (string Cat in connectedChildCats)
                    {
                        if (childLVI.Items[x].SubItems[1].Text == Cat)
                        {
                            childLVI.Items[x].Remove();

                        }
                    }
                }

            }
            else
            {

                connectorsCmb_SelectedIndexChanged(sender, e);
                
            }

            childCategoriesLbl.Text = "Categories: " + childLVI.Items.Count.ToString();
            childCategoriesLbl.Refresh();

                
                
            
        }
        private void LoadMasterLVI()
        {
            ResetMasterLVI();
            List<string> masterCats = Connectivity.SQLCalls.sqlQuery("SELECT ID,Name FROM MasterCategories");
            if (masterCats.Count > 0)
            {
                foreach (string cat in masterCats)
                {
                    ListViewItem tmpItm = new ListViewItem();
                    tmpItm.Text = cat.Split('|')[0];
                    tmpItm.SubItems.Add(cat.Split('|')[1]);
                    masterLVI.Items.Add(tmpItm);
                }
                masterCategoriesLbl.Text = "Master Categories: " + masterLVI.Items.Count.ToString();
                masterCategoriesLbl.Refresh();
            }
            else
            {
                masterCategoriesLbl.Text = "Master Categories: 0";
                masterCategoriesLbl.Refresh();
            }
        }

        private void connectorsCmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            ResetChildLVI();
            string childTable = connectorsCmb.Items[connectorsCmb.SelectedIndex].ToString().Split(':')[1].Trim();
                  


            List<string> childData = Connectivity.SQLCalls.sqlQuery("SELECT ID,Name FROM " + childTable + (hideHighPriceChk.Checked==true ? " WHERE Name NOT LIKE '%$$$%'":""));
            if (childData.Count > 0)
            {
                foreach (string tmpChild in childData)
                {
                    ListViewItem tmpLVI = new ListViewItem();
                    tmpLVI.Text = tmpChild.Split('|')[0];
                    tmpLVI.SubItems.Add(tmpChild.Split('|')[1]);
                    childLVI.Items.Add(tmpLVI);
                }
                childCategoriesLbl.Text = "Categories: " + childData.Count.ToString();
            }
            LoadConnectorsLVI(childCategoriesChk.Checked,sender,e);
            childLVI.Focus();
            FilterMasterLVIForGlobalUsed(connectorsCmb.SelectedIndex + 1);
        }

        private void connectBtn_Click(object sender, EventArgs e)
        {
            ConnectSelectedCategories();
            FilterMasterLVIForGlobalUsed(connectorsCmb.SelectedIndex + 1);
            
        }

        private void ConnectSelectedCategories()
        {
            List<int> MasterIndexes = new List<int>();
            foreach (int index in masterLVI.SelectedIndices)
            {
                MasterIndexes.Add(index);
            }
            List<int> ChildIndexes = new List<int>();
            foreach (int index in childLVI.SelectedIndices)
            {
                ChildIndexes.Add(index);
            }

            foreach (int MI in MasterIndexes)
            {

                foreach (int CI in ChildIndexes)
                {
                    int ConnectorType = int.Parse(connectorsCmb.Items[connectorsCmb.SelectedIndex].ToString().Split(':')[0].Trim());
                    int MasterID = int.Parse(masterLVI.Items[MI].Text);
                    //string MasterName = masterLVI.Items[MI].Text;
                    int ChildID = int.Parse(childLVI.Items[CI].Text);
                    //string ChildName = childLVI.Items[CI].Text;

                    Connectivity.SQLCalls.sqlQuery("INSERT INTO MasterCategoryConnectors (MasterID,ConnectorType,ConnectorID) VALUES (" + MasterID.ToString() + "," + ConnectorType.ToString() + "," + ChildID.ToString() + ")");

                }

            }

            LoadConnectorsLVI(childCategoriesChk.Checked, null, null);
        }

        private void hideHighPriceChk_CheckedChanged(object sender, EventArgs e)
        {
            connectorsCmb_SelectedIndexChanged(sender, e);
        }

        private void hideConnectedChildsChk_CheckedChanged(object sender, EventArgs e)
        {
            FilterConnectedChilds(sender,e);
        }

        private void childCategoriesChk_CheckedChanged(object sender, EventArgs e)
        {
            LoadConnectorsLVI(childCategoriesChk.Checked,sender,e);
        }

        private void disconnectBtn_Click(object sender, EventArgs e)
        {            
            List<int> indexes = new List<int>();
            foreach (int index in connectorsLVI.SelectedIndices)
            {
                indexes.Add(index);
            }

            foreach (int index in indexes)
            {
                string MasterCategoryName = connectorsLVI.Items[index].Text;
                string ChildCategoryName = connectorsLVI.Items[index].SubItems[2].Text;
                string ConnectorTypeName = connectorsLVI.Items[index].SubItems[1].Text;
                string MasterCategoryID = "";
                string ChildCategoryID = "";
                string ConnectorTypeID = "";
                List<string> master = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM MasterCategories WHERE Name='" + MasterCategoryName + "'");
                if (master.Count > 0)
                {
                    MasterCategoryID = master[0].Split('|')[0];
                    List<string> child = Connectivity.SQLCalls.sqlQuery("SELECT ID FROM " + ConnectorTypeName + " WHERE Name='" + ChildCategoryName + "'");
                    if (child.Count > 0)
                    {
                        ChildCategoryID = child[0].Split('|')[0];
                        List<string> connector = Connectivity.SQLCalls.sqlQuery("SELECT ConnectorType FROM MasterCategoryConnectorsTypes WHERE ConnectorTableName='" + ConnectorTypeName + "'");
                        if (connector.Count > 0)
                        {
                            ConnectorTypeID = connector[0].Split('|')[0];
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                }
                else
                {
                    return;
                }

                Connectivity.SQLCalls.sqlQuery("DELETE FROM MasterCategoryConnectors WHERE MasterID=" + MasterCategoryID + " AND ConnectorType=" + ConnectorTypeID + " AND ConnectorID=" + ChildCategoryID);


            }

            LoadConnectorsLVI(childCategoriesChk.Checked, sender, e);
            FilterMasterLVIForGlobalUsed(connectorsCmb.SelectedIndex + 1);
        }

        private void FilterMasterLVIForGlobalUsed(int connectorCategoryID)
        {
            List<string> masterIDsInUse = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT MasterID FROM MasterCategoryConnectors WHERE ConnectorType=" + connectorCategoryID.ToString());
            int count = 0;
            bool found = false;
            for (int x = 0; x < masterLVI.Items.Count; x++)
            {
                foreach (string usedID in masterIDsInUse)
                {
                    if (usedID.Split('|')[0] == masterLVI.Items[x].Text)
                    {
                        count++;
                        //masterLVI.Items.RemoveAt(x);
                        masterLVI.Items[x].BackColor = Color.MediumSeaGreen;
                        //masterLVI.Refresh();
                        found = true;
                    }
                }
                if (found == false)
                {
                    masterLVI.Items[x].BackColor = Color.White;
                }
            }
            //MessageBox.Show(count.ToString());
            masterCategoriesLbl.Text = "Master Categories: " + masterLVI.Items.Count.ToString();
        }
        private void FilterMasterLVIForChildUsed()
        {
            string id = connectorsCmb.Items[connectorsCmb.SelectedIndex].ToString().Split(':')[0].Trim();

            List<string> masterIDsInUse = Connectivity.SQLCalls.sqlQuery("SELECT DISTINCT MasterID FROM MasterCategoryConnectors WHERE ConnectorType=" + id);
            for (int x = 0; x < masterLVI.Items.Count; x++)
            {
                foreach (string usedID in masterIDsInUse)
                {
                    if (usedID.Split('|')[0] == masterLVI.Items[x].Text)
                    {
                        masterLVI.Items[x].Remove();
                    }
                }
            }
            masterCategoriesLbl.Text = "Master Categories: " + masterLVI.Items.Count.ToString();
        }
        private void globalMastersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (globalMastersChk.Checked == true)
            {
                LoadMasterLVI();
                FilterMasterLVIForGlobalUsed(connectorsCmb.SelectedIndex +1 );
                
            }
            else
            {
                LoadMasterLVI();
            }
        }

        private void childMastersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (childMastersChk.Checked == true)
            {
                LoadMasterLVI();
                FilterMasterLVIForChildUsed();
            }
            else
            {
                LoadMasterLVI();
            }
        }


    }
}
