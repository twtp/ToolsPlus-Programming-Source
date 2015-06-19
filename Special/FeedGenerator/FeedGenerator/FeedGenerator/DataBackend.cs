using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FeedGenerator
{
    public partial class DataBackend : Form
    {
        public DataBackend()
        {
            InitializeComponent();
        }

        private void DataBackend_Load(object sender, EventArgs e)
        {
            LoadFeedLVI();
            SetupDataLocationLVI();
            feedsLVI.MouseDoubleClick += new MouseEventHandler(feedsLVI_MouseDoubleClick);
            feedsLVI.ColumnClick += new ColumnClickEventHandler(feedsLVI_ColumnClick);
        }
        void feedsLVI_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            ListViewSorter Sorter = new ListViewSorter();
            feedsLVI.ListViewItemSorter = Sorter;
            if (!(feedsLVI.ListViewItemSorter is ListViewSorter))
                return;
            Sorter = (ListViewSorter)feedsLVI.ListViewItemSorter;

            if (Sorter.LastSort == e.Column)
            {
                if (feedsLVI.Sorting == SortOrder.Ascending)
                    feedsLVI.Sorting = SortOrder.Descending;
                else
                    feedsLVI.Sorting = SortOrder.Ascending;
            }
            else
            {
                feedsLVI.Sorting = SortOrder.Descending;
            }
            Sorter.ByColumn = e.Column;

            feedsLVI.Sort();

        }
        void feedsLVI_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int feedIndex = feedsLVI.SelectedIndices[0];
            ListViewItem tmpItm = feedsLVI.Items[feedIndex];
            string FeedName = tmpItm.Text;
            string FeedActive = tmpItm.SubItems[1].Text;
            if (FeedActive.ToUpper() == "FALSE")
            {
                Connectivity.SQLCalls.sqlQuery("UPDATE ShoppingEngineFeeds SET Active = 1 WHERE Feed='" + FeedName + "'");
                List<string> verify = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT Active FROM ShoppingEngineFeeds WHERE Feed='" + FeedName + "'")).Rows;
                if (verify.Count > 0)
                {
                    if (verify[0].Split('|')[0].ToUpper() != "TRUE")
                    {
                        MessageBox.Show("Error verifying setting " + FeedName + " to Active");
                    }
                    else
                    {
                        feedsLVI.Items[feedIndex].SubItems[1].Text = "True";
                        feedsLVI.Items[feedIndex].BackColor = Color.MediumSeaGreen;
                    }
                }
                else
                {
                    MessageBox.Show("Error trying to set " + FeedName + " to Active");
                }

            }
            else
            {
                Connectivity.SQLCalls.sqlQuery("UPDATE ShoppingEngineFeeds SET Active = 0 WHERE Feed='" + FeedName + "'");
                List<string> verify = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT Active FROM ShoppingEngineFeeds WHERE Feed='" + FeedName + "'")).Rows;
                if (verify.Count > 0)
                {
                    if (verify[0].Split('|')[0].ToUpper() != "FALSE")
                    {
                        MessageBox.Show("Error verifying setting " + FeedName + " to Inactive");
                    }
                    else
                    {
                        feedsLVI.Items[feedIndex].SubItems[1].Text = "False";
                        feedsLVI.Items[feedIndex].BackColor = Color.PaleVioletRed;
                    }
                }
                else
                {
                    MessageBox.Show("Error trying to set " + FeedName + " to Inactive");
                }
            }
            //MessageBox.Show("Feed Name: " + FeedName + "   -  Feed Active: " + FeedActive + ".");
            //feedsLVI.Items[feedIndex].BackColor = Color.MediumPurple;
        }
        private void SetupDataLocationLVI()
        {
            dataLocationLVI.Items.Clear();
            dataLocationLVI.Columns.Clear();
            dataLocationLVI.Clear();
            dataLocationLVI.GridLines = true;
            dataLocationLVI.FullRowSelect = true;
            dataLocationLVI.View = View.Details;
            dataLocationLVI.MultiSelect = false;
            dataLocationLVI.Refresh();

            ColumnHeader colID = new ColumnHeader();
            ColumnHeader colName = new ColumnHeader();
            ColumnHeader colData = new ColumnHeader();
            colID.Text = "ID";
            colName.Text = "Column Name";
            colData.Text = "Code Behind";

            dataLocationLVI.Columns.Add(colID);
            dataLocationLVI.Columns.Add(colName);
            dataLocationLVI.Columns.Add(colData);
        }
        private void SetupFeedLVI()
        {
            feedsLVI.Items.Clear();
            feedsLVI.Columns.Clear();
            feedsLVI.Clear();
            feedsLVI.View = View.Details;
            feedsLVI.GridLines = true;
            feedsLVI.FullRowSelect = true;
            feedsLVI.MultiSelect = false;


            ColumnHeader feedName = new ColumnHeader();
            ColumnHeader feedActive = new ColumnHeader();
            feedName.Text = "Feed Name";
            feedName.Width = (int)(feedsLVI.Width * 0.6f);
            feedActive.Text = "Active";
            feedsLVI.Columns.Add(feedName);
            feedsLVI.Columns.Add(feedActive);
            feedsLVI.Refresh();
        }
        private void LoadFeedLVI()
        {
            SetupFeedLVI();
            List<string> feedInfo = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT Feed,Active FROM FeedMaster")).Rows;
            foreach (string feed in feedInfo)
            {
                string FeedName = feed.Split('|')[0];
                string FeedActive = feed.Split('|')[1];
                ListViewItem feedItm = new ListViewItem();
                feedItm.Text = FeedName;
                feedItm.SubItems.Add(FeedActive);
                if (FeedActive.ToUpper() == "FALSE")
                {
                    feedItm.BackColor = Color.PaleVioletRed;
                }
                else
                {
                    feedItm.BackColor = Color.MediumSeaGreen;
                }

                feedsLVI.Items.Add(feedItm);

            }
        }
        public class ListViewSorter : System.Collections.IComparer
        {
            public int Compare(object o1, object o2)
            {
                if (!(o1 is ListViewItem))
                    return (0);
                if (!(o2 is ListViewItem))
                    return (0);

                ListViewItem lvi1 = (ListViewItem)o2;
                string str1 = lvi1.SubItems[ByColumn].Text;
                ListViewItem lvi2 = (ListViewItem)o1;
                string str2 = lvi2.SubItems[ByColumn].Text;

                int result;
                if (lvi1.ListView.Sorting == SortOrder.Ascending)
                    result = String.Compare(str1, str2);
                else
                    result = String.Compare(str2, str1);

                LastSort = ByColumn;

                return (result);
            }


            public int ByColumn
            {
                get { return Column; }
                set { Column = value; }
            }
            int Column = 0;

            public int LastSort
            {
                get { return LastColumn; }
                set { LastColumn = value; }
            }
            int LastColumn = 0;
        }

        private void feedsLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            SetupDataLocationLVI();
            try
            {
                int index = feedsLVI.SelectedIndices[0];
            }
            catch { }
            MessageBox.Show("Selecting this listview in any way does nothing at the moment.");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetupDataLocationLVI();
            List<string> res = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT ID,FeedColumn,DataLocation FROM FeedColumns")).Rows;
            foreach (string data in res)
            {
                string ColumnID = data.Split('|')[0];
                string ColumnName = data.Split('|')[1];
                if (ColumnName == "description")
                {
                }
                
                    
                string ColumnData = data.Split('|')[2];
                ListViewItem tmpItm = new ListViewItem();
                tmpItm.Text = ColumnID;
                tmpItm.SubItems.Add(ColumnName);
                tmpItm.SubItems.Add(ColumnData);
                dataLocationLVI.Items.Add(tmpItm);
            }

        }

        private void dataLocationLVI_SelectedIndexChanged(object sender, EventArgs e)
        {
            try { MessageBox.Show(dataLocationLVI.Items[dataLocationLVI.SelectedIndices[0]].SubItems[2].Text); }
            catch { }
        }

        private void addNoDuplicatesToList(string newListItem, ref ListBox lstBox)
        {
            bool test = false;
            foreach (string s in lstBox.Items)
            {
                if (s == newListItem)
                {
                    test = true;
                    break;
                }
            }
            if (test == false)
            {
                lstBox.Items.Add(newListItem);
            }
        }
        private string createFromQueryLine()
        {
            try
            {
                string FromQuery = "FROM InventoryMaster ";
               /* foreach (string str in tablesLst.Items)
                {
                    if (str.ToUpper().Contains("INVENTORYCOMPONENTMAP") == true)
                    {
                        FromQuery += "LEFT JOIN " + str + " ON " + str + ".ItemNumber = InventoryMaster.ItemNumber ";                        
                    }
                    if (str.ToUpper().Contains("INVENTORYCOMPONENTS") == true)
                    {
                        FromQuery += "LEFT JOIN " + str + " ON " + str + ".ID = InventoryComponentMap.ComponentID ";
                    }
                    if (str.ToUpper().Contains("INVENTORYCOMPONENTBARCODES") == true)
                    {
                        FromQuery += "LEFT JOIN " + str + " ON " + str + ".ComponentID = InventoryComponentMap.ComponentID ";
                    }

                }*/
                FromQuery += "LEFT JOIN InventoryComponentMap ON InventoryComponentMap.ItemNumber = InventoryMaster.ItemNumber ";
                FromQuery += "LEFT JOIN InventoryComponents ON InventoryComponents.ID = InventoryComponentMap.ComponentID ";
                FromQuery += "LEFT JOIN InventoryComponentBarcodes ON InventoryComponentBarcodes.ComponentID = InventoryComponentMap.ComponentID ";
                FromQuery += "LEFT JOIN PartNumbers ON PartNumbers.ItemNumber=InventoryMaster.ItemNumber ";
                FromQuery += "LEFT JOIN PartNumberSpecials ON PartNumberSpecials.ItemNumber=InventoryMaster.ItemNumber ";
                FromQuery += "LEFT JOIN Specials ON Specials.SpecialName=PartNumberSpecials.SpecialName ";                
                FromQuery += "LEFT JOIN PartNumberGroupLines ON PartNumberGroupLines.ItemNumber=InventoryMaster.ItemNumber ";
                FromQuery += "LEFT JOIN PartNumberGroupMaster ON PartNumberGroupMaster.ID=PartNumberGroupLines.GroupID ";
                FromQuery += "LEFT JOIN ProductLine ON ProductLine.ProductLine=InventoryMaster.ProductLine ";
                FromQuery += "LEFT JOIN PartNumberWebPaths ON PartNumberWebPaths.ItemNumber=InventoryMaster.ItemNumber ";
                FromQuery += "LEFT JOIN WebPaths ON WebPaths.ID=PartNumberWebPaths.WebPathID ";
                FromQuery += "LEFT JOIN MasterCategoryConnectors ON MasterCategoryConnectors.MasterID=WebPaths.ID ";
                FromQuery += "LEFT JOIN GoogleCategories ON GoogleCategories.ID = MasterCategoryConnectors.ConnectorID ";

                FromQuery += "LEFT JOIN WebOffload ON WebOffload.RealItemNumber=InventoryMaster.ItemNumber ";



                return FromQuery;
            }
            catch
            {
                return "[Errored!]";
            }
        }
        private void queryRefreshBtn_Click(object sender, EventArgs e)
        {
            List<string> res = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT ID,FeedColumn,DataLocation FROM FeedColumns")).Rows;
            List<string> tablesToJoin = new List<string>();
            foreach (string data in res)
            {
                //string ColumnID = data.Split('|')[0];
                //string ColumnName = data.Split('|')[1];
                string ColumnData = data.Split('|')[2];


                string[] segments = ColumnData.Split(' ');
                bool shouldAdd = true;
                foreach (string seg in segments)
                {
                    if (seg.Contains('.') == true)
                    {
                        string table = seg.Split('.')[0];
                        if (table.Contains("(") == true)
                        {
                            table = table.Split('(')[1];
                        }
                        if (table.Contains(")") == true)
                        {
                            table = table.Split(')')[0];
                        }
                        if (table.Contains("'") == true | table.Contains(":") == true |
                            table.ToUpper().Contains("VARCHAR") == true | table.ToUpper().Contains("CONVERT") == true |
                            table.ToUpper().Contains("REPLACE") == true)
                        {
                            shouldAdd = false;
                        }
                        if (shouldAdd == true)
                        {
                            tablesToJoin.Add(table);
                        }
                    }
                }


                //Added these in case, but NONE SHOULD BE USED!!!
                if (ColumnData.ToUpper().Contains("FROM ") == true)
                {
                    tablesToJoin.Add(ColumnData.Split(new string[] { "FROM " }, StringSplitOptions.None)[1].Split(' ')[0]);
                }
                if (ColumnData.ToUpper().Contains("INNER JOIN ") == true)
                {
                    tablesToJoin.Add(ColumnData.Split(new string[] { "INNER JOIN " }, StringSplitOptions.None)[1].Split(' ')[0]);
                }
                if (ColumnData.ToUpper().Contains("LEFT JOIN ") == true)
                {
                    tablesToJoin.Add(ColumnData.Split(new string[] { "LEFT JOIN " }, StringSplitOptions.None)[1].Split(' ')[0]);
                }
                //--------------------------------------------


                foreach (string s in tablesToJoin)
                {
                    addNoDuplicatesToList(s, ref tablesLst);
                    queryFromLineTxt.Text = createFromQueryLine();
                    //tablesLst.Items.Add(s);
                }


            }
        }

    }
}
