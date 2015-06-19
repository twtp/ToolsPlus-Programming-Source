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
    public partial class Form1 : Form
    {
        DataBackend dbFrm = new DataBackend();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label2.Text = "\uD83D\uDD0D" + " Search";
            LoadFeedLVI();
            SetupStatusLVI();
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
                Connectivity.SQLCalls.sqlQuery("UPDATE FeedMaster SET Active = 1 WHERE Feed='" + FeedName + "'");
                List<string> verify = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT Active FROM FeedMaster WHERE Feed='" + FeedName + "'")).Rows;
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
                Connectivity.SQLCalls.sqlQuery("UPDATE FeedMaster SET Active = 0 WHERE Feed='" + FeedName + "'");
                List<string> verify = ((Connectivity.SQLCalls.QueryResults)Connectivity.SQLCalls.sqlQuery("SELECT Active FROM FeedMaster WHERE Feed='" + FeedName + "'")).Rows;
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
            feedName.Width = (int)(feedsLVI.ClientSize.Width * 0.5f);
            feedActive.Text = "Active";
            feedActive.Width = feedsLVI.ClientSize.Width - feedName.Width - 17;
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
        public void SetupStatusLVI()
        {
            statusLVI.Items.Clear();
            statusLVI.Columns.Clear();
            statusLVI.Clear();
            statusLVI.View = View.Details;
            statusLVI.FullRowSelect = true;
            statusLVI.GridLines = true;
            statusLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader statusCol = new ColumnHeader();
            ColumnHeader dateCol = new ColumnHeader();
            ColumnHeader userCol = new ColumnHeader();
            idCol.Width = 1;
            idCol.Text = "";
            statusCol.Text = "Status";
            statusCol.Width = statusLVI.ClientSize.Width / 2;
            dateCol.Text = "TimeStamp";
            dateCol.Width = (statusLVI.ClientSize.Width - statusCol.Width) / 2;
            userCol.Text = "Username";
            userCol.Width = (statusLVI.ClientSize.Width - statusCol.Width - dateCol.Width) - 15;
            statusLVI.Columns.Add(idCol);
            statusLVI.Columns.Add(statusCol);
            statusLVI.Columns.Add(dateCol);
            statusLVI.Columns.Add(userCol);

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

        private void createSingleFeedBtn_Click(object sender, EventArgs e)
        {
            

            int index = feedsLVI.SelectedIndices[0];
            string FeedName = "Feeds_" + feedsLVI.Items[index].Text;
            string FeedActive = feedsLVI.Items[index].SubItems[1].Text;
            if (FeedActive.ToUpper() == "FALSE")
            {
                MessageBox.Show("The selected feed is not active so cannot create. Double click a feed to Activate/DeActivate.");
                return;
            }
            else
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = statusLVI.Items.Count.ToString();
                lvi.SubItems.Add("Manual Creation Of " + FeedName.Split('_')[1] + " Feed.");
                lvi.SubItems.Add(DateTime.Now.ToString());
                lvi.SubItems.Add("Temp User");
                statusLVI.Items.Add(lvi);
                statusLVI.Refresh();
                //ShoppingFeeds.ShoppingFeed SF = ShoppingFeeds.CreateFeed(FeedName, syncImagesChk.Checked, sendFeedsChk.Checked, createKeywordsChk.Checked);
                
                Connectivity.SQLCalls.QueryResults results = Connectivity.SQLCalls.sqlQuery("SELECT * FROM " + FeedName);
                previewLVI.Items.Clear();
                previewLVI.Columns.Clear();
                previewLVI.Clear();
                previewLVI.View = View.Details;
                previewLVI.FullRowSelect = true;
                previewLVI.GridLines = true;
                previewLVI.Refresh();
                          
                ColumnHeader id = new ColumnHeader();
                id.Text= "";
                id.Width = 1;
                previewLVI.Columns.Add(id);
                foreach (string col in results.Columns)
                {
                    ColumnHeader column = new ColumnHeader();
                    column.Text = col;
                    column.Width = (previewLVI.Width / results.Columns.Count) - 2;
                    previewLVI.Columns.Add(col);
                }
                int count = 0;
                foreach (string line in results.Rows)
                {
                    ListViewItem lvi2 = new ListViewItem();
                    lvi2.Text = count.ToString();

                    foreach (string itm in line.Split('|'))
                    {

                        lvi2.SubItems.Add(itm);
                        
                    }
                    previewLVI.Items.Add(lvi2);
                    
                    count++;
                }
                
            }
        }

        private void feedsLVI_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void debugBtn_Click(object sender, EventArgs e)
        {

            dbFrm.Show();
        }
        public List<int> ColumnOrder = new List<int>();
        private void exportCSV_Click(object sender, EventArgs e)
        {
            ColumnOrder.Clear();

            for (int x = 0; x < previewLVI.Columns.Count; x++)
            {
                ColumnOrder.Add(previewLVI.Columns[x].DisplayIndex);
            }
            
            string filePath = "";
            saveFileDialog1.Filter = "*.csv|*.csv|*.*|*.*";

            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.Yes | dr == System.Windows.Forms.DialogResult.OK)
            {
                filePath = saveFileDialog1.FileName;
                System.IO.File.WriteAllText(filePath, "");
            }
            else
            {
                MessageBox.Show("Aborted Saving.");
                return;
            }


            string fileout = "";
           /* foreach (ColumnHeader col in previewLVI.Columns)
            {
                if (col.Text != "")
                {
                    fileout += col.Text.Replace("\t","     ") + "\t";
                }
                
            }*/
            for (int x = 1; x < previewLVI.Columns.Count; x++)
            {
                for (int y = 0; y < ColumnOrder.Count; y++)
                {
                    if (ColumnOrder[y] == x)
                    {
                        fileout += previewLVI.Columns[y].Text.Replace("\t", "     ") + "\t";
                        break;
                    }
                }
            }
            fileout = fileout.Substring(0, fileout.Length - 1);
            fileout += "\r\n";
            int count = 0;
            foreach (ListViewItem lvi in previewLVI.Items)
            {
                for (int x = 1; x < previewLVI.Columns.Count; x++)
                {
                    for (int y = 0; y < ColumnOrder.Count; y++)
                    {
                        if (ColumnOrder[y] == x)
                        {
                            fileout += lvi.SubItems[y].Text.Replace("\t", "     ") + "\t";
                            break;
                        }
                    }
                }
                //fileout.TrimEnd('\t');
                //fileout.TrimEnd();
                fileout = fileout.Substring(0,fileout.Length - 1);
                fileout += "\r\n";
                count++;

                System.IO.File.AppendAllText(filePath, fileout);
                fileout = "";
            }

 
            MessageBox.Show("Complete");



        }

        private void exportXML_Click(object sender, EventArgs e)
        {
            int columns = previewLVI.Columns.Count;
            string filePath = "";
            saveFileDialog1.Filter = "*.xml|*.xml|*.*|*.*";
            string xmlHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\r\n";
            string xmlDataTag = "<products>";
            string xmlLineTag = " <product>";
            

            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.Yes | dr == System.Windows.Forms.DialogResult.OK)
            {
                filePath = saveFileDialog1.FileName;
            }
            else
            {
                MessageBox.Show("Aborted Saving.");
                return;
            }
            string Header = xmlHeader + xmlDataTag + "\r\n";
            //System.IO.File.AppendAllText(filePath, Header);
            string results = "";
            foreach (ListViewItem lvi in previewLVI.Items)
            {
                string xmlItem = xmlLineTag + "\r\n";
                for (int x = 0; x < columns-1; x++)
                {
                    xmlItem += "  <" + previewLVI.Columns[x+1].Text + ">";
                    xmlItem += lvi.SubItems[x+1].Text;
                    xmlItem += "</" + previewLVI.Columns[x+1].Text + ">\r\n";
                }
                xmlItem += xmlLineTag.Replace("<", "</") + "\r\n";
                results += xmlItem;
                //System.IO.File.AppendAllText(filePath, xmlItem);
            }
            string Footer = xmlDataTag.Replace("<", "</");
            //System.IO.File.AppendAllText(filePath, Footer);
            System.IO.File.WriteAllText(filePath, Header + results + Footer);

        }

        private void previewLVI_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void previewLVI_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int offset = previewLVI.SelectedIndices[0];

            ModifyLine ml = new ModifyLine();
            ml.CreateFormArray(previewLVI.Columns, previewLVI.SelectedItems[0],offset,this);
            ml.ShowDialog();
            

            //previewLVI.Items.Remove(ml.currentItem);
            //previewLVI.Items.Add(ml.modifiedItem);
        }

        private void previewLVI_ColumnReordered(object sender, ColumnReorderedEventArgs e)
        {
            //List<int> tmpColumnOrder = new List<int>();

            

        }

        private void button1_Click(object sender, EventArgs e)
        {
           
        }

        private void newFeedBtn_Click(object sender, EventArgs e)
        {
            NewFeed nf = new NewFeed();
            nf.OnFeedCreation += nf_OnFeedCreation;
            nf.SetCreationMode();
            nf.ShowDialog();
        }

        void nf_OnFeedCreation(object sender, EventArgs e)
        {          
            //event: when feed created, refresh feeds list.
            LoadFeedLVI();
        }



        private void delFeedBtn_Click(object sender, EventArgs e)
        {
            if (feedsLVI.SelectedIndices[0] > -1)
            {

                if (feedsLVI.Items[feedsLVI.SelectedIndices[0]].SubItems[1].Text == "True")
                {
                    MessageBox.Show("The feed must be disabled before you can delete it.");
                    return;
                }


                DialogResult res = MessageBox.Show("Are you sure you want to delete the feed \"" + feedsLVI.Items[feedsLVI.SelectedIndices[0]].SubItems[0].Text + "\"?","Delete Feed",MessageBoxButtons.YesNo);
                if (res == System.Windows.Forms.DialogResult.OK | res == System.Windows.Forms.DialogResult.Yes)
                {
                    Connectivity.SQLCalls.sqlQuery("DROP VIEW Feeds_" + feedsLVI.Items[feedsLVI.SelectedIndices[0]].SubItems[0].Text);
                    Connectivity.SQLCalls.sqlQuery("DELETE FROM FeedMaster WHERE Feed='" +feedsLVI.Items[feedsLVI.SelectedIndices[0]].SubItems[0].Text + "'");
                    LoadFeedLVI();
                }
            }
            else
            {
                MessageBox.Show("Please select a feed first!");
                return;
            }
        }

        private void editFeedBtn_Click(object sender, EventArgs e)
        {
            if (feedsLVI.SelectedIndices[0] > -1)
            {
                string FeedName = feedsLVI.Items[feedsLVI.SelectedIndices[0]].SubItems[0].Text;
                NewFeed editFeed = new NewFeed();
                editFeed.SetUpdateMode(FeedName);
                editFeed.OnFeedCreation += editFeed_OnFeedCreation;
                editFeed.ShowDialog();
            }
            else
            {
                MessageBox.Show("Please select a feed to edit first!");
                return;
            }
        }

        void editFeed_OnFeedCreation(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }


    }
}
