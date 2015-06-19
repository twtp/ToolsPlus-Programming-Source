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
    public partial class NewFeed : Form
    {

        public bool InCreationMode = false;
        public bool InEditMode = false;

        private string FromStatement = "";


        public event EventHandler OnFeedCreation;



        public NewFeed()
        {
            InitializeComponent();
        }

        private void itemNumberTxt_Leave(object sender, EventArgs e)
        {
            //MessageBox.Show("HERE");

        }

        private void NewFeed_Load(object sender, EventArgs e)
        {
            OnFeedCreation += NewFeed_OnFeedCreation;
            SetupAvailableLVI();
            SetupNewFeedLVI();
            newFeedLVI.AfterLabelEdit += newFeedLVI_AfterLabelEdit;
        }

        public void SetCreationMode()
        {
            InCreationMode = true;
            InEditMode = false;
            feedNameTxt.Enabled = true;
            feedSaveBtn.Visible = true;
            feedSaveBtn.BringToFront();
        }
        public void SetUpdateMode(string FeedName)
        {
            InEditMode = true;
            InCreationMode = false;
            feedNameTxt.Enabled = false;
            feedUpdateBtn.Visible = true;
            feedUpdateBtn.BringToFront();
            feedNameTxt.Text = FeedName;
            LoadFeed(FeedName);
        }
        private void LoadFeed(string FeedName)
        {

            Connectivity.SQLCalls.QueryResults colCount = Connectivity.SQLCalls.sqlQuery("SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE Table_Name='Feeds_" + FeedName + "'");
            if (colCount.Rows.Count <= 0)
            {
                MessageBox.Show("Cannot determine number of rows in this view. This should be more fatal...");
                return;
            }


            ColumnCodeBehind.Clear();
            List<string> SQLDataTypes = new List<string>();
            SQLDataTypes.Add("VARCHAR");
            SQLDataTypes.Add("INT");
            SQLDataTypes.Add("DATETIME");
            SQLDataTypes.Add("MONEY");

            Connectivity.SQLCalls.QueryResults qc = Connectivity.SQLCalls.sqlQuery("sp_helptext 'dbo.Feeds_" + FeedName + "'");
            string Query = "";
            if (qc.Rows.Count > 0)
            {
                foreach (string row in qc.Rows)
                {
                    Query += row;
                }
            }
            else
            {
                MessageBox.Show("Error: No column codebehind data found. This should be more fatal than it is...");
                return;
            }

            Query = Query.Split(new string[] { "SELECT DISTINCT" }, StringSplitOptions.None)[1];
            FromStatement = "FROM " + Query.Split(new string[] { "FROM " }, StringSplitOptions.None)[1];
            FromStatement = FromStatement.Replace("|", "");
            FromStatement = FromStatement.Replace("  ", "");
            Query = Query.Split(new string[] { "FROM " }, StringSplitOptions.None)[0];
            Query = Query.Replace("\r\n", "");
            Query = Query.Replace("|", "");

            List<string> tmpColumnCodeBehind = new List<string>();
            List<int> actualOffsets = new List<int>();
            bool finish = false;
            int lastColumnEnd = 0;
            int lastStartIndex = 0;
            int currentStartIndex = 0;
            int count = 0;
            bool lastLoopSuccess = false;
            //int LengthOfAS = 0;
            while (finish == false)
            {
                if (count == 95)
                { }
                if (lastColumnEnd > 0 & lastLoopSuccess == true)
                {
                    lastStartIndex = lastColumnEnd;
                }
                try
                {
                    currentStartIndex = Query.IndexOf(" AS ", lastStartIndex);
                    //LengthOfAS = lastStartIndex - currentStartIndex;
                }
                catch (Exception e)
                {
                    finish = true;
                    break;
                }
                string test = (Query.Substring(currentStartIndex).Split(',')[0]);
                bool checkDataType = false;
                foreach (string dType in SQLDataTypes)
                {
                    if (test.Contains(dType) == true)
                    {
                        checkDataType = true;
                        break;
                    }
                }
                if (checkDataType == false)
                {
                    lastLoopSuccess = true;
                    //find first comma after index, and that is the real end of the line
                    int indexOfEnd = Query.IndexOf(',', currentStartIndex);
                    //string result = Query.Substring(lastStartIndex + 4, (currentStartIndex + 4) - (lastStartIndex + 4));
                    string result = "";
                    if (lastColumnEnd > 0)
                    {
                        try
                        {
                            result = Query.Substring(lastColumnEnd, (indexOfEnd) - (lastColumnEnd));
                            result = result.Substring(0, result.Length - test.Length);
                        }
                        catch { result = "FAILURE"; break; }
                    }
                    else
                    {
                        try
                        {
                            result = Query.Substring(lastStartIndex, (indexOfEnd) - (lastStartIndex));
                            result = result.Substring(0, result.Length - test.Length);
                        }
                        catch { result = "FAILURE."; break; }
                    }
                    lastStartIndex = indexOfEnd;
                    lastColumnEnd = lastStartIndex + 1;// +result.Length;
                    ///MessageBox.Show(result);
                    ColumnCodeBehind.Add(result.Trim());
                    count++;
                    if (count >= totColumns - 1)
                    {
                        finish = true;
                        break;
                    }
                }
                else
                {
                    lastLoopSuccess = false;
                    lastStartIndex = currentStartIndex + 1;
                }
                //finish = true;
                //if (Query.Substring(currentStartIndex + 3))
            }
        }

        void NewFeed_OnFeedCreation(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            EventHandler handler = OnFeedCreation;            
        }

        void newFeedLVI_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
           
            int index = e.Item;
            if (e.Label.Contains(" ")==true & (e.Label.Contains("[")==false | e.Label.Contains("]")==false))
            {

                    newFeedLVI.Items[index].Text = "[" + e.Label + "]";
                    e.CancelEdit = true;
                
            }
        }
        
        private void SetupAvailableLVI()
        {
            availableLVI.Items.Clear();
            availableLVI.Columns.Clear();
            availableLVI.Clear();
            availableLVI.View = View.Details;
            availableLVI.FullRowSelect = true;
            availableLVI.GridLines = true;
            availableLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader nameCol = new ColumnHeader();
            ColumnHeader valueCol = new ColumnHeader();
            ColumnHeader codeCol = new ColumnHeader();
            idCol.Text = "";
            idCol.Width = 1;
            availableLVI.Columns.Add(idCol);
            nameCol.Text = "Column Name";
            nameCol.Width = availableLVI.ClientSize.Width / 2 - 8;
            availableLVI.Columns.Add(nameCol);
            valueCol.Text = "Column Data";
            valueCol.Width = availableLVI.ClientSize.Width / 2 - 8;            
            availableLVI.Columns.Add(valueCol);
            codeCol.Text = "Code Behind";
            codeCol.Width = 1;
            availableLVI.Columns.Add(codeCol);

        }

        private void SetupNewFeedLVI()
        {
            newFeedLVI.Items.Clear();
            newFeedLVI.Columns.Clear();
            newFeedLVI.Clear();
            newFeedLVI.View = View.Details;
            newFeedLVI.FullRowSelect = true;
            newFeedLVI.GridLines = true;
            newFeedLVI.LabelEdit = true;
            newFeedLVI.Refresh();

            ColumnHeader idCol = new ColumnHeader();
            ColumnHeader nameCol = new ColumnHeader();
            ColumnHeader exampleCol = new ColumnHeader();
            ColumnHeader codeCol = new ColumnHeader();
            idCol.Text = "";
            idCol.Width = 1;
            
            nameCol.Text = "Column Name";
            nameCol.Width = newFeedLVI.ClientSize.Width / 2 - 8;
            newFeedLVI.Columns.Add(nameCol);
            exampleCol.Text = "Example Data";
            exampleCol.Width = newFeedLVI.ClientSize.Width / 2 - 8;
            codeCol.Text = "Code Behind";
            codeCol.Width = 1;
            newFeedLVI.Columns.Add(exampleCol);
            newFeedLVI.Columns.Add(idCol);
            newFeedLVI.Columns.Add(codeCol);
        }

        public List<string> ColumnCodeBehind = new List<string>();
        private void GetColumnCodeBehind(int totColumns)
        {
            ColumnCodeBehind.Clear();
            List<string> SQLDataTypes = new List<string>();
            SQLDataTypes.Add("VARCHAR");
            SQLDataTypes.Add("INT");
            SQLDataTypes.Add("DATETIME");
            SQLDataTypes.Add("MONEY");

            Connectivity.SQLCalls.QueryResults qc = Connectivity.SQLCalls.sqlQuery("sp_helptext 'dbo.Feeds_Template'");
            string Query = "";
            if (qc.Rows.Count > 0)
            {
                foreach (string row in qc.Rows)
                {
                    Query += row;
                }
            }
            else
            {
                MessageBox.Show("Error: No column codebehind data found.");
                return;
            }

            Query = Query.Split(new string[] { "SELECT DISTINCT" }, StringSplitOptions.None)[1];
            FromStatement = "FROM " + Query.Split(new string[] { "FROM " }, StringSplitOptions.None)[1];
            FromStatement = FromStatement.Replace("|", "");
            FromStatement = FromStatement.Replace("  ", "");
            Query = Query.Split(new string[] { "FROM " }, StringSplitOptions.None)[0];            
            Query = Query.Replace("\r\n", "");
            Query = Query.Replace("|", "");

            List<string> tmpColumnCodeBehind = new List<string>();
            List<int> actualOffsets = new List<int>();
            bool finish = false;
            int lastColumnEnd = 0;
            int lastStartIndex = 0;
            int currentStartIndex = 0;
            int count = 0;
            bool lastLoopSuccess = false;
            //int LengthOfAS = 0;
            while (finish == false)
            {
                if (count==95)
                { }
                if (lastColumnEnd > 0 & lastLoopSuccess == true)
                {
                    lastStartIndex = lastColumnEnd;
                }
                try
                {
                    currentStartIndex = Query.IndexOf(" AS ", lastStartIndex);
                    //LengthOfAS = lastStartIndex - currentStartIndex;
                }
                catch (Exception e)
                {
                    finish = true;
                    break;
                }
                string test = (Query.Substring(currentStartIndex).Split(',')[0]);
                bool checkDataType = false;
                foreach (string dType in SQLDataTypes)
                {
                    if (test.Contains(dType)==true)
                    {
                        checkDataType = true;
                        break;
                    }
                }
                if (checkDataType == false)
                {
                    lastLoopSuccess = true;
                    //find first comma after index, and that is the real end of the line
                    int indexOfEnd = Query.IndexOf(',', currentStartIndex);
                    //string result = Query.Substring(lastStartIndex + 4, (currentStartIndex + 4) - (lastStartIndex + 4));
                    string result = "";
                    if (lastColumnEnd > 0)
                    {
                        try
                        {
                            result = Query.Substring(lastColumnEnd, (indexOfEnd) - (lastColumnEnd));
                            result = result.Substring(0, result.Length - test.Length);
                        }
                        catch { result = "FAILURE"; break; }
                    }
                    else
                    {
                        try
                        {
                            result = Query.Substring(lastStartIndex, (indexOfEnd) - (lastStartIndex));
                            result = result.Substring(0, result.Length - test.Length);
                        }
                        catch { result = "FAILURE."; break; }
                    }
                    lastStartIndex = indexOfEnd;
                    lastColumnEnd = lastStartIndex+1;// +result.Length;
                    ///MessageBox.Show(result);
                    ColumnCodeBehind.Add(result.Trim());
                    count++;
                    if (count >= totColumns-1)
                    {
                        finish = true;
                        break;
                    }
                }
                else
                {
                    lastLoopSuccess = false;
                    lastStartIndex = currentStartIndex+1;
                }
                //finish = true;
                //if (Query.Substring(currentStartIndex + 3))
            }


        }



        private void availableLVI_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ListViewItem tmpItm = new ListViewItem();
            tmpItm.Text = availableLVI.SelectedItems[0].SubItems[1].Text;
            tmpItm.SubItems.Add(availableLVI.SelectedItems[0].SubItems[2].Text);
            tmpItm.SubItems.Add(availableLVI.SelectedItems[0].SubItems[1].Text);
            tmpItm.SubItems.Add(availableLVI.SelectedItems[0].SubItems[3].Text);
            newFeedLVI.Items.Add(tmpItm);
        }

        private void newFeedLVI_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            newFeedLVI.Items.RemoveAt(newFeedLVI.SelectedIndices[0]);
        }







        private int NextCustomID = 0;
        private void customColumnAddBtn_Click(object sender, EventArgs e)
        {
            NextCustomID++;
            string CustomID = "C" + NextCustomID.ToString();
            string ColumnName = customColumnNameTxt.Text;
            string ColumnExample = "N/A (Custom Column)";
            string ColumnCode = customColumnCodeTxt.Text;

            if (ColumnName.Contains(" ")==true & (ColumnName.Contains("[")==false | ColumnName.Contains("]")==false))
            {
                ColumnName = "[" + ColumnName + "]";
            }
            ListViewItem lvi = new ListViewItem();
            lvi.Text = ColumnName;
            lvi.SubItems.Add(ColumnExample);
            lvi.SubItems.Add(CustomID);
            lvi.SubItems.Add(ColumnCode);
            newFeedLVI.Items.Add(lvi);

        }

        private void loadItemDataBtn_Click(object sender, EventArgs e)
        {


            string ItemNumber = itemNumberTxt.Text;
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT * FROM Feeds_Template WHERE id='" + ItemNumber + "'");
            GetColumnCodeBehind(qr.Columns.Count);
            if (qr.Rows.Count > 0 & qr.Columns.Count > 0)
            {
                List<string> cols = new List<string>();
                List<string> rows = new List<string>();
                foreach (string col in qr.Columns)
                {
                    cols.Add(col);
                }
                foreach (string row in qr.Rows)
                {
                    rows.Add(row);
                }

                for (int x = 0; x < cols.Count; x++)
                {
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = x.ToString();
                    string colToAdd = cols[x];
                    if (colToAdd.Contains(" ")==true & (colToAdd.Contains("[")==false | colToAdd.Contains("]")==false))
                    {
                        colToAdd = "[" + colToAdd + "]";
                    }
                    lvi.SubItems.Add(colToAdd);
                    lvi.SubItems.Add(rows[0].Split('|')[x]);


                    try
                    {
                        lvi.SubItems.Add(ColumnCodeBehind[x]);
                    }
                    catch (Exception f)
                    {
                        lvi.SubItems.Add("FAILURE");
                    }

                    availableLVI.Items.Add(lvi);
                }


            }
            else
            {
                MessageBox.Show("This item does not exist.");
            }

        }

        protected virtual void FinalizeCreatedFeed()
        {
            
        }
        private void feedSaveBtn_Click(object sender, EventArgs e)
        {
            string FeedName = "Feeds_";
            if (feedNameTxt.Text.Length > 0 & feedNameTxt.Text.IsNormalized() == true & feedNameTxt.Text.Contains(" ")==false)
            {
                FeedName += feedNameTxt.Text;
            }
            else
            {
                MessageBox.Show("The name of the feed must be greater than 0 characters, must be characters defined in UTF-8, and may not contain spaces.");
                return;
            }
            string CreatedFeed = "CREATE VIEW " + FeedName + " AS\r\nSELECT DISTINCT\r\n";
            foreach (ListViewItem lvi in newFeedLVI.Items)
            {
                string ColumnName = lvi.Text;
                string ColumnCode = lvi.SubItems[3].Text;
                string SQLCode = ColumnCode + " AS " + ColumnName + ",\r\n";
                SQLCode = SQLCode.Replace("  ", "");
                CreatedFeed += SQLCode;
            }
            CreatedFeed = CreatedFeed.Substring(0, CreatedFeed.Length - 3) + " ";
            CreatedFeed += FromStatement;


          

            //now we also need to add to the master feed table this feed and set it to inactive.
            string masterInsert = "INSERT INTO FeedMaster (Feed,Active) VALUES('" + FeedName.Substring(6) + "',0)";


            Connectivity.SQLCalls.sqlQuery(CreatedFeed);
            Connectivity.SQLCalls.sqlQuery(masterInsert);
            System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\woot.txt", CreatedFeed + "\r\n\r\n" + masterInsert);
            
            //now update form1...hmmmm

//            FinalizeCreatedFeed();
            OnFeedCreation(sender, e);
            this.Close();

        }

        
    }
}
