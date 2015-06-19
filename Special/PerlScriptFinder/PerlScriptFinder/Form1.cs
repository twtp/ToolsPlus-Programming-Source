using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace PerlScriptFinder
{
    public partial class Form1 : Form
    {
        long totScripts = 0;
        long totScriptCalls = 0;

        bool isPaused = false;

        public const string VERSION = "0.9.0";

        private void GETVERSION()
        {
            VersionLbl.Text = "Version: " + VERSION;
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            GETVERSION();
            ClearList();
            textBox1.Text = Application.StartupPath;

            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
    
        }

        void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Environment.Exit(0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text == "Start" | button1.Text == "Continue")
            {
                button1.Text = "Pause";
                if (isPaused == false)
                {
                    ClearList();
                    EnumerateScripts();
                    
                }

                button2.Enabled = false;
                isPaused = false;
                // testListSubItemOffsets();
            }
            else
            {
                button2.Enabled = true;
                button1.Text = "Continue";
                isPaused = true;
            }
        }

        private void testListSubItemOffsets()
        {
            ClearList();
            ListViewItem newScript = new ListViewItem();
            ListViewItem.ListViewSubItem newSubItemCount = new ListViewItem.ListViewSubItem();
            ListViewItem.ListViewSubItem newSubItemDate = new ListViewItem.ListViewSubItem();
            newSubItemCount.Text = "SCRIPT COUNT";
            newSubItemDate.Text = "SCRIPT DATE";
            newScript.Text = "SCRIPT NAME";
            newScript.SubItems.Add(newSubItemCount);
            newScript.SubItems.Add(newSubItemDate);
            listView1.Items.Add(newScript);

        }

        private void ClearList()
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            ColumnHeader scriptName = new ColumnHeader();
            ColumnHeader scriptUsedCount = new ColumnHeader();
            ColumnHeader scriptDateLastUsed = new ColumnHeader();

            scriptName.Width = (int)(listView1.Width * 0.4F);
            scriptUsedCount.Width = (int)(listView1.Width * 0.2F);
            scriptDateLastUsed.Width = (int)(listView1.Width * 0.4F);

            scriptName.Text = "Script Name: ";
            scriptUsedCount.Text = "Logged Count:";
            scriptDateLastUsed.Text = "Date Last Used:";

            scriptUsedCount.Tag = "Numeric";

            listView1.Columns.Add(scriptName);
            listView1.Columns.Add(scriptUsedCount);
            listView1.Columns.Add(scriptDateLastUsed);
        }

        private void EnumerateScripts()
        {
            long logFileCount = Directory.GetFiles(textBox1.Text, "*.log").GetLength(0);
            long y = 1;
            
            

            foreach (string file in Directory.GetFiles(textBox1.Text,"*.log"))
            {
                CheckForPause();
                statusLbl.Text = "File #" + y.ToString() + "/" + logFileCount.ToString();
                statusLbl.Refresh();
                System.Threading.Thread.Sleep(5);
                if (file.Contains(".log") == true)
                {
                    //found log file... lets get it's data and send it to parsing.
                    ParseLogFileData(File.ReadAllLines(file));

                }
                y++;
            }
            button2.Enabled = true;
        }

        private void ParseLogFileData(string[] Data)
        {
            long z = 0;
            long totLines = Data.GetLength(0);


            foreach (string dataStr in Data)
            {
                CheckForPause();
                if (dataStr.Contains(".plex") == true | dataStr.Contains(".pl") == true | dataStr.Contains(".pm") == true)
                {
                    //found a used script...
                    totScriptCalls++;
                    AddScriptToList(dataStr);                    
                }
                status2Lbl.Text = "Line #" + z.ToString() + "/" + totLines.ToString();
                status2Lbl.Refresh();
                System.Threading.Thread.Sleep(5);
                Application.DoEvents();
                z++;
            }
        }


        private void AddScriptToList(string scriptLine)
        {
            string scriptName = scriptLine.Split(' ')[3];
            DateTime scriptUsed = DateTime.Parse(scriptLine.Split(' ')[0] + " " + scriptLine.Split(' ')[1]);
            

            long scriptCount = 0;
            foreach (ListViewItem lvi in listView1.Items)
            {
                CheckForPause();
                if (lvi.Text == scriptName)
                {
                    scriptCount = long.Parse(lvi.SubItems[1].Text) + 1;
                }
            }

            if (scriptCount == 0)
            {
                //new script; just add to list...
                totScripts++;
                ListViewItem newScript = new ListViewItem();
                ListViewItem.ListViewSubItem newSubItemCount = new ListViewItem.ListViewSubItem();
                ListViewItem.ListViewSubItem newSubItemDate = new ListViewItem.ListViewSubItem();
                newSubItemCount.Text = "1";
                newSubItemDate.Text = scriptUsed.ToString();
                newScript.Text = scriptName;
                newScript.SubItems.Add(newSubItemCount);
                newScript.SubItems.Add(newSubItemDate);
                listView1.Items.Add(newScript);
                listView1.Refresh();
                System.Threading.Thread.Sleep(5);
            }
            else
            {
                for (long x = 0; x < listView1.Items.Count; x++)
                {
                    CheckForPause();
                    if (listView1.Items[(int)x].Text == scriptName)
                    {
                        //so add to our running count for this particular script
                        listView1.Items[(int)x].SubItems[1].Text = scriptCount.ToString();

                        //now let's find which date is the latter of the two...
                        long result = DateTime.Compare(DateTime.Parse(listView1.Items[(int)x].SubItems[2].Text), scriptUsed);

                        if (result > 0)
                        {
                            //first date is later than the second
                            //so do nothing
                        }
                        else
                        {
                            //first date is earlier than the second
                            //so change the date to the latest date.
                            listView1.Items[(int)x].SubItems[2].Text = scriptUsed.ToString();
                        }
                        

                    }
                }
                listView1.Refresh();
                System.Threading.Thread.Sleep(5);

            }
            listView1.ListViewItemSorter = new ListViewItemComparer();

            totScriptsFoundLbl.Text = "Unique Scripts Found: " + totScripts.ToString();
            totScriptsUsageLbl.Text = "Total Script Calls:" + totScriptCalls.ToString();


        }



        private void ExportToCSV()
        {
            try
            {
                string tempFile = "Script Name,Script Run Count, Script Last Ran,\r\n";

                foreach (ListViewItem lvi in listView1.Items)
                {
                    tempFile += lvi.Text + "," + lvi.SubItems[1].Text + "," + lvi.SubItems[2].Text + ",\r\n";
                }


                File.WriteAllText(Application.StartupPath + "\\PerlScriptsUsed.csv", tempFile);
                MessageBox.Show("Wrote CSV File to: " + Application.StartupPath + "\\PerlScriptUsed.csv");
            }
            catch (Exception WTF)
            {
                MessageBox.Show("There was an error: " + WTF.Message);
            }



        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExportToCSV();
        }

        private void CheckForPause()
        {
            while (isPaused)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
            
            
        }


    }

    /*public class ListViewItemComparer : System.Collections.IComparer
    {
        private long col;
        public ListViewItemComparer()
        {
            col = 0;
        }
        public ListViewItemComparer(long column)
        {
            col = column;
        }
        public long Compare(object x, object y)
        {           
                
                return String.Compare(((ListViewItem)y).SubItems[col].Text, ((ListViewItem)x).SubItems[col].Text);
            
            
        }
    }

*/

    class ListViewItemComparer : System.Collections.IComparer
    {
        public int Column = 1;
        public System.Windows.Forms.SortOrder Order = SortOrder.Ascending;
        public int Compare(object x, object y) // IComparer Member
        {
            if (!(x is ListViewItem))
                return (0);
            if (!(y is ListViewItem))
                return (0);

            ListViewItem l1 = (ListViewItem)x;
            ListViewItem l2 = (ListViewItem)y;

            if (l1.ListView.Columns[Column].Tag == null)
            {
                l1.ListView.Columns[Column].Tag = "Text";
            }

            if (l1.ListView.Columns[Column].Tag.ToString() == "Numeric")
            {
                float fl1 = float.Parse(l1.SubItems[Column].Text);
                float fl2 = float.Parse(l2.SubItems[Column].Text);

                if (Order == SortOrder.Descending)
                {
                    return fl1.CompareTo(fl2);
                }
                else
                {
                    return fl2.CompareTo(fl1);
                }
            }
            else
            {
                string str1 = l1.SubItems[Column].Text;
                string str2 = l2.SubItems[Column].Text;

                if (Order == SortOrder.Descending)
                {
                    return str1.CompareTo(str2);
                }
                else
                {
                    return str2.CompareTo(str1);
                }
            }
        }
    }

}
