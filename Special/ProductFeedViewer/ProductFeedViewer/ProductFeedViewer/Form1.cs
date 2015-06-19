using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ProductFeedViewer
{
    public partial class Form1 : Form
    {
        string ToolsPlusFroogleFeed = "\\\\TOOLSPLUS06\\c$\\Inetpub\\wwwroot\\feeds\\toolsplus-froogle.txt";
        string CurrentConCatLineUnmodified = "";
        string CurrentConCatLineModified = "";
        public Form1()
        {
            InitializeComponent();
        }
        //ScrollBar vScrollBar1 = new VScrollBar();
        private void Form1_Load(object sender, EventArgs e)
        {
            button1_Click(sender, e);
            this.Resize += new EventHandler(refreshForm);
            //vScrollBar1.Dock = DockStyle.Right;
            //vScrollBar1.Scroll += new ScrollEventHandler(vScrollBar1_Scroll);
            //panel1.Controls.Add(vScrollBar1);
        }
        public void refreshForm(object sender, EventArgs e)
        {
            feedLbl.Top = 5;
            feedLbl.Left = 5;
            listView1.Top = feedLbl.Bottom + 2;
            listView1.Left = 5;
            listView1.Width = this.ClientSize.Width - 10;
            listView1.Height = (this.ClientSize.Height / 2) -10;

            panel1.Left = 5;
            panel1.Top = listView1.Bottom;
            panel1.Width = this.ClientSize.Width / 2;
            panel1.Height = (this.ClientSize.Height / 2) - 10;

            button3.Top = panel1.Top + (panel1.Height / 2);
            button3.Left = panel1.Right + 5;


            button4.Top = this.ClientSize.Height - button4.Height - 5;
            button4.Left = this.ClientSize.Width - button4.Width - 5;

        }
        void vScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //panel1.VerticalScroll.Value = vScrollBar1.Value;
        }
        private void CopyDataToLVI(ListView listViewData, ListView destinationView)
        {
            if (destinationView.InvokeRequired == true)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    destinationView.Items.Clear();
                    destinationView.Columns.Clear();
                    destinationView.Clear();

                    destinationView.FullRowSelect = listViewData.FullRowSelect;
                    destinationView.GridLines = listViewData.GridLines;
                    destinationView.View = listViewData.View;

                    destinationView.Refresh();

                    foreach (ColumnHeader col in listViewData.Columns)
                    {
                        destinationView.Columns.Add((ColumnHeader)col.Clone());
                    }
                    foreach (ListViewItem lvi in listViewData.Items)
                    {
                        destinationView.Items.Add((ListViewItem)lvi.Clone());
                    }

                    destinationView.Refresh();

                });
            }
            else
            {
                destinationView.Items.Clear();
                destinationView.Columns.Clear();
                destinationView.Clear();

                destinationView.FullRowSelect = listViewData.FullRowSelect;
                destinationView.GridLines = listViewData.GridLines;
                destinationView.View = listViewData.View;

                destinationView.Refresh();

                foreach (ColumnHeader col in listViewData.Columns)
                {
                    destinationView.Columns.Add((ColumnHeader)col.Clone());
                }
                foreach (ListViewItem lvi in listViewData.Items)
                {
                    destinationView.Items.Add((ListViewItem)lvi.Clone());
                }

                destinationView.Refresh();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            listView1.View = View.Details;
            listView1.GridLines = true;
            listView1.FullRowSelect = true;
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.Refresh();

            feedLbl.Text = "Current Loaded Feed: " + ToolsPlusFroogleFeed;
            ListView wtf = ListViewTools.ListViewTools.TSV_to_LVI(ToolsPlusFroogleFeed);
            int columns = wtf.Items[0].SubItems.Count;
            for (int x = 0; x < columns; x++)
            {
                listView1.Columns.Add(x.ToString());
                Label lbl = new Label();
                lbl.Width = 150;
                lbl.Height = 20;
                lbl.Top =(x * lbl.Height) + 2;
                lbl.Left = 10;
                lbl.Text = wtf.Items[0].SubItems[x].Text;
                panel1.Controls.Add(lbl);

                TextBox txt = new TextBox();
                txt.Name = "txt" + x.ToString() + "Txt";
                txt.Parent = panel1;
                txt.Height = 20;
                txt.Width = 200;
                txt.Top = x * txt.Height;
                txt.Left = lbl.Right + 5;
                panel1.Controls.Add(txt);
            }
            CopyDataToLVI(wtf,listView1);


           



        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                CurrentConCatLineUnmodified = "";
                int Index = listView1.SelectedIndices[0];

                ListViewItem lvi = listView1.Items[Index];
                int counter = 0;
                foreach (ListViewItem.ListViewSubItem lvsi in lvi.SubItems)
                {
                    CurrentConCatLineUnmodified += lvsi.Text + "\t";
                    foreach (Control c in panel1.Controls)
                    {
                        if (c.Name == "txt" + counter.ToString() + "Txt")
                        {
                            c.Text = lvsi.Text;
                        }
                    }
                    counter++;
                }
                //CurrentConCatLineUnmodified += "\r\n";
                CurrentConCatLineUnmodified = CurrentConCatLineUnmodified.Substring(0, CurrentConCatLineUnmodified.Length - 3);
               
            }
            catch { }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string HighlightOnReload = "";
            feedLbl.Text = "Saving, please wait...";
            feedLbl.Refresh();
            List<string> columns = new List<string>();
            for (int x = 0; x < panel1.Controls.Count/2; x++)
            {
                if (panel1.Controls["txt" + x.ToString() + "Txt"].Text.Contains("\\t") == true | panel1.Controls["txt" + x.ToString() + "Txt"].Text.Contains("\\r") == true | panel1.Controls["txt" + x.ToString() + "Txt"].Text.Contains("\\n") == true)
                {
                    MessageBox.Show("Input field #" + x.ToString() + " has either a \t, \r, or \n, which cannot be used. Correct this issue before trying to save again.");
                    feedLbl.Text = "Save failed. Invalid input.";
                    feedLbl.Refresh();
                    return;
                }
                columns.Add(panel1.Controls["txt" + x.ToString() + "Txt"].Text);
                if (x == 0)
                {
                    HighlightOnReload = panel1.Controls["txt" + x.ToString() + "Txt"].Text;
                }
            }
            CurrentConCatLineModified = "";
            foreach (string s in columns)
            {
                CurrentConCatLineModified += s + "\t";
            }
            CurrentConCatLineModified = CurrentConCatLineModified.Substring(0, CurrentConCatLineModified.Length - 3);
            

            //CurrentConCatLineModified += "\r\n";
            //string[] file = System.IO.File.ReadAllLines(ToolsPlusFroogleFeed);
            //string line = file[3];
            
           // if (line.Contains("\t") == true)
           // {
           //     MessageBox.Show("OK");

           // }
           // else
           // {
           //     MessageBox.Show("I WILL STRANGLE YOU!!!!");
           // }
           // if (line.Contains(CurrentConCatLineUnmodified) == true)
           // {
           //     MessageBox.Show("YA BOY");
           // }
        
            //System.IO.File.WriteAllText(@"c:\users\tomwestbrook\desktop\wtf.txt", line);
            
            
            string feed = System.IO.File.ReadAllText(ToolsPlusFroogleFeed);
            if (feed.Contains(CurrentConCatLineUnmodified) == true)
            {
                feed = feed.Replace(CurrentConCatLineUnmodified, CurrentConCatLineModified);
                System.IO.File.WriteAllText(ToolsPlusFroogleFeed, feed);
                button1_Click(sender, e);
                int index = 0;
                foreach (ListViewItem lvi in listView1.Items)
                {
                    if (lvi.Text == HighlightOnReload)
                    {
                        index = lvi.Index;
                        break;
                    }
                }
                MessageBox.Show("Saved OK!");
                listView1.Focus();
                listView1.Items[index].Selected = true;
                listView1.Items[index].Focused = true;
                listView1.TopItem = listView1.Items[index];
                feedLbl.Text = "Feed saved.";
                feedLbl.Refresh();                


            }
            else
            {
                MessageBox.Show("Did not find line! Not sure why. Tell Tom.");
            }
            
            

        }

        private void button4_Click(object sender, EventArgs e)
        {

            feedLbl.Text = "Preparing to upload...";
            feedLbl.Refresh();
            DialogResult res = MessageBox.Show("Are you certain you wish to publish this feed? This will send it out!", "Warning!", MessageBoxButtons.YesNo);
            if (res == System.Windows.Forms.DialogResult.Yes | res == System.Windows.Forms.DialogResult.OK)
            {
                feedLbl.Text = "Uploading... Please wait as this could take a minute or two...";
                feedLbl.Refresh();
                //send the file off!
                Connectivity.FTPCalls.FTP_UPLOAD("ftp://uploads.google.com", "toolsplus-froogle.txt", ToolsPlusFroogleFeed, "toolsplus", "summer1", 21);
                feedLbl.Text = "Finished Uploading!";
                feedLbl.Refresh();
            }
            else
            {
                feedLbl.Text = "Did not run upload.";
                feedLbl.Refresh();
            }


        }
    }
}
