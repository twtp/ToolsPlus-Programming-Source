using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FeedRunner
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listFeeds();
        }
        private void listFeeds()
        {
            listBox1.Items.Clear();
            foreach (string file in System.IO.Directory.GetFiles(@"\\toolsplus04\perl\site\lib\toolsplus\feeds\"))
            {
                listBox1.Items.Add(file.Split('\\')[file.Split('\\').GetLength(0)-1]);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string startFeedLine = "/C \\\\toolsplus04\\perl\\bin\\perl -MToolsPlus::Feeds -e \"create_feeds(sync_images=>" + ((checkBox1.Checked) ? "1":"0") + ",send_feeds=>" + ((checkBox2.Checked) ? "1":"0") + ", engines=>[qw/" + listBox1.SelectedItem.ToString().Split('.')[0] + "/]);\"";
            //System.IO.File.WriteAllText("c:\\users\\tomwestbrook\\desktop\\testin.txt", startFeedLine);
            System.Diagnostics.Process.Start("cmd.exe", startFeedLine);
        }
    }
}
