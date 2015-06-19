using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FeedDifferences
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            openFileDialog1.ShowDialog();
            string path = openFileDialog1.FileName;
            string feed1 = System.IO.File.ReadAllText(path);
            string[] feed = feed1.Split(new string[]{"\r\n"},StringSplitOptions.None);

            foreach (string s in feed)
            {
                listBox1.Items.Add(s.Split('\t')[0]);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            openFileDialog1.ShowDialog();
            string path = openFileDialog1.FileName;
            string feed2 = System.IO.File.ReadAllText(path);
            string[] feed = feed2.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            foreach (string s in feed)
            {
                listBox2.Items.Add(s.Split('\t')[0]);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox3.Items.Clear();
            foreach (string itm in listBox1.Items)
            {
                bool found = false;
                foreach (string itm2 in listBox2.Items)
                {
                    if (itm == itm2)
                    {
                        found = true;
                        break;
                    }
                }
                if (found == false)
                {
                    listBox3.Items.Add(itm);
                }
            }
        }
    }
}
