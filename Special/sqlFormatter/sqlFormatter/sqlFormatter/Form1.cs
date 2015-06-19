using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace sqlFormatter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sqlLine = textBox3.Text;
            string sqlListLine = "";
            List<string> listLines = new List<string>();
            foreach (string str in textBox1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None))
            {
                listLines.Add("'" + str + "'");
            }
            foreach (string s in listLines)
            {
                sqlListLine += s + ",";
            }
            sqlListLine = sqlListLine.Substring(0, sqlListLine.Length - 1);
            sqlLine = sqlLine.Replace("{LIST}", sqlListLine);
            textBox2.Text = sqlLine;

            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Connectivity.SQLCalls.sqlQuery(textBox2.Text);
        }
    }
}
