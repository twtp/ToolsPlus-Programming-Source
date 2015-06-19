using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> cats = new List<string>();
            string file = System.IO.File.ReadAllText(@"c:\users\tomwestbrook\desktop\sphere1cats.txt");
            foreach (string s in file.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                cats.Add(s);
            }
            int counter = 1;
            foreach (string s in cats)
            {
                Connectivity.SQLCalls.sqlQuery("INSERT INTO SphereOneCategories (ID,Name) VALUES(" + counter.ToString() + ",'" + s + "')");
                counter++;
            }

        }
    }
}
