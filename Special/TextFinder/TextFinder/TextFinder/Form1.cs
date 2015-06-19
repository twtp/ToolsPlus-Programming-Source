using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TextFinder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string appPath = Application.StartupPath;
            foreach (string file in System.IO.Directory.GetFiles(appPath))
            {
                if (file.Contains(".log") == true)
                {
                    string File = System.IO.File.ReadAllText(file);
                    if (File.Contains(textBox1.Text) == true)
                    {
                        MessageBox.Show("Exists in file \"" + file + "\"");
                    }
                }
            }
        }
    }
}
