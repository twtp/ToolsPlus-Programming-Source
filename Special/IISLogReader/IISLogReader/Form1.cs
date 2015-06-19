using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IISLogReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] files = System.IO.Directory.GetFiles("\\\\10.0.50.16\\c$\\windows\\system32\\LogFiles\\W3SVC1\\");

            foreach (string file in files)
            {
                if (file.Contains(".log")==true)
                {
                    string data = System.IO.File.ReadAllText(file);
                    if (data.Contains(textBox1.Text)==true)
                    {
                        MessageBox.Show("Found in file" + file);
                    }

                }
            }

        }
    }
}
