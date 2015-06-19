using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EBayOffloadReportParser
{
    public partial class Form1 : Form
    {
        string OutputFile = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.Yes | dr == System.Windows.Forms.DialogResult.OK)
            {
                string file = System.IO.File.ReadAllText(openFileDialog1.FileName);
                OutputFile = "";
                foreach (string str in file.Split(new string [] {"\r\n"},StringSplitOptions.RemoveEmptyEntries))
                {
                    if (str.Contains("Invalid category") == true)
                    {
                        //OutputFile += "'" + str.Split(':')[0].Trim() + "',";
                    }
                    else
                    {
                        OutputFile += str + "\r\n";
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                System.IO.File.WriteAllText(saveFileDialog1.FileName, OutputFile);
                MessageBox.Show("All set!");
            }
        }
    }
}
