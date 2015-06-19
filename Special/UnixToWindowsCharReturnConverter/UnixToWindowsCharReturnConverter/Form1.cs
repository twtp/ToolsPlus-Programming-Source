using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace UnixToWindowsCharReturnConverter
{
    public partial class Form1 : Form
    {
        public string origFile = "";
        public string newFile = "";


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                origFile = openFileDialog1.FileName;


            }
            fileLbl.Text = origFile;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = saveFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK | dr == System.Windows.Forms.DialogResult.Yes)
            {
                newFile = saveFileDialog1.FileName;


            }
            label1.Text = newFile;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                string tmpFile = System.IO.File.ReadAllText(origFile);
                tmpFile = tmpFile.Replace("\r\n", "\n").Replace("\n", "\r\n");
                System.IO.File.WriteAllText(newFile, tmpFile);
                MessageBox.Show("done");
            }
            if (radioButton2.Checked == true)
            {
                string tmpFile = System.IO.File.ReadAllText(origFile);
                tmpFile = tmpFile.Replace("\r\n", "\n");
                System.IO.File.WriteAllText(newFile, tmpFile);
                MessageBox.Show("done");
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
