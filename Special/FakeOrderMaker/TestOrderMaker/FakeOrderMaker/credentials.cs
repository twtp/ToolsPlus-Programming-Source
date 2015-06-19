using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FakeOrderMaker
{
    public partial class credentials : Form
    {
        private string Password = "override";

        public credentials()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == Password)
            {
                Form1 loadFrm = new Form1();
                loadFrm.Show();
                this.Visible = false;
            }
            else
            {
                MessageBox.Show("Invalid Password. Please don't try again.");
                textBox1.Text = "";
            }
        }

        private void credentials_Load(object sender, EventArgs e)
        {
            textBox1.KeyPress += new KeyPressEventHandler(textBox1_KeyPress);
        }

        void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                button1_Click(sender, (EventArgs)e);
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }




    }
}
