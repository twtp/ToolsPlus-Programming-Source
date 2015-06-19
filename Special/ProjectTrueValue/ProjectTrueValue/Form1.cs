using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectTrueValue
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            debugFrm df = new debugFrm();
            df.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tvFrm tv = new tvFrm();
            tv.Show();
        }
    }
}
