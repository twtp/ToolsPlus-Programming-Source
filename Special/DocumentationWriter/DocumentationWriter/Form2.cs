using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DocumentationWriter
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        public void LoadPreviewDocument(string HTMLIN)
        {
            webBrowser1.DocumentText = HTMLIN;
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
