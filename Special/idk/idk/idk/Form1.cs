using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace idk
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            for (int x = 0; x < 126; x++)
            {
                List<string> res = Connectivity.SQLCalls.sqlQuery("SELECT ItemNumber FROM PartNumberWebPaths WHERE WebPathID IN (SELECT MasterID FROM MasterCategoryConnectors WHERE ConnectorType=4 AND ConnectorID=" + x.ToString() + ")");
                if (res.Count > 0)
                {
                    listBox1.Items.Add(x.ToString());
                }
            }
        }
    }
}
