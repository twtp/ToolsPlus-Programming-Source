using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AddEBayCatsToPOINV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int count = 0;
            Connectivity.SQLCalls.QueryResults qr = Connectivity.SQLCalls.sqlQuery("SELECT CategoryName,CategoryID FROM EBayHomeGardenConversionID WHERE BusinessIndustrialID IS NOT NULL AND CategoryID<>122929");
            if (qr.Rows.Count >0)
            {
                foreach (string row in qr.Rows)
                {
                    count++;
                    string catName = row.Split('|')[0];
                    string catID = row.Split('|')[1];
                    Connectivity.SQLCalls.sqlQuery("INSERT INTO EBayCategory (CategoryID,CategoryType,CategoryName,ConditionNewID,ConditionReconID) VALUES(" + catID + ",0,'" + catName + "',1000,2000)");
                    this.Text = "Working on " + count.ToString() + "/" + qr.Rows.Count.ToString();
                    this.Refresh();
                }
            }
                
        }
    }
}
