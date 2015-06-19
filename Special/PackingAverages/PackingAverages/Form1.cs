using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PackingAverages
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string todaysDate = DateTime.Now.ToString();
        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ShortName FROM Users WHERE Deleted=0 AND SpecialUser=0");
            comboBox1.Items.Clear();
            foreach (string row in results)
            {
                comboBox1.Items.Add(row.Split('|')[0]);
            }


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string user = comboBox1.Items[comboBox1.SelectedIndex].ToString();
            todaysDate = DateTime.Now.ToString();
            todaysDate = todaysDate.Split(' ')[0];

            string query = "SELECT COUNT(EndTime) AS TotalPackagesPacked,AVG(DATEDIFF(ss,StartTime,EndTime)) AS AveragePackingTime FROM totePackingTime INNER JOIN Users ON Users.ID=totePackingTime.UserID WHERE StartTime > '" + todaysDate + "' AND Users.ShortName='" + user + "'";
            List<string> results = Connectivity.SQLCalls.sqlQuery(query);
            if (results.Count > 0)
            {
                packedLbl.Text = "Packed: " + results[0].Split('|')[0];
                timeLbl.Text = "AVG Time: " + results[0].Split('|')[1];
            }
            else
            {
                packedLbl.Text = "Packed: N/A";
                timeLbl.Text = "AVG Time: N/A";
                MessageBox.Show("no data returned");
            }
        }
    }
}
