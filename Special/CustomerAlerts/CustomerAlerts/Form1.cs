using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CustomerAlerts
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            listView1.Columns.Clear();
            listView1.Clear();
            listView1.View = View.Details;
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            ColumnHeader id = new ColumnHeader();
            ColumnHeader alertType = new ColumnHeader();
            ColumnHeader email = new ColumnHeader();
            ColumnHeader threshold = new ColumnHeader();
            ColumnHeader sentNotification = new ColumnHeader();
            ColumnHeader requestTimestamp = new ColumnHeader();
            ColumnHeader sentTimestamp = new ColumnHeader();

            id.Width = 1;
            id.Text = "ID";
            alertType.Width = 50;
            alertType.Text = "Alert";
            email.Width = 150;
            email.Text = "Email Address";
            threshold.Width = 60;
            threshold.Text = "Threshold";
            sentNotification.Width = 45;
            sentNotification.Text = "Sent";
            requestTimestamp.Width = 150;
            requestTimestamp.Text = "Request Time";
            sentTimestamp.Width = 150;
            sentTimestamp.Text = "Sent Time";

            listView1.Columns.Add(id);
            listView1.Columns.Add(alertType);
            listView1.Columns.Add(email);
            listView1.Columns.Add(threshold);
            listView1.Columns.Add(sentNotification);
            listView1.Columns.Add(requestTimestamp);
            listView1.Columns.Add(sentTimestamp);

            listView1.Refresh();
            List<string> results = Connectivity.SQLCalls.sqlQuery("SELECT ID,AlertType,EmailAddress,Threshold,SentNotification,RequestTimestamp,SentNotificationTimestamp FROM CustomerAlerts WHERE ItemNumber='" + textBox1.Text + "'");
            if (results.Count > 0)
            {
                foreach (string res in results)
                {
                    string sid = res.Split('|')[0];
                    string salertType = res.Split('|')[1];
                    string semail = res.Split('|')[2];
                    string sthreshold = res.Split('|')[3];
                    string ssentNotification = res.Split('|')[4];
                    string sreqTime = res.Split('|')[5];
                    string ssentTime = res.Split('|')[6];
                    ListViewItem lvi = new ListViewItem();
                    lvi.Text = sid;
                    lvi.SubItems.AddRange(new string[] { salertType, semail, sthreshold, ssentNotification, sreqTime, ssentTime });
                    listView1.Items.Add(lvi);
                }
            }

        }
    }
}
