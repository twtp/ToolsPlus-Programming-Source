using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace tickerTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
       
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.SizeChanged += new EventHandler(Form1_SizeChanged);
            
            tpTicker1.SetHeaderBackgroundImage(Properties.Resources.tickerHeaderBackground2);
            //tpTicker1.SetHeaderBackgroundColor(Color.Turquoise);
            tpTicker1.SetHeaderTextForeColor(Color.Black);
            tpTicker1.showSubHeader = true;
            tpTicker1.Top = 50;
            tpTicker1.Left = 25;
            tpTicker1.Width = this.Width-50;
            tpTicker1.Height = this.Height-100;
            tpTicker1.SetHeaderText("Testin'\r\n");
            tpTicker1.SetHeaderSubText("subtitle");
            tpTicker1.SetTickerText( "Test");
            
            tpTicker1.SetTickerTextBackgroundColor(Color.Orange);
            tpTicker1.SetTickerTextBackgroundImage(Properties.Resources.tickerBackgroundMiddle);
            tpTicker1.SetTickerEndImage(Properties.Resources.tickerBackgroundEnd);
            //tpTicker1.SetTickerTextBackgroundImage
        }

        void Form1_SizeChanged(object sender, EventArgs e)
        {
            
            lock (tpTicker1)
            {
                tpTicker1.DisableTicker();
                tpTicker1.Top = 50;
                tpTicker1.Left = 25;
                tpTicker1.Width = this.Width-50;
                tpTicker1.Height = this.Height - 100;
                tpTicker1.Refresh();
                tpTicker1.EnableTicker();
                
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tpTicker1.SetTickerMethod((Func<TPTickerCtl.TPTicker.TickerData>)tickerMethod);
            //tpTicker1.EnableUpdates(10);
            
        }
        private TPTickerCtl.TPTicker.TickerData tickerMethod()
        {
            TPTickerCtl.TPTicker.TickerData TD = new TPTickerCtl.TPTicker.TickerData();

            List<string> userNames = Connectivity.SQLCalls.sqlQuery("SELECT ShortName From Users WHERE Deleted=0");
            TD.HeaderText = "Users";
            TD.HeaderSubText = "Tot. Users: " + userNames.Count.ToString();
            TD.TickerText = "Listing all users @ " + DateTime.Now.ToString() + "....       ";
            string userList = "";
            foreach (string user in userNames)
            {
                userList += user.Split('|')[0] + ", ";
            }
            userList = userList.Substring(0, userList.Length - 1);
            TD.TickerText += userList;
            TD.ShowSubHeader = true;
            TD.SetTickerBackgroundColor = Color.Orange;
            TD.Enabled = true;
            TD.TickerSpeed = 15;
            TD.UpdateSpeedInSeconds = 10;
            return TD;
        }

    }
}
