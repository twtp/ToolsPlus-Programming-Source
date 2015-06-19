using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ShippingMonitor2
{
    public partial class Form2 : Form
    {

       


        public bool hasStarted = false;
        public Screen SelectedMonitor;
        public List<Screen> MonitorList;
        public Form1 tvMain = new Form1();
       

        public List<Screen> GetMonitors()
        {
            List<Screen> ScreenList = new List<Screen>();
            try
            {
                foreach (Screen screen in Screen.AllScreens)
                {
                    ScreenList.Add(screen);
                }
            }
            catch
            {
            }

            return ScreenList;
        }
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            this.FormClosing += new FormClosingEventHandler(Form2_FormClosing);
            MonitorList = GetMonitors();
            listBox1.Items.Clear();
            foreach (Screen scr in MonitorList)
            {
                listBox1.Items.Add(scr.DeviceName.ToString());
            }
        }

        void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                SelectedMonitor = MonitorList[listBox1.SelectedIndex];
                //mainWindow.CurrentMonitor = SelectedMonitor;
                tvMain.selectedScreen = SelectedMonitor;
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tvMain.Show();


        }

        private void amEbayOrdersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (amEbayOrdersChk.Checked == true)
            {
                tvMain.ebayAMTicker.Visible = true;
            }
            else
            {
                tvMain.ebayAMTicker.Visible = false;
            }
        }

        private void pmEbayOrdersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (pmEbayOrdersChk.Checked == true)
            {
                tvMain.ebayPMTicker.Visible = true;
            }
            else
            {
                tvMain.ebayPMTicker.Visible = false;
            }
        }

        private void yahooOrdersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (yahooOrdersChk.Checked == true)
            {
                tvMain.yahooTicker.Visible = true;
            }
            else
            {
                tvMain.yahooTicker.Visible = false;
            }
        }

        private void otherOrdersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (otherOrdersChk.Checked == true)
            {
                tvMain.otherOrdersTicker.Visible = true;
            }
            else
            {
                tvMain.otherOrdersTicker.Visible = false;
            }
        }

        private void funFactsChk_CheckedChanged(object sender, EventArgs e)
        {
            if (funFactsChk.Checked == true)
            {
                tvMain.funFactsTicker.Visible = true;
                tvMain.testTicker.Visible = true;
            }
            else
            {
                tvMain.funFactsTicker.Visible = false;
                tvMain.testTicker.Visible = false;
            }
        }

        private void showPackersLVIChk_CheckedChanged(object sender, EventArgs e)
        {
            if (showPackersLVIChk.Checked == true)
            {
                tvMain.packersLVI.Visible = true;
            }
            else
            {
                tvMain.packersLVI.Visible = false;
            }
        }

        private void toteInformationChk_CheckedChanged(object sender, EventArgs e)
        {
            if (toteInformationChk.Checked == true)
            {
                tvMain.totesInfoTicker.Visible = true;
            }
            else
            {
                tvMain.totesInfoTicker.Visible = false;
            }
        }

        private void willCallChk_CheckedChanged(object sender, EventArgs e)
        {
            if (willCallChk.Checked == true)
            {
                tvMain.tpWillCall1.Visible = true;
            }
            else
            {
                tvMain.tpWillCall1.Visible = false;
            }
        }

        private void poChk_CheckedChanged(object sender, EventArgs e)
        {
            if (poChk.Checked == true)
            {
                tvMain.tppoReceivedAlert1.Visible = true;
            }
            else
            {
                tvMain.tppoReceivedAlert1.Visible = false;
            }
        }

        private void transfersChk_CheckedChanged(object sender, EventArgs e)
        {
            if (transfersChk.Checked == true)
            {
                tvMain.tpTransferAlert1.Visible = true;
            }
            else
            {
                tvMain.tpTransferAlert1.Visible = false;
            }
        }

        private void restockingChk_CheckedChanged(object sender, EventArgs e)
        {
            tvMain.tpRestockingAlert1.Visible = restockingChk.Checked;
        }

        private void titleChk_CheckedChanged(object sender, EventArgs e)
        {
            tvMain.TopCaptionLbl.Visible = titleChk.Checked;
        }

        private void timeChk_CheckedChanged(object sender, EventArgs e)
        {
            tvMain.tpClock1.Visible = timeChk.Checked;
        }



    }
}
