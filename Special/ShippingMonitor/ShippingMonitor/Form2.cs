using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ShippingMonitor
{
    public partial class Form2 : Form
    {

        public Form1.FormLayout FormLayoutInfo = new Form1.FormLayout();


        public bool hasStarted = false;
        public Screen SelectedMonitor;
        public List<Screen> MonitorList;
       
        public SplashScreenFrm loadingFrm = new SplashScreenFrm();

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
                loadingFrm.selectedScreen = SelectedMonitor;
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormLayoutInfo = new Form1.FormLayout();
            FormLayoutInfo.ShowAMEBayOrders = this.amEbayOrdersChk.Checked;
            FormLayoutInfo.ShowPMEBayOrders = this.pmEbayOrdersChk.Checked;
            FormLayoutInfo.ShowYahooOrders = this.yahooOrdersChk.Checked;
            FormLayoutInfo.ShowOtherOrders = this.otherOrdersChk.Checked;
            FormLayoutInfo.ShowFunFacts = this.funFactsChk.Checked;
            FormLayoutInfo.ShowPOs = this.poChk.Checked;
            FormLayoutInfo.ShowToteStatus = this.toteInformationChk.Checked;
            FormLayoutInfo.ShowTransfers = this.transfersChk.Checked;
            FormLayoutInfo.ShowWillCalls = this.willCallChk.Checked;

            loadingFrm.FormLayoutInfo = FormLayoutInfo;
            
            loadingFrm.Show();
            loadingFrm.selectedScreen = SelectedMonitor;
            loadingFrm.SetFormToMonitor(SelectedMonitor);
            loadingFrm.SetProgressMax(100);
            loadingFrm.SetProgressLevel(0);            
            loadingFrm.BringToFront();


        }
    }
}
