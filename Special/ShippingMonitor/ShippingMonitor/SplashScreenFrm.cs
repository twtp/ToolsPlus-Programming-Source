using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ShippingMonitor
{
    public partial class SplashScreenFrm : Form
    {

       
        public Form1.FormLayout FormLayoutInfo = new Form1.FormLayout();

        public Screen selectedScreen;
        public bool hasStarted = false;
        public Form1 mainWindow = new Form1();
        public SplashScreenFrm()
        {
            InitializeComponent();
        }

        public void SetProgressLevel(int Progress)
        {
            try
            {
                progressBar1.Value = Progress;
            }
            catch
            {
                if (progressBar1.Maximum < Progress)
                {
                    progressBar1.Value = progressBar1.Maximum;
                }
                else
                {
                    progressBar1.Value = progressBar1.Minimum;
                }
            }
        }
        public int GetProgressLevel()
        {
            return progressBar1.Value;
        }
        public void SetProgressMax(int ProgressMax)
        {
            try
            {
                progressBar1.Maximum = ProgressMax;
            }
            catch
            {
                progressBar1.Maximum = 100;
            }
        }
        public int GetProgressMax()
        {
            return progressBar1.Maximum;
        }
      
        private void SplashScreenFrm_Load(object sender, EventArgs e)
        {
            RefreshDisplay();
        }
        public void LoadMainForm()
        {

            if (hasStarted == false)
            {
                mainWindow.FormLayoutInfo = FormLayoutInfo;
                mainWindow.CurrentMonitor = selectedScreen;
                mainWindow.Show();
                mainWindow.Form1_LoadThisStuff(ref progressBar1);
                hasStarted = true;
                this.Visible = false;
            }
            else
            {
                mainWindow.WindowState = FormWindowState.Normal;
                mainWindow.CurrentMonitor = selectedScreen;
                mainWindow.SetFormToMonitor(selectedScreen);
            }
        }
        public void RefreshDisplay()
        {
            SetFormToMonitor(selectedScreen);
            progressBar1.Left = 0;
            progressBar1.Width = this.Width;
            progressBar1.Height = this.Height / 4;
            progressBar1.Top = (int)(this.Height * .6f);
            splashTmr.Enabled = true;

        }
        public void SetFormToMonitor(Screen TheMonitor)
        {
            Point ScreenCoords = new Point();
            ScreenCoords.X = TheMonitor.WorkingArea.Left;
            ScreenCoords.Y = TheMonitor.WorkingArea.Top;
            this.Left = ScreenCoords.X;
            this.Top = ScreenCoords.Y;
            this.ResizeRedraw = true;
            this.WindowState = FormWindowState.Maximized;
            this.Refresh();
            
        }

        private void splashTmr_Tick(object sender, EventArgs e)
        {
            LoadMainForm();
            splashTmr.Enabled = false;
        }
    }
}
