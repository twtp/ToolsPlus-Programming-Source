using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TPClock
{
    public partial class TPClock : UserControl
    {
        public TPClock()
        {
            InitializeComponent();

            timeLbl.Parent = this;
            controlResize();
            this.SizeChanged += new EventHandler(UserControl1_SizeChanged);
            
        }

        void UserControl1_SizeChanged(object sender, EventArgs e)
        {
            controlResize();
        }

        private void timeLbl_Click(object sender, EventArgs e)
        {
            
        }
        private void controlResize()
        {
            timeLbl.Left = 0;
            timeLbl.Width = this.Width;
            timeLbl.Height = this.Height;
            timeLbl.TextAlign = ContentAlignment.MiddleCenter;
        }
        public void SetBackgroundColor(Color backgroundColor)
        {
            this.BackColor = backgroundColor;

        }
        private void timeTmr_Tick(object sender, EventArgs e)
        {
            timeLbl.Text = DateTime.Now.ToString("h:mm:ss tt");
        }


   
    }
}
