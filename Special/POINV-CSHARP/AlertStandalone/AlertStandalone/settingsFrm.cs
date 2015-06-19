using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace AlertStandalone
{
    public partial class settingsFrm : Form
    {
        public settingsFrm()
        {
            InitializeComponent();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void settingsFrm_Load(object sender, EventArgs e)
        {
            menuStrip1.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            menuStrip1.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            leftWrench.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            leftWrench.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            rightWrench.MouseMove += new MouseEventHandler(menuStrip1_MouseMove);
            rightWrench.MouseDown += new MouseEventHandler(menuStrip1_MouseDown);
            titleLbl.MouseMove +=new MouseEventHandler(menuStrip1_MouseMove);
            titleLbl.MouseDown +=new MouseEventHandler(menuStrip1_MouseDown);
        }

        void menuStrip1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        void menuStrip1_MouseMove(object sender, MouseEventArgs e)
        {
            if (!Focused)
            {
                Focus();
            }
        }



        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd,
                         int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        private static extern bool ReleaseCapture();

    }
}
