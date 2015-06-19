using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Threading;


namespace TPWillCall
{
    
    public partial class TPWillCall : UserControl
    {
        public TPWillCall()
        {
            InitializeComponent();
            this.SizeChanged += new EventHandler(TPWillCall_SizeChanged);
        }

        void TPWillCall_SizeChanged(object sender, EventArgs e)
        {
            RefreshWillCall();   
        }

        public void RefreshWillCall()
        {
            WillCallHeader.Left = 0;
            //WillCallHeader.Height = this.Height / 3;
            WillCallHeader.Width = this.Width;
            WillCallHeader.Top = 0;

            detailsLbl.Left = 0;
           //detailsLbl.Height = this.Height / 5;
            detailsLbl.Width = this.Width;
            detailsLbl.Top = WillCallHeader.Bottom-15;

            subDetailsLbl.Left = 0;
           // subDetailsLbl.Height = this.Height / 5;
            subDetailsLbl.Width = this.Width;
            subDetailsLbl.Top = detailsLbl.Bottom;
        }
        private void UserControl1_Load(object sender, EventArgs e)
        {
            RefreshWillCall();
            LoadWillCalls();
            if (this.Visible == true)
            {
                strobeTmr.Enabled = true;
            }
            else
            {
                strobeTmr.Enabled = false;
            }
        }
        public void SetBackColor(Color color)
        {
            this.BackColor = color;
            
        }
        public Color GetBackColor()
        {
            return this.BackColor;
        }
        public void SetForeColor(Color color)
        {
            WillCallHeader.ForeColor = color;
            detailsLbl.ForeColor = color;
            subDetailsLbl.ForeColor = color;
        }
        public Color GetForeColor()
        {
            return WillCallHeader.ForeColor;
        }

        public delegate void AsyncMethodCaller();
        public void LoadWillCalls()
        {
            AsyncMethodCaller caller = new AsyncMethodCaller(LoadWillCallsThread);
            caller.BeginInvoke(null, null);
        }



        public void LoadWillCallsThread()
        {
            
            
            //List<string> willCallList = Connectivity.SQLCalls.sqlQuery("SELECT ID,TransactionID FROM WhseTaskWillCallHeader WHERE TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
            List<string> willCallList = ConnectivityWillCall.SQLCalls.sqlQuery("SELECT ID,TransactionID FROM WhseTaskWillCallHeader WHERE TimeInserted > '" + DateTime.Now.ToShortDateString() + "'");
            if (willCallList.Count > 0)
            {
                int totComponentsRequired = 0;
                int totComponentsPicked = 0;
                string willCallListstr = "";
                foreach (string willCall in willCallList)
                {
                    willCallListstr += "#" + willCall.Split('|')[1] + ", "; 
                    List<string> willCallInfo = ConnectivityWillCall.SQLCalls.sqlQuery("SELECT ComponentID,QuantityRequired,QuantityPicked FROM WhseTaskWillCallLines WHERE HeaderID=" + willCall.Split('|')[0]);

                    if (willCallInfo.Count > 0)
                    {
                       
                        foreach (string str in willCallInfo)
                        {
                            totComponentsRequired += int.Parse(str.Split('|')[1]);
                            totComponentsPicked += int.Parse(str.Split('|')[2]);
                        }

                        if (totComponentsPicked < totComponentsRequired)
                        {
                            if (willCallList.Count > 1)
                            {
                                WillCallHeader.Invoke((MethodInvoker)(() => WillCallHeader.Text = "Will Calls (" + willCallList.Count.ToString() + ")"));
                            }
                            else
                            {
                                WillCallHeader.Text = "Will Call";
                            }
                            this.subDetailsLbl.Invoke((MethodInvoker)(() => this.subDetailsLbl.Text = "Qty: " + totComponentsRequired.ToString() + "  Picked: " + totComponentsPicked.ToString()));
                            this.detailsLbl.Invoke((MethodInvoker)(() => this.detailsLbl.Text = willCallListstr.Substring(0, willCallListstr.Length - 2)));
                            this.Invoke((MethodInvoker)(() => this.Visible = true));
                            this.Invoke((MethodInvoker)(() => this.BringToFront()));
                            this.Invoke((MethodInvoker)(() => this.strobeTmr.Enabled = true));
                            //break;
                        }
                        else
                        {
                            this.Invoke((MethodInvoker)(() => this.Visible = false));
                        }

                    }


                }
            }
            else
            {
                this.Invoke((MethodInvoker)(()=>this.Visible = false));
                strobeTmr.Enabled = false;
            }
            if (this.Visible == false)
            {
                strobeTmr.Enabled = false;
            }


        }

        private void strobeTmr_Tick(object sender, EventArgs e)
        {
            if (strobeTmr.Interval == 3000)
            {
                strobeTmr.Interval = 300;
                this.Visible = true;
            }
            else
            {
                if (strobeTmr.Interval == 301)
                {
                    strobeTmr.Interval = 302;
                    this.Visible = true;
                }
                else if (strobeTmr.Interval == 300)
                {
                    strobeTmr.Interval = 301;
                    this.Visible = false;
                }
                else if (strobeTmr.Interval == 302)
                {
                    strobeTmr.Interval = 303;
                    this.Visible = false;
                }
                else
                {
                    strobeTmr.Interval = 3000;
                    this.Visible = true;
                }
            }
        }

        private void subDetailsLbl_Click(object sender, EventArgs e)
        {

        }

        private void updateTmr_Tick(object sender, EventArgs e)
        {
            LoadWillCalls();
            if (this.Visible == true)
            {
                strobeTmr.Enabled = true;
            }
            else
            {
                strobeTmr.Enabled = false;
            }
        }



    }
}
