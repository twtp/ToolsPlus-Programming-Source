using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Threading;

/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
//     Tools Plus TV Ticker Version 1.0                        //
//            created by: Tom Westbrook                        //
//                          10/31/2014                         //
/////////////////////////////////////////////////////////////////
//                                                             //
//  Two ways to use the TPTicker. First the manual way.        //
//  Call any one of these listed methods to change the ticker. //
//                                                             //
//  SetTickerSpeed / GetTickerSpeed -                          //
//      These methods Set/Get speed of ticker bar scrolling    //
//  SetHeaderText  / GetHeaderText -                           //
//      These Set/Get the text displayed on the left of ticker.//
//  SetHeaderSubText / GetHeaderSubText -                      //
//      These Set/Get the text displayed below the header. If  //
//      this field is Set, the sub header is displayed auto.   //
//  SetTickerText / GetTickerText -                            //
//      These Set/Get the text displayed in the ticker itself. //
//      If ticker text is longer than ticker, it auto-scrolls. //
//  SetTickerEndImage / GetTickerEndImage                      //
//      These Set/Get the image used to terminate the ticker.  //
//  SetTickerEndColor / GetTickerEndColor                      //
//      These Set/Get the color to terminate the ticker with.  //
//  SetTickerTextBackgroundColor / GetTickerTextBackgroundColor//
//      These Set/Get the color behind the ticker text.        //
//  SetTickerTextBackgroundImage / GetTickerTextBackgroundImage//
//      These Set/Get the image that encapsulates the ticker's //
//      mid section (as in around the ticker text itself).     //
//  SetHeaderBackgroundImage / GetHeaderBackgroundImage -      //
//      These Set/Get the image of the header of the ticker.   //
//  SetHeaderBackgroundColor / GetHeaderBackgroundColor -      //
//      These Set/Get the color of the ticker's header.        //
//  EnableSubHeader / DisableSubHeader -                       //
//      Enables/Disables the viewing of the sub-header.        //
//  EnableTicker / DisableTicker -                             //
//      Enables/Disables the ticker.                           //
/////////////////////////////////////////////////////////////////
//  The second way to use the TPTicker is to create a method in//
//  the calling program using the type                         //
//  'TPTickerCtl.TPTicker.TickerData'. Create a new TickerData //
//  variable and populate it's info using the included         //
//  'Connectivity.SQLCalls.sqlQuery()' method to aquire data.  //
//  Once the method is written to populate as many variables   //
//  in the TickerData type as possible, then have another      //
//  method or trigger that sends this method to the            //
//  SetTickerMethod() method in the TPTicker.                  //



namespace TPTickerCtl
{
    public delegate void AsyncMethodCaller();

    public partial class TPTicker : UserControl
    {
        private const int headerPadding = 15;
        private const int headerDefaultWidthDivisor = 3;
        private const int tickerWidthBuffer = 75;
        private const int tickerShimAmount = 1;
        private const int tickerBackgroundOverset = 1;
        private const int headerSubVerticalInset = 50;

        public int updateTimerInSeconds;
        private int updateTimerSecondsPassed;
        private Func<TickerData> UpdateMethod;
        public bool showSubHeader = true;
        public bool isUpdateEnabled = false;
        public bool isEnabled = false;

        public TPTicker()
        {
            InitializeComponent();
            
        }
        public struct TickerData
        {
            public string TickerText;
            public string HeaderText;
            public string HeaderSubText;
            public Image SetHeaderBackgroundImage;
            public Color SetHeaderBackgroundColor;
            public Color SetHeaderTextForeColor;
            public Image SetTickerBackgroundImage;
            public Color SetTickerBackgroundColor;
            public Color SetTickerForegroundColor;
            public Image SetEndTickerImage;
            public Color SetEndTickerColor;
            public int TickerSpeed;
            public int UpdateSpeedInSeconds;
            public bool Enabled;
            public bool ShowSubHeader;

        }
        public void processTickerData(TickerData data)
        {
            if (tickerTextLbl.InvokeRequired == true)
            {
                tickerTextLbl.Invoke((MethodInvoker)(() => tickerTextLbl.Text = ""));
                headerTextLbl.Invoke((MethodInvoker)(() => headerTextLbl.Text = ""));
            }
            else
            {
                tickerTextLbl.Text = "";
                headerTextLbl.Text = "";
            }
            if (data.TickerText != "")
            {
                SetTickerText(data.TickerText);
            }
            if (data.SetTickerForegroundColor != null)
            {
                SetTickerTextForegroundColor(data.SetTickerForegroundColor);
            }
            if (data.HeaderText != "")
            {
                SetHeaderText(data.HeaderText);
            }
            if (data.HeaderSubText != "")
            {
                SetHeaderSubText(data.HeaderSubText);
            }
            if (data.SetHeaderBackgroundImage != null)
            {
                SetHeaderBackgroundImage(data.SetHeaderBackgroundImage);
            }
            if (data.SetHeaderBackgroundColor != null)
            {
                SetHeaderBackgroundColor(data.SetHeaderBackgroundColor);
            }
            if (data.SetHeaderTextForeColor != null)
            {
                SetHeaderBackgroundColor(data.SetHeaderTextForeColor);
            }
            if (data.SetTickerBackgroundImage != null)
            {
                SetTickerTextBackgroundImage(data.SetTickerBackgroundImage);
            }
            if (data.SetTickerBackgroundColor != null)
            {
                SetTickerTextBackgroundColor(data.SetTickerBackgroundColor);
            }
            if (data.TickerSpeed > 0)
            {
                Invoke((MethodInvoker)(()=>SetTickerSpeed(data.TickerSpeed)));
            }
            if (data.SetEndTickerImage != null)
            {
                SetTickerEndImage(data.SetEndTickerImage);
            }
            if (data.SetEndTickerColor != null)
            {
                SetTickerEndColor(data.SetEndTickerColor);
            }
           // if (data.Enabled == true)
           // {
           //     Invoke((MethodInvoker)(() =>EnableTicker()));
           // }else{ Invoke((MethodInvoker)(() =>DisableTicker()));}
            if (data.Enabled == true)
            {
                Invoke((MethodInvoker)(()=>tickerTextLbl_TextChanged(null,null)));
            }
            if (data.ShowSubHeader == true)
            {
                Invoke((MethodInvoker)(() => EnableSubHeader()));
            }
            else{Invoke((MethodInvoker)(()=>DisableSubHeader()));}
            if (data.UpdateSpeedInSeconds > 0)
            {
                Invoke((MethodInvoker)(()=> EnableUpdates(data.UpdateSpeedInSeconds)));
                
            }
            else
            {
                updateTimerInSeconds = 0;
                updateTimerSecondsPassed = 0;
                isUpdateEnabled = false;               
                updateTmr.Enabled = false;
            }


        }

        public void ThreadedProcess(object inputMethod)
        {
            Func<TickerData> threadedFunction = (Func<TickerData>)inputMethod;
            processTickerData(threadedFunction());
                     
        }

        public void SetTickerMethod(Func<TickerData> inputMethod)
        {
            UpdateMethod = inputMethod;
            //ParameterizedThreadStart pts = new ParameterizedThreadStart(ThreadedProcess);
            //Thread processThread = new Thread(pts);
            //processThread.Start(inputMethod);
            launchMethodEvent();
        }

        private void TPTicker_Load(object sender, EventArgs e)
        {
            SetPreDefaults();
            SetDefaults();
        }



        private void SetPreDefaults()
        {
            SetStyle(System.Windows.Forms.ControlStyles.SupportsTransparentBackColor, true);
            this.BackColor = Color.Transparent;
            headerTextLbl.Parent = headerBackgroundPic;
            tickerBackDropColorLbl.Parent = tickerBackgroundPic;
            //tickerTextLbl.Parent = tickerBackgroundPic;
            tickerTextLbl.Parent = tickerBackDropColorLbl;
            //headerSubLbl.Parent = headerBackgroundPic;
            headerSubLbl.Parent = headerTextLbl;
            tickerEndPic.Parent = this;
            headerTextLbl.TextAlign = ContentAlignment.MiddleCenter;
            this.SizeChanged += new EventHandler(TPTicker_SizeChanged);
            tickerTextLbl.TextChanged += new EventHandler(tickerTextLbl_TextChanged);
           

        }
        private void SetDefaults()
        {
            //label1.Visible = false;
            //tickerTextLbl.Text = "";
            //headerTextLbl.Text = "";

           
            //headerBackgroundPic.BackColor = Color.Transparent;
            //headerTextLbl.ForeColor = Color.White;
            //tickerTextLbl.ForeColor = Color.Black;
            //tickerTextLbl.BackColor = Color.AliceBlue;
            //this.BackColor = tickerTextLbl.BackColor;
            headerBackgroundPic.Top = 0;
            headerBackgroundPic.Left = 0;
            headerBackgroundPic.Height = this.Height;
            headerBackgroundPic.Width = this.Width / headerDefaultWidthDivisor;

            if (showSubHeader == true)
            {
                headerTextLbl.Top = 0;
                headerTextLbl.Left = 0;
                headerTextLbl.Height = (int)(this.Height);
                headerTextLbl.Width = headerBackgroundPic.Width - headerPadding;

                headerSubLbl.Visible = true;
                headerSubLbl.Top = headerTextLbl.Bottom - headerSubVerticalInset;
                headerSubLbl.Height = (int)(this.Height);
                headerSubLbl.Width =  headerTextLbl.Width;
                headerSubLbl.Left = 0;
                headerSubLbl.BringToFront();
                
            }
            else
            {

                headerTextLbl.Top = 0;
                headerTextLbl.Left = 0;
                headerTextLbl.Height = this.Height;
                headerTextLbl.Width = headerBackgroundPic.Width - headerPadding;
                headerSubLbl.Visible = false;
                headerSubLbl.SendToBack();
            }
            
            
            tickerTextLbl.Top = 0;
            tickerTextLbl.Height = this.Height-(tickerShimAmount*2);
            tickerTextLbl.Left = headerBackgroundPic.Right;            
            
          
            tickerBackgroundPic.Left = headerBackgroundPic.Right - tickerBackgroundOverset;
            tickerBackgroundPic.Top = 0;
            tickerBackgroundPic.Width = this.Width - tickerBackgroundPic.Left;
            tickerBackgroundPic.Height = this.Height;
            tickerBackDropColorLbl.Text = "";
            //tickerBackDropColorLbl.BackColor = ;
            tickerBackDropColorLbl.Left = 0;
            tickerBackDropColorLbl.Top = tickerShimAmount;
            tickerBackDropColorLbl.Height = this.Height - (tickerShimAmount * 2);
            tickerBackDropColorLbl.Width = this.Width - tickerBackDropColorLbl.Left;
            tickerBackDropColorLbl.SendToBack();
           
            tickerBackgroundPic.SendToBack();
            tickerTextLbl.BringToFront();
            
             // tickerBackgroundPic;            
            tickerEndPic.Height = this.Height;
            tickerEndPic.Width = tickerEndPic.Height / 3;
            tickerEndPic.Left = this.Right - tickerEndPic.Width;
            tickerEndPic.Top = 0;
            
            
        }

        void TPTicker_SizeChanged(object sender, EventArgs e)
        {
            SetDefaults();
        }

        void tickerTextLbl_TextChanged(object sender, EventArgs e)
        {
            
            int textSize = System.Windows.Forms.TextRenderer.MeasureText(tickerTextLbl.Text, new Font(tickerTextLbl.Font.FontFamily, tickerTextLbl.Font.Size, tickerTextLbl.Font.Style)).Width;
           
            Invoke((MethodInvoker)(()=>tickerTextLbl.Width = textSize + tickerWidthBuffer));

            if (tickerTextLbl.Width > tickerBackDropColorLbl.Width)
            {
                //MessageBox.Show("Appearantly tickerTextLbl.Width = " + tickerTextLbl.Width.ToString() + " is greater than tickerBackDropColorLbl.Width = " + tickerBackDropColorLbl.Width.ToString());
                //isEnabled = true;
                Invoke((MethodInvoker)(() => EnableTicker()));
                
                    //Invoke((MethodInvoker)(()=>tickerTmr.Enabled = true));
                
            }
            else
            {
                Invoke((MethodInvoker)(()=>tickerTmr.Enabled = false));
                Invoke((MethodInvoker)(()=>tickerTextLbl.Left = headerPadding));
                Invoke((MethodInvoker)(()=>tickerTextLbl.Width = this.Width - tickerTextLbl.Left));
            }
        }
        public void SetTickerSpeed(int tickerSpeed)
        {
            
            tickerTmr.Interval = tickerSpeed;
        }
        public int GetTickerSpeed()
        {
            return tickerTmr.Interval;
        }



        public void SetHeaderText(string headerText)
        {
            if (headerTextLbl.InvokeRequired == true)
            {
                headerTextLbl.Invoke((MethodInvoker)(() => headerTextLbl.Text = headerText));
                headerBackgroundPic.Invoke((MethodInvoker)(() => headerBackgroundPic.Width = headerTextLbl.Width + headerPadding));
            }
            else
            {
                headerTextLbl.Text = headerText;
                headerBackgroundPic.Width = headerTextLbl.Width + headerPadding;
            }
        }
        public string GetHeaderText()
        {
            return headerTextLbl.Text;
        }



        public void SetHeaderSubText(string tickerSubText)
        {
            if (headerSubLbl.InvokeRequired == true)
            {
                EnableSubHeader();
                headerSubLbl.Invoke((MethodInvoker)(() => headerSubLbl.Text = tickerSubText));
                headerSubLbl.Invoke((MethodInvoker)(() => headerSubLbl.BringToFront()));

            }
            else
            {
                EnableSubHeader();
                headerSubLbl.Text = tickerSubText;
                headerSubLbl.BringToFront();
            }
        }
        public string GetSubHeaderText()
        {
            return headerSubLbl.Text;
        }



        public void SetHeaderTextForeColor(Color headerTextForeColor)
        {
            if (headerTextLbl.InvokeRequired == true)
            {
                headerTextLbl.Invoke((MethodInvoker)(() => headerTextLbl.ForeColor = headerTextForeColor));
            }
            else
            {
                headerTextLbl.ForeColor = headerTextForeColor;
            }
        }
       

        public void SetTickerText(string tickerText)
        {
            if (tickerTextLbl.InvokeRequired == true)
            {
                tickerTextLbl.Invoke((MethodInvoker)(() => tickerTextLbl.Text = tickerText));
            }
            else
            {
                tickerTextLbl.Text = tickerText;
                tickerTextLbl.BringToFront();
            }
        }
        public string GetTickerText()
        {
            return tickerTextLbl.Text;
        }



        public void SetTickerEndImage(Image tickerEndImage)
        {
            if (tickerEndPic.InvokeRequired == true)
            {
                tickerEndPic.Invoke((MethodInvoker)(() => tickerEndPic.Image = tickerEndImage));
            }
            else
            {
                tickerEndPic.Image = tickerEndImage;
            }
        }
        public Image GetTickerEndImage()
        {
            return tickerEndPic.Image;
        }


        public void SetTickerEndColor(Color tickerEndColor)
        {
            if (tickerEndPic.InvokeRequired == true)
            {
                tickerEndPic.Invoke((MethodInvoker)(() => tickerEndPic.BackColor = tickerEndColor));
            }
            else
            {
                tickerEndPic.BackColor = tickerEndColor;
            }
        }
        public Color GetTickerEndColor()
        {
            return tickerEndPic.BackColor;
        }




        public void SetTickerTextForegroundColor(Color tickerForegroundColor)
        {
            if (tickerTextLbl.InvokeRequired == true)
            {
                tickerTextLbl.Invoke((MethodInvoker)(()=>tickerTextLbl.ForeColor = tickerForegroundColor));
            }
        }
        public void SetTickerTextBackgroundColor(Color tickerBackgroundColor)
        {
            if (tickerTextLbl.InvokeRequired == true)
            {
                tickerTextLbl.Invoke((MethodInvoker)(()=>tickerTextLbl.BackColor = tickerBackgroundColor));
                tickerBackDropColorLbl.Invoke((MethodInvoker)(()=>tickerBackDropColorLbl.BackColor = tickerBackgroundColor));
                tickerTextLbl.Invoke((MethodInvoker)(()=>tickerTextLbl.BringToFront()));
            }
            else
            {

                tickerTextLbl.BackColor = tickerBackgroundColor;
                tickerBackDropColorLbl.BackColor = tickerBackgroundColor;
                tickerTextLbl.BringToFront();
            }
            //tickerTextLbl.BackColor = tickerBackgroundColor;
            //this.BackColor = tickerBackgroundColor;
            
        }
        public Color GetTickerTextBackgroundColor()
        {
            return tickerBackDropColorLbl.BackColor;
        }



        public void SetTickerTextBackgroundImage(Image tickerBackgroundImage)
        {
            if (tickerBackgroundPic.InvokeRequired == true)
            {
                tickerBackgroundPic.Invoke((MethodInvoker)(() => tickerBackgroundPic.Image = tickerBackgroundImage));
            }
            else
            {
                tickerBackgroundPic.Image = tickerBackgroundImage;
            }
        }
        public Image GetTickerTextBackgroundImage()
        {
            return tickerBackgroundPic.Image;
        }






        public void SetHeaderBackgroundImage(Image headerBackgroundImage)
        {
            if (headerBackgroundPic.InvokeRequired == true)
            {
                
                headerBackgroundPic.Invoke((MethodInvoker)(()=>headerBackgroundPic.Image = headerBackgroundImage));
                headerBackgroundPic.Invoke((MethodInvoker)(()=>headerBackgroundPic.SizeMode=PictureBoxSizeMode.StretchImage));
                headerBackgroundPic.Invoke((MethodInvoker)(()=>headerBackgroundPic.Refresh()));

            }
            else
            {                
                headerBackgroundPic.Image = headerBackgroundImage;
                headerBackgroundPic.SizeMode = PictureBoxSizeMode.StretchImage;
                headerBackgroundPic.Refresh();
            }
        }
        public Image GetHeaderBackgroundImage()
        {
            return headerBackgroundPic.Image;
        }



        public void SetHeaderBackgroundColor(Color headerBackgroundColor)
        {
            if (headerBackgroundPic.InvokeRequired == true)
            {
                headerBackgroundPic.Invoke((MethodInvoker)(() => headerBackgroundPic.BackColor = headerBackgroundColor));
                headerBackgroundPic.Invoke((MethodInvoker)(() => headerBackgroundPic.Refresh()));
                //tickerTextLbl.Invoke((MethodInvoker)(()=>tickerTextLbl.BackColor=headerBackgroundPic.BackColor));
            }
            else
            {
                headerBackgroundPic.BackColor = headerBackgroundColor;
                headerBackgroundPic.Refresh();
                //tickerTextLbl.BackColor = headerBackgroundPic.BackColor;
            }
        }
        public Color GetHeaderBackgroundColor()
        {
            return headerBackgroundPic.BackColor;
        }



       


        public void EnableSubHeader()
        {
            showSubHeader = true;
            Invoke((MethodInvoker)(() =>headerSubLbl.Visible = true));
        }
        public void DisableSubHeader()
        {
            showSubHeader = false;
            Invoke((MethodInvoker)(() =>headerSubLbl.Visible = false));
        }

        public void EnableTicker()
        {
            if (updateTimerInSeconds == 0)
            { updateTimerInSeconds = 60; }
            //MessageBox.Show("Ticker " + this.Name.ToString() + " has been enabled... next update in " + updateTimerInSeconds.ToString());
       
           isEnabled = true;
           tickerTmr.Enabled = true;
        }
        public void DisableTicker()
        {
            isEnabled = false;
            Invoke((MethodInvoker)(() =>tickerTmr.Enabled = false));
        }

        public void EnableUpdates(int UpdateSpeedInSeconds)
        {
            updateTimerInSeconds = UpdateSpeedInSeconds;
            updateTimerSecondsPassed = 0;
            isUpdateEnabled = true;
            updateTmr.Interval = 1000;
            updateTmr.Enabled = true;
        }
        public void DisableUpdates()
        {
            isUpdateEnabled = false;
            updateTmr.Enabled = false;
        }


        private void tickerTmr_Tick(object sender, EventArgs e)
        {            
            if (isEnabled == false)
            {
                tickerTmr.Enabled = false;
                
            }
           
            tickerTextLbl.SendToBack();
            tickerTextLbl.Left -= 5;
            if (tickerTextLbl.Right < 0)
            {
                tickerTextLbl.Left = this.Width + 200;
            }
        }

        private void tickerTextLbl_Click(object sender, EventArgs e)
        {

        }

        private void tickerBackDropColorLbl_Click(object sender, EventArgs e)
        {

        }

        private void updateTmr_Tick(object sender, EventArgs e)
        {
           updateTimerSecondsPassed++;
            if (updateTimerSecondsPassed >= updateTimerInSeconds)
            {
                //update the ticker using stored method...
                AsyncMethodCaller caller = new AsyncMethodCaller(launchMethodEvent);
                caller.BeginInvoke(null,null);
                //launchMethodEvent();
            }
        }
        private void launchMethodEvent()
        {
            ParameterizedThreadStart pts = new ParameterizedThreadStart(ThreadedProcess);
            Thread processThread = new Thread(pts);
            processThread.Start(UpdateMethod);
            updateTimerSecondsPassed = 0;
        }
        public int GetSecondsPassedDuringCurrentUpdate()
        {
            return updateTimerSecondsPassed;
        }
        public int GetUpdateInterval()
        {
            return updateTimerInSeconds;
        }


    }
}
